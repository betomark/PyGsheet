# -*- coding: utf-8 -*-
from oauth2client import tools
from webcolors import name_to_rgb

class SpreadsheetManager:
    """Utility class ad hoc for drive spreadsheets interaction
     """
    def __init__(self, app_name, spreadsheetId=None, with_pipeline=False, cred_path=None):
        """Class parameters
        Args:
            spreadsheetId (str): id of spreadsheet (For: https://docs.google.com/spreadsheets/d/<spreadsheetId>/).
            app_name (str): Just a name for class instance.
        """
        self.app_name = app_name
        self.cred_path = None
        if cred_path:
            self.service = self.get_service(cred_path)
            self.cred_path = cred_path
        else:
            self.service = self.get_service()

        self.spreadsheetId = None
        if spreadsheetId:
            self.spreadsheetId = spreadsheetId
        else:
            self.create_spreadsheet(app_name)

        self.sheets_id = self.get_sheets_id()

        self.app_name = app_name
        self.flags = None
        try:
            import argparse
            self.flags = argparse.ArgumentParser(parents=[tools.argparser], conflict_handler='resolve').parse_args()
        except ImportError:
            self.flags = None

        self.with_pipeline = with_pipeline
        if self.with_pipeline:
            self.pipeline = []
            
        self.drive_manager = None

    def write_data_in_range(self, data, sheet, sheet_range=None, value_input='USER_ENTERED'):
        """This function write into a sheet a specified data list
        Args:
            data (:obj:`list` of :obj:`tuple` of :obj:`tuple`): data list which we want to write.
            sheet (str): Sheet name.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for write, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        range_formated = self.format_range(sheet, sheet_range)
        if self.with_pipeline:

            self.pipeline.append(None)  # TODO: implementation
        else:
            self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheetId,
                range=range_formated, body={'values': data}, valueInputOption=value_input).execute()

    def read_data_in_range(self, sheet, sheet_range=None, omit_empty=False):
        """This function read from a sheet
        Args:
            sheet (str): Sheet name.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for read, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
            omit_empty (bool, optional): True if you want to retrieve empty cells as None.
        Returns:
            list: Data stored in sheet's specified range.
        """
        response = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheetId,
            range=self.format_range(sheet, sheet_range), valueRenderOption='UNFORMATTED_VALUE').execute()
        if 'values' in response and not omit_empty:
            return [val[0] if len(val) == 1 else {0: None}.get(len(val), val) for val in response['values']]
        elif 'values' in response and omit_empty:
            return [val[0] if len(val) == 1 else val for val in response['values'] if len(val) > 0]
        else:
            return []

    def append_data(self, data, sheet, value_input='USER_ENTERED'):
        """This function appends data to a specified sheet
        Args:
            data (:obj:`list` of :obj:`tuple` of :obj:`tuple`): data list which we want to write.
            sheet (str): Sheet name.
        """
#        if self.with_pipeline:
#            data_list = []
#            for line in data:
#                line_dict = {"values": []}
#                for element in line:
#
#                    cell_data = {  # TODO: finish it
#                        "userEnteredValue": {object(ExtendedValue)},
#                        "effectiveValue": {object(ExtendedValue)},
#                        "formattedValue": string,
#                        "userEnteredFormat": {object(CellFormat)},
#                        "effectiveFormat": {object(CellFormat)},
#                        "hyperlink": string,
#                        "note": string,
#                        "textFormatRuns": [{object(TextFormatRun)}],
#                        "dataValidation": {object(DataValidationRule)},
#                        "pivotTable": {object(PivotTable)},
#                    }
#
#                    line_dict["values"].append(cell_data)
#                data_list.append(line_dict)
#
#            data = { "appendCells": {
#                "sheetId": self.spreadsheetId,
#                "rows": data_list,
#                "fields": "*",
#                }
#            }
#            self.pipeline.append(data)
#        else:
        self.service.spreadsheets().values().append(spreadsheetId=self.spreadsheetId,
                range=sheet, body={'values': data}, valueInputOption=value_input).execute()

    def delete_data(self, sheet, sheet_range=None):
        """This function delete a specified range from a sheet
        Args:
            sheet (str): Sheet name.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        self.service.spreadsheets().values().clear(spreadsheetId=self.spreadsheetId,
            range=self.format_range(sheet, sheet_range), body={}).execute()

    def update_borders(self, sheet, style, width, color=name_to_rgb('black'), alpha=None, sheet_range=None, top=True, bottom=True, left=True, right=True, inner_horizontal=True, inner_vertical=True):
        """This function changes borders style of a specified range
        Args:
            sheet (str): Sheet name.
            style(str): A choice from the following options ().
            width(int): Width of the border.
            color(:obj: `tuple` of :obj: 'float', optional): A tuple which have the RGB code (Black for instance: (0.0, 0.0, 0.0})
            color(str, optional): color name according web nomenclature.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
            top(bool, optional): True in case you want to modify a cell's top, otherwise False. True by default.
            bottom(bool, optional): True in case you want to modify a cell's bottom, otherwise False. True by default.
            inner_horizontal(bool, optional): True in case you want to modify a cell's inner_horizontal, otherwise False. True by default.
        """
        sheet_id = self.get_sheets_id(sheet)
        range_points = sheet_range if type(sheet_range, tuple) else self.get_range_points(sheet_range)
        data = {
            "updateBorders": {
                "range": {
                    "sheetId": sheet_id,
                }
            }
        }

        if range_points[0][0]:
            data["updateBorders"]["range"]["startColumnIndex"] = range_points[0][0]

        if range_points[0][1]:
            data["updateBorders"]["range"]["startRowIndex"] = range_points[0][0] + 1

        if range_points[1][0]:
            data["updateBorders"]["range"]["endColumnIndex"] = range_points[1][0]

        if range_points[1][1]:
            data["updateBorders"]["range"]["endRowIndex"] = range_points[1][0] + 1

        if top:
            data["updateBorders"]["top"] = {
                    "color": {
                        "red": color[0],
                        "green": color[1],
                        "blue": color[2]
                    },
                    "width": width,
                    "style": style
                }

            if alpha:
                 data["updateBorders"]["top"]["color"]["alpha"] = alpha

        if bottom:
            data["updateBorders"]["bottom"] = {
                    "color": {
                        "red": color[0],
                        "green": color[1],
                        "blue": color[2]
                    },
                    "width": width,
                    "style": style
                }
            if alpha:
                 data["updateBorders"]["bottom"]["color"]["alpha"] = alpha
        if left:
            data["updateBorders"]["left"] = {
                    "color": {
                        "red": color[0],
                        "green": color[1],
                        "blue": color[2]
                    },
                    "width": width,
                    "style": style
                }
            if alpha:
                 data["updateBorders"]["left"]["color"]["alpha"] = alpha

        if right:
            data["updateBorders"]["right"] = {
                    "color": {
                        "red": color[0],
                        "green": color[1],
                        "blue": color[2]
                    },
                    "width": width,
                    "style": style
                }
            if alpha:
                 data["updateBorders"]["right"]["color"]["alpha"] = alpha


        if inner_horizontal:
            data["updateBorders"]["innerHorizontal"] = {
                    "color": {
                        "red": color[0],
                        "green": color[1],
                        "blue": color[2]
                    },
                    "width": width,
                    "style": style
                }
            if alpha:
                 data["updateBorders"]["innerHorizontal"]["color"]["alpha"] = alpha

        if inner_vertical:
            data["updateBorders"]["innerVertical"] = {
                    "color": {
                        "red": color[0],
                        "green": color[1],
                        "blue": color[2]
                    },
                    "width": width,
                    "style": style
                }
            if alpha:
                 data["updateBorders"]["innerVertical"]["color"]["alpha"] = alpha

        if self.with_pipeline:
            self.pipeline.append(data)
        else:
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(data)).execute()

    def create_spreadsheet(self, title):
        """This function creates a new spreadsheet and asign it to current 'spreadsheetless' class for futher work.
        Args:
            title (str): Name of spreadsheet
        """
        data = {
            "properties": {
                "title": title,
                "locale": 'es',
                "timeZone": 'GMT+01:00'
                }
        }
        request = self.service.spreadsheets().create(body=data).execute()
        self.spreadsheetId = request["spreadsheetId"]
        return self.spreadsheetId

    def create_sheet(self, title, index=None, rows=1000, cols=1000, froz_rows=0, froz_cols=0, hid_grid=False, hidden_sheet=False, tab_color=None):
        """This function appends a sheet to a spreadsheet.
        Args:
            title (str): Sheet name.
            index (int): Sheet tab position.
            rows (int, optional): Number of rows of new sheet.
            cols (int, optional): Number of columns of new sheet.
            froz_rows (int, optional): Number of frozen rows.
            froz_cols (int, optional): Number of frozen columns.
            hid_grid (bool, optional): True if you want to not show grid borders.
            hidden_sheet (bool, optional): True if you want to create a hidden sheet.
            tab_color (:obj: `tuple` of :obj: 'float', optional): Change background color of sheet's tab.
        """
        data = {
            "addSheet": {
                "properties": {
                    "title": title,
                    "gridProperties": {
                        "columnCount": cols,
                        "rowCount": rows,
                        "frozenColumnCount": froz_cols,
                        "frozenRowCount": froz_rows,
                        "hideGridlines": hid_grid
                    },
                    "hidden": hidden_sheet
                }
            }
        }

        if index:
            data["addSheet"]["properties"]["index"] = index

        if tab_color:
            data["addSheet"]["properties"]["tabColor"] = {
                "red": tab_color[0],
                "green": tab_color[1],
                "blue": tab_color[2],
                "alpha": 0.9
            }

        sheet_id = sorted([sheet["properties"]["sheetId"] for sheet in self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()["sheets"]])[-1] + 1
        data["addSheet"]["properties"]["sheetId"] = sheet_id
        if self.with_pipeline:
            self.pipeline.append(data)
        else:
            response = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(data)).execute()
            if response:  # TODO: test it
                self.sheets_id[title] = sheet_id

    def delete_sheet(self, sheet):
        data = {
            "deleteSheet": {
                "sheetId": self.get_sheets_id(sheet)
            }
        }
        if self.with_pipeline:
            self.pipeline.append(("batchUpdate", data))
        else:
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(data)).execute()
        
    def cell_format(self, sheet, number_format=None, background=name_to_rgb('white'), h_alignment=None, v_alignment=None, top_padding=None, right_padding=None, bottom_padding=None, left_padding=None, sheet_range=None):
        """This function changes cell's format
        Args:
            sheet (str): Sheet name.
            background (:obj: `tuple` of :obj: 'float', optional): background color for cell range, white by default.
            h_alignment (str, optional): Horizontal alignment. A chocie between ['left', 'center', 'right'].
            v_alignment (str, optional): Vertical alignment. A chocie between ['top', 'center', 'bottom'].
            top_padding (int, optional): cell's top padding.
            right_padding (int, optional): cell's right padding.
            bottom_padding (int, optional): cell's bottom padding.
            left_padding (int, optional): cell's left padding.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        sheet_id = self.get_sheets_id(sheet)
        range_points = sheet_range if isinstance(sheet_range, tuple) else self.get_range_points(sheet_range)

        data = {
            "UpdateCellsRequest": {
                "rows": [],
                "range": {"sheetId": sheet_id},
            }
        }

        if range_points[0][0]:
            data["UpdateCellsRequest"]["range"]["startColumnIndex"] = range_points[0][0]

        if range_points[0][1]:
            data["UpdateCellsRequest"]["range"]["startRowIndex"] = range_points[0][0] + 1

        if range_points[1][0]:
            data["UpdateCellsRequest"]["range"]["endColumnIndex"] = range_points[1][0]

        if range_points[1][1]:
            data["UpdateCellsRequest"]["range"]["endRowIndex"] = range_points[1][0] + 1

        n_rows = range_points[1][1] - range_points[0][1] if range_points[1][1] and range_points[0][1] else 1
        n_cols = range_points[1][0] - range_points[0][0] if range_points[1][0] and range_points[0][0] else 1

        for i in range(n_rows):
            row_data = {"values": []}

            for j in range(n_cols):
                cell_data = {
                    "userEnteredFormat": {}
                }
                if number_format:
                    cell_data["userEnteredFormat"]["numberFormat"]["type"] = number_format

                if background:
                    cell_data["userEnteredFormat"]["backgroundColor"] = {
                        "red": background[0],
                        "green": background[1],
                        "blue": background[2]
                    }
                if top_padding:
                    cell_data["userEnteredFormat"]["padding"]["top"] = top_padding
                if right_padding:
                    cell_data["userEnteredFormat"]["padding"]["right"] = right_padding
                if bottom_padding:
                    cell_data["userEnteredFormat"]["padding"]["bottom"] = bottom_padding
                if left_padding:
                    cell_data["userEnteredFormat"]["padding"]["left"] = left_padding
                if h_alignment:
                    cell_data["userEnteredFormat"]["padding"]["horizontalAlignment"] = h_alignment
                if v_alignment:
                    cell_data["userEnteredFormat"]["padding"]["verticalAlignment"] = v_alignment

                row_data["values"].append(cell_data)

            data["UpdateCellsRequest"]["rows"].append(row_data)

        if self.with_pipeline:
            self.pipeline.append(("batchUpdate", data))
        else:
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(data)).execute()


    def text_format(self, sheet, color=name_to_rgb('black'), sheet_range=None, font='Comic Sans MS', size=None, bold=False, italic=False):
        """This function changes text's format
        Args:
            sheet (str): Sheet name.
            color(:obj: `tuple` of :obj: 'float', optional): A tuple which have the RGB code (Black for instance: (0.0, 0.0, 0.0})
            color(str, optional): color name according web nomenclature.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
            font (str, optional): Name of text font. Comic Sans MS as default.
            size (int, optional): A number which represents text's size.
            bold (bool, optional): True if you want to bold it. False as default.
            italic (bool, optional): True if you want to italic it. False as default.
        """
        sheet_id = self.get_sheets_id(sheet)
        range_points = sheet_range if type(sheet_range, tuple) else self.get_range_points(sheet_range)

        data = {
            "UpdateCellsRequest": {
                "rows": [],
                "range": {"sheetId": sheet_id},
            }
        }

        if range_points[0][0]:
            data["UpdateCellsRequest"]["range"]["startColumnIndex"] = range_points[0][0]

        if range_points[0][1]:
            data["UpdateCellsRequest"]["range"]["startRowIndex"] = range_points[0][0] + 1

        if range_points[1][0]:
            data["UpdateCellsRequest"]["range"]["endColumnIndex"] = range_points[1][0]

        if range_points[1][1]:
            data["UpdateCellsRequest"]["range"]["endRowIndex"] = range_points[1][0] + 1

        n_rows = range_points[1][1] - range_points[0][1] if range_points[1][1] and range_points[0][1] else 1
        n_cols = range_points[1][0] - range_points[0][0] if range_points[1][0] and range_points[0][0] else 1

        for i in range(n_rows):
            row_data = {"values": []}

            for j in range(n_cols):
                cell_data = {
                    "userEnteredFormat": {"textFormat": {}}
                }
                if color:
                    cell_data["userEnteredFormat"]["textFormat"]["foregroundColor"] = {
                        "red": color[0],
                        "green": color[1],
                        "blue": color[2]
                    }
                if font:
                     cell_data["userEnteredFormat"]["textFormat"]["fontFamily"] = font
                if size:
                    cell_data["userEnteredFormat"]["textFormat"]["fontSize"] = size
                if bold:
                    cell_data["userEnteredFormat"]["textFormat"]["bold"] = bold
                if italic:
                    cell_data["userEnteredFormat"]["textFormat"]["italic"] = italic


                row_data["values"].append(cell_data)

            data["UpdateCellsRequest"]["rows"].append(row_data)

        if self.with_pipeline:
            self.pipeline.append(("batchUpdate", data))
        else:
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(data)).execute()

    def filter_view(self, sheet, id, title, condition_type, condition, sort_order=None, sheet_range=None):
        """
        This function creates a filter view in a specified sheet
        Args:
            sheet (str): Sheet name.
            id (number): An id for this filter.
            title (str): A name for this filter.
            condition_type (str): Which kind of filter you want. Here a list with supported filter kinds: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#conditiontype
            condition (str): Condition predicate which must have excel formula format.
            sort_order (srt, optional): Sorting criteria. A choice between ['ascending', 'descending']. Chaos as default.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        pass


#    def protected_range(self, sheet, editors, protected_range=None, sheet_range=None):
#        """This funtion allow to a specified set of users(using their emails) to edit a certain range.
#        Args:
#            sheet (str): Sheet name.
#            editors (:obj: `list` of :obj: `str`): List of user's emails who you want to conced edit permission.
#            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
#            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
#        """
#        request = {  # TODO: finish it
#            "addProtectedRange": {object(AddProtectedRangeRequest)
#            },
#            "updateProtectedRange": {object(UpdateProtectedRangeRequest)
#            },
#            "deleteProtectedRange": {object(DeleteProtectedRangeRequest)
#            }
#        }
#        if self.with_pipeline:
#            self.pipeline.append(request)
#        else:
#            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(request)).execute()


    def conditional_format(self, sheet, condition_type, condition, text_format_class=None, cell_format_class=None, gradient_format_class=None, sheet_range=None):
        """This funtion set conditional format rule for a specified range in a sheet.
        Args:
            sheet (str): Sheet name.
            condition_type (str): Which kind of filter you want. Here a list with supported filter kinds: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#conditiontype
            condition (str): Condition predicate which must have excel formula format.
            text_format_class (TextFormat, optional): Gets a text format class for use it in conditional formating.
            cell_format_class (CellFormat, optional): Gets a cell format class for use it in conditional formating.
            gradient_format_class (GradientFormat, optional): Gets a gradient format class for use it in conditional formating.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        pass

#    def find_and_replace(self, find, replace, sheet_range=None, sheet=None,  case_sensitive=True, entire_cell=True, regrex=False, search_formula=False):
#        request = {
#            "findReplace": {
#                "find": find,
#                "replacement": replace,
#                "matchCase": case_sensitive,
#                "matchEntireCell": entire_cell,
#                "searchByRegex": regrex,
#                "includeFormulas": search_formula,
#                "range": {  # TODO: finish it
#                  "sheetId": number,
#                  "startRowIndex": number,
#                  "endRowIndex": number,
#                  "startColumnIndex": number,
#                  "endColumnIndex": number,
#                }
#            }
#        }
#        if sheet:
#            request['findReplace']["sheetId"] = self.sheets_id[sheet]
#        else:
#            request['findReplace']["allSheets"] = True
#
#        if self.with_pipeline:
#            self.pipeline.append(request)
#        else:
#            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(request)).execute()

    def get_sheets_id(self, sheet=None):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        sheets_dic = {sheet["properties"]["title"]: sheet["properties"]["sheetId"] for sheet in spreadsheet["sheets"]}
        if sheet:
            return sheets_dic[sheet]
        else:
            return sheets_dic

#    def export_csv(self, sheet=None, file_name=None):
#        """This function allow to export a specified sheet as csv
#        Args:
#            sheet (str): Sheet name.
#        """
#        if file_name:
#            file = open(file_name + '.csv', 'w')
#        else:
#            from datetime import datetime
#            file = open('pygsheet_export_at_' + str(datetime.now().isoformat))
#
#        if sheet:
#            spreadsheet = [element for element in self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId)["sheets"] if sheet["properties"]["title"] == sheet][0]
#
#        else:


    def share_spreadsheet(self, domain=None, user_list=None):
        """This function allow to share the spreadsheet to users or a domain
        Args:
            domain (str): name of the domain which you want to share the spreadsheet, commenter permission only.
            user_list (list): an email list for full spreadsheet sharing.
        """
        if not self.drive_manager:
            self.get_drive_manager()
            
        if user_list:
            if domain:
                self.drive_manager.share_file(self.spreadsheetId, domain, user_list)
            else:
                self.drive_manager.share_file(self.spreadsheetId, user_list=user_list)
        else:
            if domain:
                self.drive_manager.share_file(self.spreadsheetId, domain=domain)
            else:
                raise AttributeError('You must specify a domain and/or email list')
                
    def move_file_to_folder(self, folder_id, remove_parents=False):
        if not self.drive_manager:
            self.get_drive_manager()
        if remove_parents:
            self.drive_manager.move_file_to_folder(self.spreadsheetId, folder_id, remove_parents=True)
        else:
            self.drive_manager.move_file_to_folder(self.spreadsheetId, folder_id)
        
                
    def get_drive_manager(self):
        from pygsheet import drive_manager
        self.drive_manager = drive_manager.DriveManager(app_name=self.app_name, cred_path=self.cred_path)

    def get_credentials(self, cred_path, cred_file='client_secret.json'):
        import os
        from oauth2client.file import Storage
        from oauth2client import client
        if not cred_path:
            home_dir = os.path.expanduser('~')
            credential_dir = os.path.join(home_dir, '.credentials')
            if not os.path.exists(credential_dir):
                try: 
                    os.system("sudo mkdir {}".format(credential_dir))
                except:
                    os.umask(0)
                    os.makedirs(credential_dir, mode=0o777)
        else:
            credential_dir = cred_path
        
        credential_path = os.path.join(credential_dir, 'python-quickstart.json')
        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            if cred_path is None:
                flow = client.flow_from_clientsecrets(cred_file, scope=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
            else:
                flow = client.flow_from_clientsecrets(cred_path + cred_file, scope=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
            flow.user_agent = self.app_name
            try:
                credentials = tools.run_flow(flow, store, self.flags)
            except:  # Needed only for compatibility with Python 2.6
                try:
                    credentials = tools.run_flow(flow, store)
                except:
                    credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def get_service(self, cred_path=None):
        import httplib2
        from apiclient import discovery
        credentials = self.get_credentials(cred_path=cred_path)
        http = credentials.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        return discovery.build('sheets', 'v4', http=http, discoveryServiceUrl=discoveryUrl)

    def format_range(self, sheet, sheet_range):
        def excel_style(row, col):
            LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
            result = []
            while col:
                col, rem = divmod(col-1, 26)
                result[:0] = LETTERS[rem]
            return ''.join(result) + str(row)

        formated = "'" + sheet + "'"
        if sheet_range:
            if isinstance(sheet_range, list):
                pass
            elif isinstance(sheet_range, tuple):
                init = excel_style(sheet_range[0][0], sheet_range[0][1])
                end = excel_style(sheet_range[1][0], sheet_range[1][1])
                formated += '!' + init + ':' + end
            elif isinstance(sheet_range, str):
                formated += '!' + sheet_range
        return formated
    
    def get_range_points(self, range):
        pass  # TODO: finish it

    def create_request_body(self, data):
        if isinstance(data, list):
            body = {
                "requests": data,
                "includeSpreadsheetInResponse": False,
                "responseIncludeGridData": False,
            }
        else:
            body = {
                "requests": [data],
                "includeSpreadsheetInResponse": False,
                "responseIncludeGridData": False,
            }
        return body

#    def execute_pipeline(self):
#        response = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=self.create_request_body(self.pipeline)).execute()
#        if response and str(self.pipeline)



class TextFormat:
    def __init__(self, color=name_to_rgb('black'), sheet_range=None, font='Comic Sans MS', size=None, bold=False, italic=False):
        pass

class CellFormat:
    def __init__(self, background=name_to_rgb('white'), h_alignment=None, v_alignment=None, top_padding=None, right_padding=None, bottom_padding=None, left_padding=None):
        pass

class GradientFormat:
    def __init__(self, init, mid, end, init_col, mid_col, end_col, interpolation_type):
        pass
