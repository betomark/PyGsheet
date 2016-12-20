# -*- coding: utf-8 -*-
from oauth2client import tools
from webcolors import name_to_rgb

class DriveWriter:
    """Utility class ad hoc for drive spreadsheets interaction
     """
    def __init__(self, spreadsheetId, app_name):
        """Class parameters
        Args:
            spreadsheetId (str): id of spreadsheet (For: https://docs.google.com/spreadsheets/d/<spreadsheetId>/).
            app_name (str): Just a name for class instance.
        """
        self.spreadsheetId = spreadsheetId
        self.app_name = app_name
        self.flags = None
        try:
            import argparse
            self.flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
        except ImportError:
            self.flags = None
        self.service = self.get_service()

    def write_data_in_range(self, data, sheet, sheet_range=None):
        """This function write into a sheet a specified data list
        Args:
            data (:obj:`list` of :obj:`tuple` of :obj:`tuple`): data list which we want to write.
            sheet (str): Sheet name.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for write, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheetId,
            range=self.format_range(sheet, sheet_range), body={'values': data}, valueInputOption='USER_ENTERED').execute()

    def read_data_in_range(self, sheet, sheet_range=None):
        """This function read from a sheet
        Args:
            sheet (str): Sheet name.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for read, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        Returns:
            list: Data stored in sheet's specified range.
        """
        self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheetId,
            range=self.format_range(sheet, sheet_range)).execute()

    def append_data(self, data, sheet):
        """This function appends data to a specified sheet
        Args:
            data (:obj:`list` of :obj:`tuple` of :obj:`tuple`): data list which we want to write.
            sheet (str): Sheet name.
        """
        self.service.spreadsheets().values().append(spreadsheetId=self.spreadsheetId,
            range=sheet, body={'values': data}, valueInputOption='USER_ENTERED').execute()

    def delete_data(self, sheet, sheet_range=None):
        """This function delete a specified range from a sheet
        Args:
            sheet (str): Sheet name.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        self.service.spreadsheets().values().clear(spreadsheetId=self.spreadsheetId,
            range=self.format_range(sheet, sheet_range), body={}).execute()

    def update_borders(self, sheet, style, width, color=name_to_rgb('black'), sheet_range=None, top=True, bottom=True, inner_horizontal=True):
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
        pass

    def create_spreadsheet(self, title):
         """This function creates a new spreadsheet and asign it to current 'spreadsheetless' class for futher work.
        Args:
            title (str): Name of spreadsheet
        """
        pass

    def create_sheet(self, title, index, rows=1000, cols=1000, froz_rows=0, froz_cols=0, hid_grid=False, hidden_sheet=False, tab_color=None):
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
        pass

    def cell_format(self, sheet, background=name_to_rgb('white'), h_alignment=None, v_alignment=None, top_padding=None, right_padding=None, bottom_padding=None, left_padding=None, sheet_range=None):
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
        pass

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
        pass

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


    def protected_range(self, sheet, editors, sheet_range=None):
        """This funtion allow to a specified set of users(using their emails) to edit a certain range.
        Args:
            sheet (str): Sheet name.
            editors (:obj: `list` of :obj: `str`): List of user's emails who you want to conced edit permission.
            sheet_range (:obj: `tuple` of :obj:`tuple` of :obj: `int`, optional): A tuple with two coordinates which delimitate a sheet range for delete it, all sheet by default.
            sheet_range (str, optional): Another implementation of sheet range which suports excel range format, all sheet by default.
        """
        pass

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

    def export_csv(self, sheet):
        """This function allow to export a specified sheet as csv
        Args:
            sheet (str): Sheet name.
        """
        pass


    def get_credentials(self, cred_file='client_secret.json'):
        import os
        from oauth2client.file import Storage
        from oauth2client import client
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
        credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-python-quickstart.json')
        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(cred_file, 'https://www.googleapis.com/auth/spreadsheets')
            flow.user_agent = self.app_name
            if self.flags:
                credentials = tools.run_flow(flow, store, self.flags)
            else:  # Needed only for compatibility with Python 2.6
                credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def get_service(self):
        import httplib2
        from apiclient import discovery
        credentials = self.get_credentials()
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

        formated = sheet
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



class TextFormat:
    def __init__(self, color=name_to_rgb('black'), sheet_range=None, font='Comic Sans MS', size=None, bold=False, italic=False):
        pass

class CellFormat:
    def __init__(self, background=name_to_rgb('white'), h_alignment=None, v_alignment=None, top_padding=None, right_padding=None, bottom_padding=None, left_padding=None):
        pass

class GradientFormat:
    def __init__(self, init, mid, end, init_col, mid_col, end_col, interpolation_type):
        pass
