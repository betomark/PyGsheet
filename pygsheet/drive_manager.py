# -*- coding: utf-8 -*-
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from apiclient.http import MediaFileUpload
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

class DriveManager():
    def __init__(self, secret_file='client_secret.json', app_name='Drive API Python Quickstart', cred_path=None):
        self.secret_file = secret_file
        self.app_name = app_name
        self.service = self.get_service(cred_path=cred_path)


    def get_credentials(self, cred_path):
        """Gets valid user credentials from storage.

        If nothing has been stored, or if the stored credentials are invalid,
        the OAuth2 flow is completed to obtain the new credentials.

        Returns:
            Credentials, the obtained credential.
        """
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
        credential_path = os.path.join(credential_dir,
                                       'python-quickstart.json')

        store = Storage(credential_path)
        credentials = store.get()
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets(self.secret_file, scope=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive'])
            flow.user_agent = self.app_name
            try:
                import argparse
                flags = argparse.ArgumentParser(parents=[tools.argparser], conflict_handler='resolve').parse_args()
            except ImportError:
                flags = None
            try:
                credentials = tools.run_flow(flow, store, flags)
            except: # Needed only for compatibility with Python 2.6
                try:
                    credentials = tools.run_flow(flow, store)
                except:
                    credentials = tools.run(flow, store)
            print('Storing credentials to ' + credential_path)
        return credentials

    def get_service(self, cred_path=None):
        credentials = self.get_credentials(cred_path=cred_path)
        http = credentials.authorize(httplib2.Http())
        return discovery.build('drive', 'v3', http=http)

    def share_file(self, file_id, domain=None, user_list=None):
        def callback(request_id, response, exception):
            if exception:
                # Handle error
                print(exception)
            else:
                print("Permission Id: %s" % response.get('id'))

        batch = self.service.new_batch_http_request(callback=callback)

        if user_list:
            for email in user_list:
                batch.add(self.service.permissions().create(
                    fileId=file_id,
                    body= {'type': 'user', 'role': 'writer', 'emailAddress': email},
                    fields='id',
                ))
            if domain:
                batch.add(self.service.permissions().create(
                    fileId=file_id,
                    body={'type': 'domain', 'role': 'commenter', 'domain': domain},
                    fields='id',
                ))

        batch.execute()

    def update_sharing(self):
        pass

    def create_folder(self, name, parents=None, team_drives=True):
        file_metadata = {
            'name': name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        if parents:
            if not isinstance(parents, list):
                parents = [parents]
            file_metadata['parents'] = parents
        return self.service.files().create(body=file_metadata, supportsTeamDrives=team_drives).execute()

    def move_file_to_folder(self, file_id, folder_id, remove_parents=False, team_drives=True):

        if remove_parents:
            file = self.service.files().get(fileId=file_id,
                                         fields='parents').execute()
            previous_parents = ",".join(file.get('parents'))
            file = self.service.files().update(fileId=file_id,
                                            addParents=folder_id,
                                            removeParents=previous_parents,
                                            fields='id, parents',
                                            supportsTeamDrives=team_drives).execute()
        else:
            self.service.files().update(fileId=file_id,
                                            addParents=folder_id,
                                            fields='id, parents',
                                            supportsTeamDrives=team_drives).execute()

    def copy_file(self, file_id, new_name=None, new_folder=None, body=None, team_drives=True):
        response = None
        if new_name:
            if body:
                body['name'] = new_name
            else:
                body = {'name': new_name}
            response = self.service.files().copy(fileId=file_id, body=body, supportsTeamDrives=team_drives).execute()
        else:
            if body:
                response = self.service.files().copy(fileId=file_id, body=body, supportsTeamDrives=team_drives).execute()
            else:
                response = self.service.files().copy(fileId=file_id, supportsTeamDrives=team_drives).execute()
        if new_folder:
            self.move_file_to_folder(response['id'], new_folder, remove_parents=True)
        return response

    def upload_file(self, filename, mtype, folder=None, team_drives=True):
        mimes= {
                'ppt': 'application/vnd.mspowerpoint',
                'pdf': 'application/pdf',
                'gif': 'image/gif',
                'png': 'image/png',
                'jpeg': 'image/jpeg',
                'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                'xls': 'application/vnd.ms-excel',
                'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            }

        file_metadata = {
            'title': filename,
            'mimeType': mimes[mtype],
            'name': filename
        }
        if folder:
            file_metadata['parents'] = [{'id': folder}]

        file = MediaFileUpload(filename, mimetype=mimes[mtype])
        response = self.service.files().create(body=file_metadata,
                                     media_body=file,
                                     fields='id',
                                     supportsTeamDrives=team_drives).execute()
        return response

    def list_files(self, team_drive_id = None, team_drives=True, ):
        files = []
        corp ='teamDrive' if team_drives else 'user'
        next_page = True
        token = None
#        while next_page:
        if team_drives:
            if token:
                response = self.service.files().list(corpora= corp, supportsTeamDrives=team_drives, includeTeamDriveItems=team_drives, pageSize=1000, teamDriveId=team_drive_id).execute()
            else:
                response = self.service.files().list(corpora= corp, supportsTeamDrives=team_drives, includeTeamDriveItems=team_drives, pageSize=1000, teamDriveId=team_drive_id, pageToken=token).execute()
            for file in response['files']:
                files.append(file['id'])
            #if response["nextPageToken"]:
    #                    token = response["nextPageToken"]
    #                    continue
    #                else:
#                break
        else:
            if token:
                response = self.service.files().list(corpora= corp, supportsTeamDrives=team_drives, includeTeamDriveItems=team_drives, pageSize=1000).execute()
            else:
                response = self.service.files().list(corpora= corp, supportsTeamDrives=team_drives, pageSize=1000, pageToken=token).execute()
            for file in response['files']:
                files.append(file['id'])
#            if response["nextPageToken"]:
#                token = response["nextPageToken"]
#                continue
#            else:
#                break
        return files

    def list_team_drives(self,):
        team_drives = []
        response = self.service.teamdrives().list(pageSize=100).execute()
        for team_drive in response["teamDrives"]:
            team_drives.append({'id':team_drive['id'], 'name': team_drive['name']})
        return team_drives


