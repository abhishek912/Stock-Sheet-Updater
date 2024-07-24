from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import gspread
from openpyxl import load_workbook
import pickle
import os

class GoogleSheet:
    __TOKEN_PATH = 'token.pickle'
    __CLIENT_SECRET_FILE = 'credentials.json'
    __SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets.readonly',
    'https://www.googleapis.com/auth/spreadsheets', 
    'https://www.googleapis.com/auth/drive']

    @classmethod
    def getAuthorizationCredentials(cls):
        try :
            credentials = None
            # Check if the token.pickle file already exists
            if os.path.exists(cls.__TOKEN_PATH):
                with open(cls.__TOKEN_PATH, 'rb') as token:
                    credentials = pickle.load(token)
            # If there are no valid credentials available, let the user log in.
            if not credentials or not credentials.valid:
                if credentials and credentials.expired and credentials.refresh_token:
                    credentials.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                        cls.__CLIENT_SECRET_FILE, cls.__SCOPES)
                    credentials = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open(cls.__TOKEN_PATH, 'wb') as token:
                    pickle.dump(credentials, token)
                    
            print("Credentials obtained successfully")
            return credentials
        except Exception as e:
            print(f"Error getAuthorizationCredentials: {e}")
            return None
        
    @classmethod
    def getAuthorizedClient(cls):
        credentials = cls.getAuthorizationCredentials()
        try :
            if(credentials != None):
                return gspread.authorize(credentials)
        except Exception as e:
            print(f"Error getAuthorizedClient: {e}")
            return None
        
    def getSheetHandle(self, sheet_name):
        try :
            client = GoogleSheet.getAuthorizedClient()
            if(client != None):
                print('Received the client')
                return client.open(sheet_name)
        except Exception as e:
            print(f"Error getAuthorizedClient: {e}")
            return None
    
    def getWorksheetHandleByWorksheetName(self, sheet_handle, worksheet_name):
        try :
            return sheet_handle.worksheet(worksheet_name)
        except Exception as e:
            print(f"Error getSpecificWorksheetByName: {e}")
            return None
    
    def getWorksheetData(self, work_sheet_handle):
        try :
            return work_sheet_handle.get_all_values()
        except Exception as e:
            print(f"Error getAllDataFromWorksheet: {e}")
            return []

    def getSpecificWorksheetData(self, sheet_name, worksheet_name):
        try :
            sheetHandle = self.getSheetHandle(sheet_name)
            if sheetHandle == None:
                return None
            
            print('Received an open sheet handle')
            worksheetHandle = self.getWorksheetHandleByWorksheetName(sheetHandle, worksheet_name)
            if worksheetHandle == None:
                return None
            
            print('Received an open worksheet handle')
            return self.getWorksheetData(worksheetHandle)
        except Exception as e:
            print(f"Error getAllDataFromWorksheet: {e}")
            return []
        
    def getDefaultWorksheetData(self, sheet_name):
        try :
            sheetHandle = self.getSheetHandle(sheet_name)
            if sheetHandle == None:
                return None
            
            print('Received an open sheet handle')
            worksheetHandle = sheetHandle.get_worksheet(0)
            if worksheetHandle == None:
                return None
            
            print('Received an open worksheet handle')
            return self.getWorksheetData(worksheetHandle)
        except Exception as e:
            print(f"Error getAllDataFromWorksheet: {e}")
            return []
    
    def updateCell(self, work_sheet_handle, row, cell, value):
        try :
            work_sheet_handle.update_cell(row, cell, value)
        except Exception as e:
            print(f"Error updateCell: {e}")
            return None
    
    