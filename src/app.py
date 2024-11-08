import os
import google.auth
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import tkinter as tk
import requests
from googleapiclient.errors import HttpError

def authenticate_sheets():
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    creds = None
    # Verificar se já temos um token salvo
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # Se não houver token, solicitar login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret_958485074740-8ciq0ms7uovp87ng5evki88fvsokcopn.apps.googleusercontent.com.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Salvar o token para reutilização
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    service = build('sheets', 'v4', credentials=creds)
    return service

def fetch_data_from_api(api_url):
    response = requests.get(api_url, verify= False)
    if response.status_code == 200:
        return response.json()  # Retorna os dados da API
    else:
        print("Erro na requisição", response.status_code)
        return None

def update_sheet(service, spreadsheet_id, range_name, values):
    body = {
        'values': values
    }
    try:
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_name,
            valueInputOption="RAW", body=body).execute()
        print(f"{result.get('updatedCells')} células atualizadas.")
    except HttpError as err:
        print(err)

def startprocess():
    service = authenticate_sheets()
    api_url = "https://docs.google.com/spreadsheets/d/1HKr0xjlH5ArMgpcQPJnF3f_RIfvE3EimKEnQ2Zq01c4/edit?gid=2037694645#gid=2037694645"
    data = fetch_data_from_api(api_url)
    if data:
        spreadsheet_id = "ID_DA_PLANILHA"
        range_name = "Página1!A1:D10"
        update_sheet(service, spreadsheet_id, range_name, data)

def create_gui():
    root = tk.Tk()
    root.title("Atualizador de Planilha")

    start_button = tk.Button(root, text="Iniciar Processo", command=startprocess)
    start_button.pack(pady=20)

    reset_button = tk.Button(root, text="Reiniciar Processo", command=startprocess)
    reset_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    create_gui()