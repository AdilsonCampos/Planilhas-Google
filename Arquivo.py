#!/usr/bin/env python
# coding: utf-8

# In[8]:


from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


# ## FOLHA DE PAGAMENTO

# In[9]:


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = ''
SAMPLE_RANGE_NAME = 'Folha de pagamento!A1:I30'


# In[10]:


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        
        values = result['values']
        #print(values)   # cada item é uma linha da tabela
        
        
        funcionarios = []
        nomes = []
        valores = []
        
        for i, dados in enumerate(values):
            if i > 0:
                
                valores.append(float(dados[1].replace("R$", "").replace(".","").replace(",", ".")))
                nomes.append(dados[2])
                if i < 7:
                    funcionarios.append(dados[6])
        
        
        valores_total = []
        
        for funcionario in funcionarios:
            valor = 0
            for i, nome in enumerate(nomes):
                if nome == funcionario:
                    valor += valores[i]
                    
            valores_total.append([valor])
            
        # print(funcionarios, valores_total)
        
        # adicionar/editar uma informação

        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range='Folha de pagamento!H2', valueInputOption="USER_ENTERED",
                                    body={'values': valores_total}).execute()
        
        
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()


# ## MOVIMENTAÇÃO BANCO MENSAL

# In[11]:


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = ''
SAMPLE_RANGE_NAME = 'Banco!A1:H37'


# In[12]:


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        
        values = result['values']
        #print(values)   # cada item é uma linha da tabela
        
        
        colunas_id = []
        categorias = []
        valores = []
        
        for i, dados in enumerate(values):
            
            if i > 0:    
                valores.append(float(dados[1].replace("R$", "").replace(".","").replace(",", ".").replace(" ", "")))
                categorias.append(dados[2])
                if i < 11:
                    colunas_id.append(dados[6])
        
        
        valores_total = []
        
        for coluna in colunas_id:
            valor = 0
            for i, categoria in enumerate(categorias):
                if categoria == coluna:
                    valor += valores[i]
                    
            valores_total.append([valor])

            
            
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range='Banco!H2', valueInputOption="USER_ENTERED",
                                    body={'values': valores_total}).execute()
        
        
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()


# ## MOVIMENTAÇÃO CAIXA MENSAL

# In[13]:


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = ''
SAMPLE_RANGE_NAME = 'Caixa!A1:H37'


# In[14]:


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        
        values = result['values']
        #print(values)   # cada item é uma linha da tabela
        
        
        colunas_id = []
        categorias = []
        valores = []
        
        for i, dados in enumerate(values):
            
            if i > 0:    
                valores.append(float(dados[1].replace("R$", "").replace(".","").replace(",", ".").replace(" ", "")))
                categorias.append(dados[2])
                if i < 11:
                    colunas_id.append(dados[6])
        
        
        valores_total = []
        
        for coluna in colunas_id:
            valor = 0
            for i, categoria in enumerate(categorias):
                if categoria == coluna:
                    valor += valores[i]
                    
            valores_total.append([valor])

            
            
        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range='Caixa!H2', valueInputOption="USER_ENTERED",
                                    body={'values': valores_total}).execute()
        
        
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()


# In[ ]:




