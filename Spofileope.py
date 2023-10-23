import requests
import json
import openpyxl
import json

def load_config(filename):
    with open(filename, 'r') as file:
        config = json.load(file)
    return config

def get_sharepoint_folders():

    config = load_config("config.json")

    # Azure ADで登録したアプリケーションの情報
    client_id = config["client_id"]
    client_secret = config["client_secret"]
    tenant_id = config["tenant_id"]
    resource = config["sharepoint_url"]

    # OAuth 2.0のトークンを取得
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/token"
    token_data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'resource': resource
    }
    token_r = requests.post(token_url, data=token_data)
    token = token_r.json().get('access_token')

    # SharePointのフォルダ一覧を取得
    folder_url = f"{resource}/_api/web/GetFolderByServerRelativeUrl('/sites/YourSiteName/YourLibraryName')/folders"
    headers = {
        'Authorization': f'Bearer {token}',
        'Accept': 'application/json;odata=verbose'
    }

    response = requests.get(folder_url, headers=headers)

    if response.status_code == 200:
        data = json.loads(response.text)
        folders = data['d']['results']
        for folder in folders:
            folder_name = folder['Name']
            print(f"Folder: {folder_name}")
            get_excel_files_in_folder(folder_name, headers)
    else:
        print(f"Failed to retrieve folder list. Status code: {response.status_code}")

def get_excel_files_in_folder(folder_name, headers):
    excel_files_url = f"{resource}/_api/web/GetFolderByServerRelativeUrl('/sites/YourSiteName/YourLibraryName/{folder_name}')/Files"
    excel_files_response = requests.get(excel_files_url, headers=headers)

    if excel_files_response.status_code == 200:
        excel_data = json.loads(excel_files_response.text)
        excel_files = excel_data['d']['results']
        for excel_file in excel_files:
            if excel_file['Name'].endswith('.xlsx'):
                print(f"Excel File: {excel_file['Name']}")
                download_and_open_excel(excel_file['Name'], headers)
    else:
        print(f"Failed to retrieve Excel files. Status code: {excel_files_response.status_code}")

def download_and_open_excel(file_name, headers):
    excel_file_url = f"{resource}/_api/web/GetFileByServerRelativeUrl('/sites/YourSiteName/YourLibraryName/{file_name}')/$value"
    excel_file_response = requests.get(excel_file_url, headers=headers)

    if excel_file_response.status_code == 200:
        with open(file_name, 'wb') as excel_file:
            excel_file.write(excel_file_response.content)
        
        open_excel_file(file_name)
    else:
        print(f"Failed to retrieve Excel file. Status code: {excel_file_response.status_code}")

def open_excel_file(file_name):
    # Excelファイルを開く
    wb = openpyxl.load_workbook(file_name)
    sheet = wb.active
    for row in sheet.iter_rows(values_only=True):
        for cell in row:
            print(cell, end='\t')
        print()

def save_excel_file(wb, file_name):
    # Excelファイルを保存
    wb.save(file_name)

def main():


if __name__ == "__main__":
    pass
