import pandas as pd
import os
import re
import requests
from requests_toolbelt.multipart.encoder import MultipartEncoder
import time


API_KEY = "sk_test_1722862339_uYR1pgdGoNosJoVJQRSWbsmShWYQMPtC"  # Place your API Key here
EXCEL_FILE = '/Users/Downloads/Book1.xlsx'  # Place your input file path here
OUTPUT_FILE = '/Users/Downloads/Book2.xlsx'  # Place your output file path here


def read_excel(file_path):
    return pd.read_excel(file_path)


def write_responses_to_excel(responses, output_file_path):
    df = pd.DataFrame(responses)
    df.to_excel(output_file_path, index=False)


# This function is to create a Partner against a Connected User.
def create_account(Xflow_Account, email, legal_name, address, city, country, postal_code, state, nickname, type):
    url = 'https://api.xflowpay.com/v1/accounts'
    headers = {
        "Authorization": API_KEY,
        "Xflow-Account": Xflow_Account,
        'Content-Type': 'application/json'
    }
    data = {
        "business_details": {
            "email": email,
            "legal_name": legal_name,
            "physical_address": {
                "city": city,
                "country": country,
                "line1": address,
                "postal_code": postal_code,
                "state": state
            },
            "type": type
        },
        "nickname": nickname,
        "type": "partner"  # This will always be 'partner' [Don't change it]
    }
    print(f"Request payload for creating Partner: {data}")
    response = requests.post(url, headers=headers, json=data)
    response.raise_for_status()
    print(f"Response for creating partner: {response.json()}")
    return response.json()


# This function is to get the details of Activated USD VBAN of a Connected User.
def list_receive_address(Xflow_Account):
    url = 'https://api.xflowpay.com/v1/addresses?category=xflow_receive&status=activated&currency=USD'
    headers = {
        "Authorization": API_KEY,
        "Xflow-Account": Xflow_Account,
        'Content-Type': 'application/json'
    }
    print(f"Getting the USD VBAN address for user: {Xflow_Account}")

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    print(f"Response of get account status: {response.json()}")
    return response.json()


# This function will transfer the USD funds from platform to Connected User VBAN (The amount will be receivable amount)
def transfer_money(Xflow_Account, amount, currency):
    url = 'https://api.xflowpay.com/v1/transfers'
    headers = {
        "Authorization": API_KEY,
        "Xflow-Account": Xflow_Account,
        'Content-Type': 'application/json'
    }
    data = {
        "from": {
            "account_id": xflow_subject_account_id,
            "amount": amount,
            "currency": currency
        },
        "to": {
            "account_id": Xflow_Account,
            "currency": currency
        },
        "type": "platform_debit"
    }
    print(f"Transferring the Amount for: {Xflow_Account} and request: {data}")

    response = requests.post(url, headers=headers, json=data)
    response.raise_for_status()
    print(f"Response of get transferring amount: {response.json()}")
    return response.json()


# Don't change this function [It's linked to uploading the file]
def download_file(drive_link, destination):
    file_id = extract_file_id(drive_link)
    URL = "https://drive.google.com/uc?export=download"

    session = requests.Session()

    response = session.get(URL, params={'id': file_id}, stream=True)
    token = get_confirm_token(response)

    if token:
        params = {'id': file_id, 'confirm': token}
        response = session.get(URL, params=params, stream=True)

    save_response_content(response, destination)


# Don't change this function [It's linked to uploading the file]
def extract_file_id(drive_link):
    pattern = r'(?:\/d\/|id=)([a-zA-Z0-9_-]+)'
    match = re.search(pattern, drive_link)
    if match:
        return match.group(1)
    else:
        raise ValueError("Invalid Google Drive URL provided.")


# Don't change this function [It's linked to uploading the file]
def get_confirm_token(response):
    for key, value in response.cookies.items():
        if key.startswith('download_warning'):
            return value
    return None


# Don't change this function [It's linked to uploading the file]
def save_response_content(response, destination):
    CHUNK_SIZE = 32768

    # Delete the destination file if it exists
    if os.path.exists(destination):
        os.remove(destination)
        print(f"Existing file '{destination}' deleted.")

    with open(destination, "wb") as f:
        for chunk in response.iter_content(CHUNK_SIZE):
            if chunk:
                f.write(chunk)


# This function is to upload the file of a Connected User.
def upload_file(file_path, Xflow_Account):
    url = 'https://api.xflowpay.com/v1/files'
    headers = {
        "Authorization": API_KEY,
        "Xflow-Account": Xflow_Account,
    }
    payload = {'purpose': 'finance_document'}
    if file_path.startswith('http'):
        local_file_path = 'temp_file.pdf'
        download_file(file_path, local_file_path)
        file_to_upload = open(local_file_path, 'rb')
    else:
        file_to_upload = open(file_path, 'rb')
    encoder = MultipartEncoder(
        fields={
            'file': ('invoice.pdf', file_to_upload, 'application/pdf'),
            'payload': ('', str(payload), 'application/json')
        }
    )
    headers['Content-Type'] = encoder.content_type
    print(f"Request payload for uploading file: {payload}")
    response = requests.post(url, headers=headers, data=encoder)
    response.raise_for_status()
    print(f"Response for uploading file: {response.json()}")
    if file_path.startswith('http'):
        file_to_upload.close()
        os.remove(local_file_path)
    else:
        file_to_upload.close()
    return response.json()


# This function will create the receivable for a Connected User
def create_receivable(account_id, Xflow_Account, amount, currency, due_date, reference_number, creation_date, purpose_code, transaction_type, document_id):
    url = 'http://api-internal.xflowpay.com/v1/receivables'
    headers = {
        "Authorization": API_KEY,
        'Content-Type': 'application/json',
        "Xflow-Account": Xflow_Account,
    }
    creation_date_str = creation_date.strftime('%Y-%m-%d') if isinstance(creation_date, pd.Timestamp) else creation_date
    due_date_str = due_date.strftime('%Y-%m-%d') if isinstance(due_date, pd.Timestamp) else due_date

    data = {
        "account_id": account_id,
        "amount_maximum_reconcilable": amount,
        "currency": currency,
        "invoice": {
            "amount": amount,
            "creation_date": creation_date_str,
            "currency": currency,
            "document": document_id,
            "due_date": due_date_str,
            "reference_number": reference_number
        },
        "purpose_code": purpose_code,
        "transaction_type": transaction_type
    }

    print(f"Request payload for creating receivable: {data}")
    response = requests.post(url, headers=headers, json=data)
    response.raise_for_status()
    print(f"Response for creating receivable: {response.json()}")
    return response.json()


# This function is to confirm a receivable.
def confirm_receivable(receivable_id, Xflow_Account):
    url = f'https://api.xflowpay.com/v1/receivables/{receivable_id}/confirm'
    headers = {
        "Authorization": API_KEY,
        'Content-Type': 'application/json',
        "Xflow-Account": Xflow_Account,
    }
    print(f"Request payload for confirming receivable: {Xflow_Account}")
    response = requests.post(url, headers=headers)
    response.raise_for_status()
    print(f"Response for confirming receivable: {response.json()}")
    return response.json()


# This function will create a VBAN address for a Connected User.
def create_VBAN_address(Xflow_Account, currency):
    url = f'http://api-internal.xflowpay.com/v1/addresses'
    headers = {
        "Authorization": API_KEY,
        'Content-Type': 'application/json',
        "Xflow-Account": Xflow_Account,
    }
    data = {
        "category": "xflow_receive",
        "currency": currency,
        "linked_id": Xflow_Account,
        "linked_object": "account"
    }
    print(f"Request payload for creating VBAN address for {Xflow_Account} and req payload is: {data}")
    response = requests.post(url, headers=headers, json=data)
    response.raise_for_status()
    print(f"Response for creating VBAN is: {response.json()}")
    return response.json()


# This is to call the above functions and to read and write the excel sheet. [Don't make any change in this function]
def process_excel(input_file_path, output_file_path):
    # Read the Excel file
    df = read_excel(input_file_path)
    
    responses = []

    for index, row in df.iterrows():
        try:
            print(f"Processing row: {index + 1}/{len(df)}")
            
            # Extract the values from the row
            Xflow_Account = row['Xflow_Account']
            email = row['email']
            legal_name = row['legal_name']
            address = row['address']
            city = row['city']
            country = row['country']
            postal_code = row['postal_code']
            state = row['state']
            nickname = row['nickname']
            type = row['type']
            amount = row['amount']
            currency = row['currency']
            due_date = row['due_date']
            reference_number = row['reference_number']
            creation_date = row['creation_date']
            purpose_code = row['purpose_code']
            transaction_type = row['transaction_type']
            file_path = row['file_path']

            # Create Account
            account_response = create_account(Xflow_Account, email, legal_name, address, city, country, postal_code, state, nickname, type)
            account_id = account_response['id']

            # Create VBAN Address
            vban_response = create_VBAN_address(Xflow_Account, currency)
            vban_id = vban_response['id']

            # List Receive Address
            receive_address_response = list_receive_address(Xflow_Account)

            # Upload File
            file_response = upload_file(file_path, Xflow_Account)
            document_id = file_response['id']

            # Create Receivable
            receivable_response = create_receivable(account_id, Xflow_Account, amount, currency, due_date, reference_number, creation_date, purpose_code, transaction_type, document_id)
            receivable_id = receivable_response['id']

            # Confirm Receivable
            confirm_receivable_response = confirm_receivable(receivable_id, Xflow_Account)

            # Transfer Money
            transfer_money_response = transfer_money(Xflow_Account, amount, currency)

            # Append the response
            responses.append({
                'Xflow_Account': Xflow_Account,
                'account_response': account_response,
                'vban_response': vban_response,
                'receive_address_response': receive_address_response,
                'file_response': file_response,
                'receivable_response': receivable_response,
                'confirm_receivable_response': confirm_receivable_response,
                'transfer_money_response': transfer_money_response
            })

            print(f"Completed processing row: {index + 1}/{len(df)}")

        except Exception as e:
            print(f"Failed processing row: {index + 1}/{len(df)}. Error: {str(e)}")

    # Write the responses to a new Excel file
    write_responses_to_excel(responses, output_file_path)
    print(f"Output written to {output_file_path}")


if __name__ == "__main__":
    process_excel(EXCEL_FILE, OUTPUT_FILE)
