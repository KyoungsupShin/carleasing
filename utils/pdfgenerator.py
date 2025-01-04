from pathlib import Path
import xlwings as xw
import os
import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
SCOPES = ['https://www.googleapis.com/auth/drive']

class pdf_gen():
    def __init__(self, input_data):
        self.input_data = input_data
        self.folder_id = '1r0Fd2UGxzhbwkzaaIBwWaOSn4_TKETst'

    def save_pdf(self):
        with xw.App() as app:
            app.visible = False
            self.book = app.books.open('../data/견적서양식.xlsx')
            sheet = self.book.sheets[0]
            sheet.page_setup.print_area = 'A1:AH54'
            print(self.input_data)
            sheet.range("J9").value = self.input_data['brand_name']
            sheet.range("J10").value = self.input_data['affiliates_name']
            sheet.range("J11").value = self.input_data['car_name'] + ' ' + self.input_data['trim_name']
            sheet.range("J13").value = self.input_data['car_price']
            sheet.range("J14").value = self.input_data['option_price']
            sheet.range("J15").value = self.input_data['discount_price']
            sheet.range("J16").value = self.input_data['total_price']
            sheet.range("J17").value = self.input_data['final_price'] #취득원가

            sheet.range("Y13").value = self.input_data['tax_price']
            sheet.range("Y14").value = self.input_data['bond_rate']
            sheet.range("AD14").value = self.input_data['bond_yn']
            sheet.range("Y15").value = self.input_data['delivery_price']
            sheet.range("AD15").value = self.input_data['delivery_yn']
            sheet.range("Y16").value = self.input_data['etc_price']
            sheet.range("AD16").value = self.input_data['etc_yn']
            sheet.range("Y17").value = self.input_data['total_taxprice']

            if len(self.input_data['lease_month']) >= 1:
                sheet.range("J20").value = self.input_data['lease_month'][0] 
                sheet.range("J21").value = self.input_data['distance'][0]
                sheet.range("J22").value = self.input_data['deposit_rate'][0]
                sheet.range("L22").value = self.input_data['deposit_price'][0]
                sheet.range("J24").value = self.input_data['prepayment_rate'][0]
                sheet.range("L24").value = self.input_data['prepayment_price'][0]
                sheet.range("J26").value = self.input_data['residual_rate'][0]
                sheet.range("L26").value = self.input_data['residual_price'][0]
                sheet.range("J28").value = self.input_data['monthly_lease'][0]
                sheet.range("J30").value = self.input_data['sales_rate'][0]
                sheet.range("J31").value = self.input_data['init_price'][0]

            if len(self.input_data['lease_month']) >= 2:
                sheet.range("R20").value = self.input_data['lease_month'][1]
                sheet.range("R21").value = self.input_data['distance'][1]
                sheet.range("R22").value = self.input_data['deposit_rate'][1]
                sheet.range("T22").value = self.input_data['deposit_price'][1]
                sheet.range("R24").value = self.input_data['prepayment_rate'][1]
                sheet.range("T24").value = self.input_data['prepayment_price'][1]
                sheet.range("R26").value = self.input_data['residual_rate'][1]
                sheet.range("T26").value = self.input_data['residual_price'][1]
                sheet.range("R28").value = self.input_data['monthly_lease'][1]
                sheet.range("R30").value = self.input_data['sales_rate'][1]
                sheet.range("R31").value = self.input_data['init_price'][1]

            if len(self.input_data['lease_month']) >= 3:
                sheet.range("Z20").value = self.input_data['lease_month'][2] 
                sheet.range("Z21").value = self.input_data['distance'][2]
                sheet.range("Z22").value = self.input_data['deposit_rate'][2]
                sheet.range("AB22").value = self.input_data['deposit_price'][2]
                sheet.range("Z24").value = self.input_data['prepayment_rate'][2]
                sheet.range("AB24").value = self.input_data['prepayment_price'][2]
                sheet.range("Z26").value = self.input_data['residual_rate'][2]
                sheet.range("AB26").value = self.input_data['residual_price'][2]
                sheet.range("Z28").value = self.input_data['monthly_lease'][2]
                sheet.range("Z30").value = self.input_data['sales_rate'][2]
                sheet.range("Z31").value = self.input_data['init_price'][2]

            # sheet.range("X6").value = self.input_data['customer_name']
            sheet.range("X6").value = 'VIP'
            sheet.range("C52").value = self.input_data['salesname']
            sheet.range("J53").value = self.input_data['salesnumber']

            # current_work_dir = os.getcwd()
            self.filename = "../data/pdf/{}.pdf".format(self.input_data['quoteid'])
            # pdf_path = Path('../data/pdf/', )
            sheet.to_pdf(path=self.filename, show=False)
        
    def get_drive_service(self):
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('cred.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return build('drive', 'v3', credentials=creds)

    def upload_to_drive(self):
        service = self.get_drive_service()
        file_metadata = {
            'name': self.filename.split('/')[-1],
            'parents': [self.folder_id]  # 특정 폴더에 업로드하기 위한 설정
        }
        if self.folder_id:
            file_metadata['parents'] = [self.folder_id]
        media = MediaFileUpload(self.filename, resumable=True)
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"파일이 업로드되었습니다. 파일 ID: {file.get('id')}")
        self.share_url = 'https://drive.google.com/file/d/{}/view?usp=drive_link'.format(file.get('id'))

    def main(self):
        self.save_pdf()
        self.get_drive_service()
        self.upload_to_drive()

    # def __del__(self):
    #     self.book.close()
        # self.book.kill()