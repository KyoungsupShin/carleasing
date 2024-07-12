import glob
import xlwings as xw

class dgb_calculator():
    def __init__(self, xl_app, wb):
        # self.xlsx_name = '../data/dgb.xlsm'
        self.app = xl_app
        self.wb = wb
        self.read_excel_file()
        self.fetch_master_data()
    
    def read_excel_file(self):        
        # self.app = xw.App(visible=False)        
        # self.wb = self.app.books.open(self.xlsx_name)
        self.app.calculation = 'manual'
        self.app.enable_events = False        

        # self.wb = xw.Book(self.xlsx_name)  # 파일 경로와 이름을 적절히 수정하세요
        self.sheet = self.wb.sheets['운용리스_단일']  # 시트 이름을 적절히 수정하세요
        self.sheet1 = self.wb.sheets['AG 입력시트']
        self.sheet2 = self.wb.sheets['계산_운용리스_단일']
    
    def fetch_master_data(self):
        self.sheet.range('BR18').value = True #취득세 수기 작성 여부 
        self.sheet.range('AS29').value = '차량가기준' #취득원가 선택 (고정값)
        self.sheet.range('AS28').value = 36 #리스기간 (반복 실행)
        self.sheet.range('BR22').value = False # 자동차세 포함 여부
        self.sheet.range('AN40').value = 20000 #운행거리 (반복 실행)
        self.sheet.range('AS36').value = 0 #잔가 (세부 선택값)
        self.sheet.range('AS19').value = '대구광역시' #공채 지역 (고정값)
        self.sheet.range('BR21').value = False #기타비용 포함 여부 1.포함 2.별도 
        self.sheet.range('BD24').value = 0 #기타비용 
        self.sheet.range('AS30').value = 0.3 # 보증금 비율
    
    def fetch_calculator_parameters(self, input_data, single=False):
        self.fetch_master_data()
        self.sheet1.range('S9').value = input_data['affiliates_name'] #제휴사
        self.sheet1.range('S7').value = input_data['brand_name'] #브랜드명
        self.sheet.range('AS7').value = input_data['car_name'] #차종
        self.sheet.range('BR20').value = input_data['delivery_yn'] #탁송료 부담 여부 1.포함 2.별도
        self.sheet.range('BD23').value = input_data['delivery_price'] #탁송료-
        self.sheet.range('BR19').value = input_data['bond_yn'] #공채선택 
        self.sheet.range('AS22').value = input_data['bond_rate'] #공채할인율
        self.sheet.range('AI8').value = input_data['hybrid_yn'] #하이브리드 세제혜택 여부 1.미대상 2.하이브리드 3.전기차
        self.sheet2.range('I75').value = input_data['elec_yn'] #친환경 자동차 보조금 여부 
        self.sheet.range('AV20').value = input_data['electric_subsidary'] #친환경 자동차 보조금 
        self.sheet.range('AS11').value = input_data['car_price'] #차량 가격  
        self.sheet.range('BD18').value = input_data['tax_price'] #취득세 
        self.sheet.range('BF12').value = input_data['option_price'] #옵션가격
        self.sheet.range('BF13').value = input_data['discount_price'] #할인가격

        if single == True:
            self.sheet.range('AS28').value = input_data['lease_month'] #리스기간 (반복 실행)
            self.sheet.range('AN40').value = input_data['distance'] #운행거리 (반복 실행)
            self.sheet.range('AS33').value = input_data['prepayment_rate'] # 선수금 비율
            self.sheet.range('AS30').value = input_data['deposit_rate'] # 보증금 비율
            self.sheet.range('AS43').value = input_data['sales_rate'] # CM인센티브 비율

        self.app.calculation = 'automatic'
        self.app.enable_events = True

        if single == True:
            if input_data['max_res_yn'] == True:
                limit_sum = 1 - (self.sheet.range('AS30').value + self.sheet.range('AS33').value)
                limit_sum = limit_sum if limit_sum < 0.63 else 0.62  
                self.sheet.range('AS36').value = limit_sum 
                # self.sheet.range('AS36').value = self.sheet.range('AS38').value
            else:
                if input_data['residual_rate'] < 0.3:
                    self.sheet.range('BR22').value = 0.3
                else:
                    self.sheet.range('BR22').value = input_data['residual_rate'] #잔가 (세부 선택값)

    def create_single_report(self):
        report = {
                    "_id": "5",
                    "금융사" : "DGB캐피탈" ,
                    "월리스료" : self.sheet.range('AA27').value ,
                    "최대잔가" : round(self.sheet.range('AS38').value*100,2),
                    "기준금리" : round(self.sheet.range("AS45").value*100,2),
                    "초기비용" : self.sheet.range("K28").value
                }
        return report

    def create_iter_report(self):
        leasing_iter = [36, 48, 60] #36, 48, 60
        reports = []
        for i in leasing_iter:
            self.sheet.range('AS28').value = i #리스기간
            limit_sum = 1 - (self.sheet.range('AS30').value + self.sheet.range('AS33').value)
            limit_sum = limit_sum if limit_sum < 0.63 else 0.62  
            self.sheet.range('AS36').value = limit_sum 

            # self.sheet.range('AS36').value = self.sheet.range('AS38').value
            report = {
                        "_id": "5",
                        "금융사" : "DGB캐피탈" ,
                        "월리스료" : self.sheet.range('AA27').value ,
                        "최대잔가" : round(self.sheet.range('AS38').value*100,2),
                        "기준금리" : round(self.sheet.range("AS45").value*100,2),
                        "초기비용" : self.sheet.range("K28").value
                    }
            reports.append(report)
        return reports
    
    def main(self, input_data):
        try:
            self.fetch_calculator_parameters(input_data)
            reports = self.create_iter_report()
            return reports

        except Exception as e:
            print(e)
            self.wb.save('../log/errorcheck.xlsm')
            pass

    def main_single(self, input_data):
        try:
            self.fetch_calculator_parameters(input_data, True)
            reports = self.create_single_report()
            return reports
        except Exception as e:
            print(e)
            self.wb.save('../log/errorcheck.xlsm')
            pass

    # def __del__(self):
    #     self.wb.close()
    #     self.app.kill()

if __name__ == '__main__':
    dgb = dgb_calculator()
    reports = dgb.main()