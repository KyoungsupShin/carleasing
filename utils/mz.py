import glob
import xlwings as xw

class mz_calculator():
    def __init__(self, xl_app, wb):
        self.xlsx_name = '../data/mz.xlsx'
        self.app = xl_app
        self.wb = wb

        # app = xw.App(visible=False)
        self.read_excel_file()
    
    def read_excel_file(self):        
        # self.app = xw.App(visible=False)        
        # self.wb = self.app.books.open(self.xlsx_name)
        self.app.calculation = 'manual'
        self.app.enable_events = False        

        self.sheet = self.wb.sheets['운용리스']  # 시트 이름을 적절히 수정하세요
        self.sheet.range('AG25').value = 36  #리스기간 (반복 실행)
        self.sheet.range('AG26').value = 20000  #운행거리 (반복 실행)
        self.sheet.range('AG27').value = 0  #보증금 (세부 선택값)
        self.sheet.range('AG29').value = 0  #잔가 (세부 선택값)
        self.sheet.range('AF18').value = '수기' #취득세 수기 작성 여부 
        self.sheet.range('AH24').value = '차량가 기준' #취득원가 선택 (고정값)
        self.sheet.range('AF32').value = '미포함'  # 자동차세 포함 여부
        self.sheet.range('AG19').value = '대구' #공채 지역 (고정값)
        self.sheet.range('AF22').value = '미포함' #기타비용 포함 여부 1.포함 2.별도 
    
    def fetch_calculator_parameters(self, input_data, single=False):
        self.sheet.range('AH6').value = input_data['affiliates_name'] #제휴사
        self.sheet.range('AG7').value = input_data['brand_name'] #브랜드명
        self.sheet.range('AG8').value = input_data['car_name'] #차종
        self.sheet.range('AG9').value = input_data['trim_name'] #세부 모델        
        self.sheet.range('AF21').value = input_data['delivery_yn']  #탁송료 부담 여부 1.포함 2.별도
        self.sheet.range('AG21').value = input_data['delivery_price']  #탁송료
        self.sheet.range('AH20').value = input_data['bond_rate']  #공채할인율
        self.sheet.range('AJ18').value = input_data['tax_price']  #취득세 
        self.sheet.range('AG22').value = input_data['etc_price']  #기타비용 
        self.sheet.range('AJ12').value = input_data['car_price']  # 차량 가격 
        self.sheet.range('AG13').value = input_data['option_price']  # 옵션 가격 
        self.sheet.range('AG14').value = input_data['discount_price']  # 할인 가격 
        if single == True:
            self.sheet.range('AG25').value = input_data['lease_month'] #리스기간 (반복 실행)
            self.sheet.range('AG29').value = input_data['residual_rate'] #잔가 (세부 선택값)
            self.sheet.range('AG26').value = input_data['distance'] #운행거리 (반복 실행)
            self.sheet.range('AG28').value = input_data['prepayment_rate'] # 선수금 비율
            self.sheet.range('AG27').value = input_data['deposit_rate'] # 보증금 비율
            self.sheet.range('AG37').value = input_data['sales_rate'] # CM인센티브 비율
            if input_data['max_res_yn'] == True:
                self.sheet.range('AG29').value = self.sheet.range('AG30').value
            else:
                self.sheet.range('AG29').value = input_data['residual_rate'] #잔가 (세부 선택값)
        self.app.calculation = 'automatic'
        self.app.enable_events = True

    def create_single_report(self):
        report = {
                    "_id": "6",
                    "금융사" : "메리츠캐피탈" ,
                    "월리스료" : self.sheet.range('T23').value ,
                    "최대잔가" : round(self.sheet.range('AG29').value*100,2),
                    "기준금리" : round(self.sheet.range("AG41").value*100,2),
                    "고잔가" : False
                }
        return report

    def create_iter_report(self):
        leasing_iter = [36, 48, 60] #36, 48, 60
        reports = []
        for i in leasing_iter:
            self.sheet.range('AG25').value = i #리스기간
            self.sheet.range('AG29').value = self.sheet.range('AG30').value
            report = {
                        "_id": "6",
                        "금융사" : "메리츠캐피탈" ,
                        "월리스료" : self.sheet.range('T23').value ,
                        "최대잔가" : round(self.sheet.range('AG29').value*100,2),
                        "기준금리" : round(self.sheet.range("AG41").value*100,2),
                        "고잔가" : False
                    }
            reports.append(report)
        return reports
    
    def main(self, input_data):
        self.fetch_calculator_parameters(input_data)
        reports = self.create_iter_report()
        return reports

    def main_single(self,input_data):
        self.fetch_calculator_parameters(input_data, True)
        reports = self.create_single_report()
        return reports
    # def __del__(self):
    #     self.wb.close()
    #     self.app.kill()

if __name__ == '__main__':
    mz = mz_calculator()
    reports = mz.main()