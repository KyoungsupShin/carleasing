import glob
import xlwings as xw

class mz_calculator():
    def __init__(self, xl_app, wb):
        self.xlsx_name = '../data/mz.xlsx'
        self.app = xl_app
        self.wb = wb
        self.read_excel_file()
        self.ag_incentive = 0.0005

    def read_excel_file(self):        
        self.app.calculation = 'manual'
        self.app.enable_events = False        
        self.sheet = self.wb.sheets['운용리스']  # 시트 이름을 적절히 수정하세요
    
    def fetch_master_data(self):
        self.sheet.range('AG25').value = 36  #리스기간 (반복 실행)
        self.sheet.range('AG26').value = 20000  #운행거리 (반복 실행)
        self.sheet.range('AG29').value = 0  #잔가 (세부 선택값)
        self.sheet.range('AF18').value = '수기' #취득세 수기 작성 여부 
        self.sheet.range('AH24').value = '차량가 기준' #취득원가 선택 (고정값)
        self.sheet.range('AF32').value = '미포함'  # 자동차세 포함 여부
        self.sheet.range('AG19').value = '대구' #공채 지역 (고정값)
        self.sheet.range('AF22').value = '미포함' #기타비용 포함 여부 1.포함 2.별도 

    def initialize_data(self):
        self.sheet.range('AG27').value = 0.3 # 보증금 비율
        self.sheet.range('AG28').value = 0 # 선수금 비율
        self.sheet.range('AG37').value = 0 + self.ag_incentive # CM인센티브 비율

    def fetch_calculator_parameters(self, input_data, single=False):
        self.fetch_master_data()
        self.initialize_data()
        self.sheet.range('AH6').value = input_data['affiliates_name'] #제휴사
        self.sheet.range('AG7').value = input_data['brand_name'] #브랜드명
        self.sheet.range('AG8').value = input_data['car_name'] #차종
        self.sheet.range('AG9').value = input_data['trim_name'] #세부 모델        
        self.sheet.range('AF21').value = input_data['delivery_yn']  #탁송료 부담 여부 1.포함 2.별도
        self.sheet.range('AG21').value = input_data['delivery_price']  #탁송료
        self.sheet.range('AJ20').value = input_data['bond_rate']  #공채할인율
        self.sheet.range('AJ18').value = input_data['tax_price']  #취득세 
        self.sheet.range('AG22').value = input_data['etc_price']  #기타비용 
        self.sheet.range('AF22').value = input_data['etc_yn']  #기타비용 
        self.sheet.range('AJ12').value = input_data['car_price']  # 차량 가격 
        self.sheet.range('AG13').value = input_data['option_price']  # 옵션 가격 
        self.sheet.range('AG14').value = input_data['discount_price']  # 할인 가격 
        self.sheet.range('AG37').value = 0 + self.ag_incentive # CM인센티브 비율

        if single == True:
            self.sheet.range('AG25').value = input_data['lease_month'] #리스기간 (반복 실행)
            self.sheet.range('AG26').value = input_data['distance'] #운행거리 (반복 실행)
            self.sheet.range('AH28').value = input_data['prepayment_price'] + 0.00001 # 선수금 비율
            self.sheet.range('AH27').value = input_data['deposit_price'] + 0.000001# 보증금 비율
            self.sheet.range('AG37').value = input_data['sales_rate'] + self.ag_incentive# CM인센티브 비율
        self.app.calculation = 'automatic'
        self.app.enable_events = True
        if single == True:
            if input_data['max_res_yn'] == True:
                limit_sum = 1 - (self.sheet.range('AG27').value + self.sheet.range('AG28').value)
                limit_sum = limit_sum if limit_sum < self.sheet.range('AG30').value + 0.01 else self.sheet.range('AG30').value  
                self.sheet.range('AG29').value = limit_sum 
            else:
                if self.sheet.range('AG31').value > input_data['residual_rate']:
                    self.sheet.range('AG29').value = self.sheet.range('AG31').value 
                else:
                    if input_data['residual_rate'] <= self.sheet.range('AG30').value:
                        self.sheet.range('AG29').value = input_data['residual_rate'] #잔가 (세부 선택값)
                    else:
                        limit_sum = 1 - (self.sheet.range('AG27').value + self.sheet.range('AG28').value)
                        limit_sum = limit_sum if limit_sum < self.sheet.range('AG30').value + 0.01 else self.sheet.range('AG30').value  
                        self.sheet.range('AG29').value = limit_sum 

    def create_single_report(self):
        report = {
                    "_id": "6",
                    "금융사" : "메리츠캐피탈" ,
                    "월리스료" : self.sheet.range('T23').value ,
                    "최대잔가" : round(self.sheet.range('AG29').value*100,2),
                    "기준금리" : round(self.sheet.range("AG41").value*100,2),
                    "초기비용" : self.sheet.range("G22").value
                }
        return report

    def create_iter_report(self):
        leasing_iter = [36, 48, 60] #36, 48, 60
        reports = []
        for i in leasing_iter:
            self.sheet.range('AG25').value = i #리스기간
            limit_sum = 1 - (self.sheet.range('AG27').value + self.sheet.range('AG28').value)
            limit_sum = limit_sum if limit_sum < self.sheet.range('AG30').value + 0.01 else self.sheet.range('AG30').value  
            self.sheet.range('AG29').value = limit_sum 

            # self.sheet.range('AG29').value = self.sheet.range('AG30').value
            report = {
                        "_id": "6",
                        "금융사" : "메리츠캐피탈" ,
                        "월리스료" : self.sheet.range('T23').value ,
                        "최대잔가" : round(self.sheet.range('AG29').value*100,2),
                        "기준금리" : round(self.sheet.range("AG41").value*100,2),
                        "초기비용" : self.sheet.range("G22").value
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
            # self.wb.save('../log/errorcheck.xlsm')
            pass

    def main_single(self,input_data):
        try:
            self.fetch_calculator_parameters(input_data, True)
            reports = self.create_single_report()
            return reports
        except Exception as e:
            print(e)
            # self.wb.save('../log/errorcheck.xlsm')
            pass

    # def __del__(self):
    #     self.wb.close()
    #     self.app.kill()

if __name__ == '__main__':
    mz = mz_calculator()
    reports = mz.main()