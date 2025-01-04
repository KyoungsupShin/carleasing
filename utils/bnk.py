import glob
import xlwings as xw
import datetime

app = xw.App(visible=False)

class bnk_calculator():
    def __init__(self, xl_app, wb):
        self.app = xl_app
        self.wb = wb
        self.read_excel_file()
        self.ag_incentive = 0.05
    
    def read_excel_file(self):
        self.app.calculation = 'manual'
        self.app.enable_events = False        
        self.sheet = self.wb.sheets['운용리스견적']  # 시트 이름을 적절히 수정하세요
        self.sheet1 = self.wb.sheets['Es1']

    def fetch_master_data(self):
        self.sheet1.range('B39').value = 3 #리스기간 (반복 실행)
        self.sheet1.range('B41').value = 3 #운행거리 (반복 실행)
        self.sheet.range('N36').value = 0 #보증금 (세부 선택값)
        self.sheet1.range('B45').value = 1 #잔가 (세부 선택값)
        self.sheet1.range('B98').value = True #취득세 수기 작성 여부 (고정값)
        self.sheet1.range('B194').value = 2 #기타비용 포함 여부 1.포함 2.별도  (고정값)
        self.sheet1.range('B140').value = 1 #취득원가 선택 (고정값)

    def initialize_data(self, input_data):
        self.sheet1.range('B143').value = int(input_data['car_price']) * 0.3 # 보증금 비율
        self.sheet1.range('B146').value = 0 # 선수금 비율
        self.sheet.range('N42').value = self.ag_incentive # CM인센티브 비율

    def brand_idx(self, x):
        brands = self.wb.sheets['Es1'].range('J7', 'J36').value 
        for idx, b in enumerate(brands):
            if b == x: 
                break
        return idx+1
        
    def car_idx(self, x):
        models =self.wb.sheets['Es1'].range('Y20', 'Z300').value 
        models = [item for item in models if None not in item]
        for idx, model in enumerate(models):
            if model[1].replace(u'\xa0',' ').replace(' ', '') == x.replace(' ', ''):
                model_code = model[0]
                break
        return int(model[0])
    
    def model_idx(self, x):
        models =self.wb.sheets['Es1'].range('AN20', 'AO34').value 
        models = [item for item in models if None not in item]        
        for idx, model in enumerate(models):
            if model[1].replace(u'\xa0',' ').replace(' ', '') == x.replace(' ', ''):
                model_code = model[0]
                break
        return int(model[0])
    
    def capital_idx(self, capital_name):
        capital_names = self.wb.sheets['Es1'].range('G12', 'G27').value 
        for idx, cname in enumerate(capital_names):
            if cname == capital_name:
                break
        return idx+1
    
    def fetch_calculator_parameters(self, input_data, single = False):
        self.fetch_master_data()
        self.initialize_data(input_data)
        self.sheet1.range('B31').value = input_data['delivery_yn'] #탁송료 부담 여부 1.포함 2.별도
        self.sheet.range('N16').value = input_data['delivery_price'] #탁송료-
        self.sheet1.range('B191').value = input_data['bond_yn'] #공채 포함 여부  1.포함 2.별도 
        self.sheet1.range('B119').value = input_data['bond_rate'] #공채할인율  
        self.sheet.range('N22').value = input_data['tax_price'] # 취득세 
        self.sheet.range('N24').value = input_data['etc_price'] #기타비용 
        self.sheet1.range('B194').value = input_data['etc_yn'] #기타비용 
        self.sheet.range('N18').value = input_data['electric_subsidary'] #친환경 자동차 보조금 
        self.sheet.range('N13').value = input_data['car_price'] #차량가격
        self.sheet.range('N14').value = input_data['option_price'] #옵션가격
        self.sheet.range('N15').value = input_data['discount_price'] #할인가격
        
        if single == True:
            self.sheet1.range('B39').value = input_data['lease_month'] #리스기간 (반복 실행)
            self.sheet1.range('B41').value = input_data['distance'] #운행거리 (반복 실행)
            self.sheet1.range('B146').value = input_data['prepayment_price'] # 선수금 비율
            self.sheet1.range('B143').value = input_data['deposit_price'] # 보증금 비율
            self.sheet.range('N42').value = input_data['sales_rate'] + self.ag_incentive # CM인센티브 비율

        self.app.calculation = 'automatic'
        self.app.enable_events = True
        self.sheet1.range('B9').value = self.brand_idx(input_data['brand_name']) #브랜드명
        self.sheet1.range('B13').value = self.car_idx(input_data['car_name']) #차종
        self.sheet1.range('B15').value =  self.model_idx(input_data['trim_name']) #상세모델 
        self.sheet1.range('B154').value = self.capital_idx(input_data['affiliates_name']) #제휴사

        if single == True:
            if input_data['max_res_yn'] == True:
                limit_sum = 1 - (self.sheet.range('N38').value + self.sheet.range('N36').value)
                limit_sum = limit_sum if limit_sum < self.sheet1.range('G120').value + 0.01 else self.sheet1.range('G120').value  
                self.sheet1.range('B56').value = limit_sum 
            else:
                if self.sheet1.range('B64').value > input_data['residual_rate']:
                    self.sheet1.range('B56').value = input_data['residual_rate'] # 최소 잔가로 재 설정
                else:
                    if input_data['residual_rate'] <= self.sheet1.range('G120').value:
                        self.sheet1.range('B56').value = input_data['residual_rate'] #잔가 (세부 선택값)
                    else:
                        limit_sum = 1 - (self.sheet.range('N38').value + self.sheet.range('N36').value)
                        limit_sum = limit_sum if limit_sum < self.sheet1.range('G120').value + 0.01 else self.sheet1.range('G120').value  
                        self.sheet1.range('B56').value = limit_sum 
                        
    def create_single_report(self):
        report = {
                    "_id": "7",
                    "금융사" : "BNK캐피탈" ,
                    "월리스료" : self.sheet.range('H26').value ,
                    "최대잔가" : round(self.sheet.range('N33').value*100,2),
                    "기준금리" : round(self.sheet.range("N45").value*100,2),
                    "초기비용" : self.sheet.range("F27").value
                }
        return report

    def create_iter_report(self):
        leasing_iter = [3,2,1] #36, 48, 60
        reports = []
        for i in leasing_iter:
            self.sheet1.range('B39').value = i #리스기간
            limit_sum = 1 - (self.sheet.range('N38').value + self.sheet.range('N36').value)
            limit_sum = limit_sum if limit_sum < self.sheet1.range('G120').value + 0.01 else self.sheet1.range('G120').value  
            self.sheet1.range('B56').value = limit_sum 
            report = {
                        "_id": "7",
                        "금융사" : "BNK캐피탈" ,
                        "월리스료" : self.sheet.range('H26').value ,
                        "최대잔가" : round(self.sheet.range('N33').value*100,2),
                        "기준금리" : round(self.sheet.range("N45").value*100,2),
                        "초기비용" : self.sheet.range("F27").value
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

    def main_single(self, input_data):
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
    bnk = bnk_calculator()
    reports = bnk.main()