import glob
import xlwings as xw

class nh_calculator():
    def __init__(self, xl_app, wb):
        self.app = xl_app
        self.wb = wb
        self.read_excel_file()
        self.fetch_master_data()
        

    def brand_idx(self, x):
        for car in self.brands:
            if car[1] == x:
                brand_code = car[0]
        return brand_code[1:]
    
    def capital_idx(self, capital_name):
        for idx, cname in enumerate(self.capital_names):
            if cname[0] == capital_name:
                break
        return idx+1
    
    def car_idx(self, x):
        models = self.sheet.range('BY3', 'BZ41').value 
        for idx, model in enumerate(models):
            if model[1] == x:
                model_code = model[0]
                break
        return idx+1
    
    def model_idx(self, x):
        models = self.sheet.range('CK4', 'CL40').value 
        for idx, model in enumerate(models):
            if model[1] == x:
                model_code = model[0]
                break
        return idx+1
        
    def read_excel_file(self):     
        self.sheet = self.wb.sheets['운용리스']  # 시트 이름을 적절히 수정하세요
        self.app.calculation = 'manual'
        self.app.enable_events = False        
  
    def fetch_master_data(self):
        self.brands = self.wb.sheets['1'].range('H43', 'I78').value
        self.capital_names = self.sheet.range('BN27', 'BP38').value
        self.options = self.sheet.range('CL3', 'CM15').value         
        self.sheet.range('BM10').value = 1 # 취득원가 선택 (고정값)
        self.sheet.range('BG27').value = 2 # 리스기간 (반복 실행)
        self.sheet.range('BO11').value = 2 # 운행거리 (반복 실행)
        self.sheet.range('AY28').value = 0 # 잔가 (세부 선택값)
        self.sheet.range('BN10').value = 2 # 자동차세 포함 여부 1.포함 2. 미포함
        self.sheet.range('BH25').value = 1 # 공채 지역(인천)
    
    def initialize_data(self):
        self.sheet.range('AY26').value = 0.3 # 보증금 비율
        self.sheet.range('AY24').value = 0 # 선수금 비율
        self.sheet.range('AY32').value = 0 # CM인센티브 비율

    def fetch_calculator_parameters(self, input_data, single=False):
        self.initialize_data()
        self.sheet.range('BK10').value = input_data['delivery_yn'] #탁송료 부담 여부 1.포함 2.별도 
        self.sheet.range('BA17').value = input_data['delivery_price'] #탁송료
        self.sheet.range('BJ10').value = input_data['bond_yn'] #공채선택 1.포함 2.미포함
        self.sheet.range('N12').value = input_data['bond_rate'] #공채할인율
        self.sheet.range('AY9').value = input_data['car_price'] #차량가격
        self.sheet.range('AY10').value = input_data['option_price'] #옵션 가격
        self.sheet.range('AY12').value = input_data['discount_price'] #할인 가격

        if single == True:
            self.sheet.range('BG27').value = input_data['lease_month'] #리스기간 (반복 실행)
            self.sheet.range('BO11').value = input_data['distance'] #운행거리 (반복 실행)
            self.sheet.range('AY24').value = input_data['prepayment_rate'] # 선수금 비율
            self.sheet.range('AY26').value = input_data['deposit_rate'] # 보증금 비율
            self.sheet.range('AY32').value = input_data['sales_rate'] # CM인센티브 비율

        self.app.calculation = 'automatic'
        self.app.enable_events = True
        self.sheet.range('BT3').value =  self.brand_idx(input_data['brand_name']) #브랜드명
        self.sheet.range('BT9').value = self.car_idx(input_data['car_name']) #차종
        self.sheet.range('BT11').value = self.model_idx(input_data['trim_name']) #상세모델 
        self.sheet.range('BO25').value = self.capital_idx(input_data['affiliates_name']) #제휴사
        
        if single == True:
            if input_data['max_res_yn'] == True:
                limit_sum = 1 - (self.sheet.range('AY24').value + self.sheet.range('AY26').value)
                limit_sum = limit_sum if limit_sum < self.sheet.range('AZ30').value + 0.01 else self.sheet.range('AZ30').value  
                self.sheet.range('AY28').value = limit_sum 
            else:
                if self.sheet.range('BB30').value == input_data['residual_rate']:
                    self.sheet.range('AY28').value = self.sheet.range('BB30').value
                else: 
                    if input_data['residual_rate'] <= self.sheet.range('AZ30').value: 
                        self.sheet.range('AY28').value = input_data['residual_rate'] #잔가 (세부 선택값)
                    else:
                        limit_sum = 1 - (self.sheet.range('AY24').value + self.sheet.range('AY26').value)
                        limit_sum = limit_sum if limit_sum < self.sheet.range('AZ30').value + 0.01 else self.sheet.range('AZ30').value  
                        self.sheet.range('AY28').value = limit_sum 

    def create_single_report(self):
        report = {
                    "_id": "2",
                    "금융사" : "NH농협캐피탈" ,
                    "월리스료" : self.sheet.range('AG22').value ,
                    "최대잔가" : round(self.sheet.range('AY28').value*100,2),
                    "기준금리" : round(self.sheet.range("AY38").value*100,2),
                    "초기비용" : self.sheet.range("N21").value
                }
        return report

    def create_iter_report(self):
        leasing_iter = [2, 3, 4] #36, 48, 60
        reports = []
        for i in leasing_iter:
            self.sheet.range('BG27').value = i #리스기간
            limit_sum = 1 - (self.sheet.range('AY24').value + self.sheet.range('AY26').value)
            limit_sum = limit_sum if limit_sum < self.sheet.range('AZ30').value + 0.01 else self.sheet.range('AZ30').value  
            self.sheet.range('AY28').value = limit_sum 
            report = {
            "_id": "2",
            "금융사" : "NH농협캐피탈" ,
            "월리스료" : self.sheet.range('AG22').value ,
            "최대잔가" : round(self.sheet.range('AY28').value*100,2),
            "기준금리" : round(self.sheet.range("AY38").value*100,2),
            "초기비용" : self.sheet.range("N21").value
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
    nh = nh_calculator()
    reports = nh.main()