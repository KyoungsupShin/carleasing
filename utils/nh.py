import glob
import xlwings as xw

class nh_calculator():
    def __init__(self):
        self.xlsx_name = '../data/nh.xlsx'
        # app = xw.App(visible=False)
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
        self.wb = xw.Book(self.xlsx_name)  # 파일 경로와 이름을 적절히 수정하세요
        self.sheet = self.wb.sheets['운용리스']  # 시트 이름을 적절히 수정하세요
  
    def fetch_master_data(self):
        self.brands = self.wb.sheets['1'].range('H43', 'I78').value
        self.capital_names = self.sheet.range('BN27', 'BP38').value
        self.options = self.sheet.range('CL3', 'CM15').value         

    def fetch_calculator_parameters(self, input_data):
        self.sheet.range('BO25').value = self.capital_idx(input_data['param0']) #제휴사
        self.sheet.range('BT3').value =  self.brand_idx(input_data['param1']) #브랜드명
        self.sheet.range('BT9').value = self.car_idx(input_data['param2']) #차종
        self.sheet.range('BT11').value = self.model_idx(input_data['param3']) #상세모델 
        self.sheet.range('BK10').value = input_data['param4'] #탁송료 부담 여부 1.포함 2.별도 
        self.sheet.range('BA17').value = input_data['param5'] #탁송료
        self.sheet.range('BM10').value = 1 #취득원가 선택 (고정값)
        self.sheet.range('BG27').value = input_data['param6'] #리스기간 (반복 실행)
        self.sheet.range('BO11').value = input_data['param7'] #운행거리 (반복 실행)
        self.sheet.range('AY26').value = input_data['param8'] #보증금 (세부 선택값)
        self.sheet.range('AY28').value = input_data['param9'] #잔가 (세부 선택값)
        self.sheet.range('AY24').value = input_data['param10'] #선수금 (세부 선택값)
        # sheet.range('AD25').value = 0 #Total inc (고정값)
        self.sheet.range('AY32').value = input_data['param11'] #CM 인센티브 (초기값)
        self.sheet.range('BJ10').value = input_data['param12'] #공채선택 1.포함 2.미포함
        self.sheet.range('BB15').value = input_data['param13'] #공채할인율
        self.sheet.range('BN10').value = 2 #자동차세 포함 여부 1.포함 2. 미포함
        self.sheet.range('AY9').value = input_data['param14'] #차량가격
        # sheet.range('AD33').value = int(input_data['param13']) #기타비용 포함 여부 1.별도 2.포함
        # self.sheet.range('AE33').value = int(input_data['param14']) #기타비용 
        # sheet.range('AD34').value = 2 #인지대 수납 1.차감지급 2.리스료 포함 3.수납완료 
        # self.sheet.range('AG10').value = input_data['param15'] #하이브리드 세제혜택 여부
        # self.sheet.range('AG11').value = input_data['param16'] #친환경 자동차 보조금 
        
    def create_single_report(self):
        report = {
                    "_id": "2",
                    "금융사" : "NH농협캐피탈" ,
                    "차량가격" : self.sheet.range('H7').value ,
                    "할인가격" : self.sheet.range('AP7').value , 
                    "실판매가격" : self.sheet.range('K10').value ,
                    "보증금" : self.sheet.range('AK17').value , 
                    "잔존가치" : round(self.sheet.range('AG18').value,2) ,
                    "선수금" : self.sheet.range('AK16').value ,
                    "월자동차세" : self.sheet.range('AK20').value ,
                    "연간운행거리" : self.sheet.range('AG13').value ,
                    "월리스료" : self.sheet.range('AG22').value ,
                    "등록세" : 0 ,
                    "취득세" : self.sheet.range('N11').value ,
                    "공채" : self.sheet.range('N12').value ,
                    "탁송료" : self.sheet.range('N13').value ,
                    "기타비용" : self.sheet.range('N15').value ,
                    "취득원가" : self.sheet.range('N16').value,
                    "리스기간" : self.sheet.range('AG11').value, 
                    "최대잔가" : round(self.sheet.range('AZ30').value*100,2),
                    "기준금리" : round(self.sheet.range("AY38").value*100,2),
                    "고잔가" : False
                }
        return report

    def create_iter_report(self):
        leasing_iter = [1, 2,3] #36, 48, 60
        reports = []
        for i in leasing_iter:
            self.sheet.range('BG27').value = i+1 #리스기간
            self.sheet.range('AY28').value = self.sheet.range('AZ30').value
            report = {
                        "_id": "2",
                        "금융사" : "NH농협캐피탈" ,
                        "차량가격" : self.sheet.range('H7').value ,
                        "할인가격" : self.sheet.range('AP7').value , 
                        "실판매가격" : self.sheet.range('K10').value ,
                        "보증금" : self.sheet.range('AK17').value , 
                        "잔존가치" : round(self.sheet.range('AG18').value,2) ,
                        "선수금" : self.sheet.range('AK16').value ,
                        "월자동차세" : self.sheet.range('AK20').value ,
                        "연간운행거리" : self.sheet.range('AG13').value ,
                        "월리스료" : self.sheet.range('AG22').value ,
                        "등록세" : 0 ,
                        "취득세" : self.sheet.range('N11').value ,
                        "공채" : self.sheet.range('N12').value ,
                        "탁송료" : self.sheet.range('N13').value ,
                        "기타비용" : self.sheet.range('N15').value ,
                        "취득원가" : self.sheet.range('N16').value,
                        "리스기간" : self.sheet.range('AG11').value, 
                        "최대잔가" : round(self.sheet.range('AZ30').value*100,2),
                        "기준금리" : round(self.sheet.range("AY38").value*100,2),
                        "고잔가" : False
                    }
            reports.append(report)
        return reports
    
    
    def main(self, input_data):
        self.fetch_calculator_parameters(input_data)
        reports = self.create_iter_report()
        return reports

    def main_single(self, input_data):
        self.fetch_calculator_parameters(input_data)
        reports = self.create_single_report()
        return reports

if __name__ == '__main__':
    nh = nh_calculator()
    reports = nh.main()