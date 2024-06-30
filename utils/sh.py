import glob
import xlwings as xw

class sh_calculator():
    def __init__(self):
        self.xlsx_name = '../data/sh.xlsx'
        # app = xw.App(visible=False)
        self.read_excel_file()
        self.fetch_master_data()
            
    def brand_idx(self, x):
        for idx, b in enumerate(self.brands):
            if b == x: 
                break
        return idx+1
    
    def car_idx(self, x):
        models = self.wb.sheets['오토리스(운용&금융)차량모델'].range('C7', 'C164').value 
        for idx, model in enumerate(models):
            if model.strip() == x.strip():
                model_code = model
                break
        return idx+1
    
    def capital_idx(self, capital_name):
        capital_names = self.wb.sheets['브랜드별 딜러사'].range('C7', 'C28').value 
        for idx, cname in enumerate(capital_names):
            if cname == capital_name:
                break
        return idx+1
        
    def read_excel_file(self):
        self.app = xw.App(visible=False)        
        self.wb = self.app.books.open(self.xlsx_name)
        self.app.calculation = 'manual'
        self.app.enable_events = False        

        # self.wb = xw.Book(self.xlsx_name)  # 파일 경로와 이름을 적절히 수정하세요
        self.sheet = self.wb.sheets['오토리스(운용&금융)']
        self.sheet1 = self.wb.sheets['오토리스(운용&금융)차량모델']
  
    def fetch_master_data(self):
        self.brands = self.wb.sheets['오토리스(운용&금융)차량모델'].range('B7', 'B28').value

    def fetch_calculator_parameters(self, input_data):
        self.sheet.range('AI14').value = input_data['param3'] #탁송료 부담 여부 1.포함 2.별도
        self.sheet.range('AD15').value = input_data['param4'] #탁송료
        self.sheet.range('AN5').value = 2 #취득원가 선택 (고정값)
        # self.sheet.range('AN6').value = input_data['param5'] #리스기간 (반복 실행)
        self.sheet.range('AN6').value = 2 #리스기간 (반복 실행)
        self.sheet.range('AK3').value = False # 자동차세 포함 여부 
        # self.sheet.range('AK22').value = input_data['param6'] #운행거리 (반복 실행)
        self.sheet.range('AK22').value = 2 #운행거리 (반복 실행)
        # self.sheet.range('AD26').value = input_data['param7'] #보증금 (세부 선택값)
        # self.sheet.range('AD27').value = input_data['param8'] #잔가 (세부 선택값)
        self.sheet.range('AD27').value = 1 #잔가 (세부 선택값)
        # self.sheet.range('AD28').value = input_data['param9'] #선수금 (세부 선택값)
        self.sheet.range('AD41').value = 0 #Total inc (고정값)
        # self.sheet.range('AC41').value = input_data['param10'] #CM 인센티브 (초기값)
        self.sheet.range('D2').value = 1 #공채 지역 (고정값)
        self.sheet.range('AI22').value = input_data['param11'] #공채선택 
        self.sheet.range('AD14').value = input_data['param12'] #공채할인율
        self.sheet.range('AI25').value = 2 #기타비용 포함 여부 1.포함 2.별도 
        self.sheet.range('AD16').value = 0 #기타비용 
        self.sheet.range('AI17').value = True #취득세 적용 
        self.sheet.range('AI8').value = input_data['param13'] #하이브리드 세제혜택 여부 1.미대상 2.하이브리드 3.전기차
        self.sheet.range('AD19').value = input_data['param14'] #친환경 자동차 보조금 
        self.sheet.range('AD9').value = input_data['param15'] # 차량 가격
        self.sheet.range('AD10').value = input_data['param16'] # 옵션 가격
        self.sheet.range('AD11').value = input_data['param17'] # 할인 가격
        self.sheet.range('AD19').value = input_data['param18'] # 전기차 할인 가격
        
        self.app.calculation = 'automatic'
        self.app.enable_events = True
        self.sheet1.range('B6').value = self.brand_idx(input_data['param1']) #브랜드명
        self.sheet1.range('C6').value = self.car_idx(input_data['param2']) #차종
        self.sheet.range('BO25').value = self.capital_idx(input_data['param0']) #제휴사
    def create_single_report(self):
        report = {
                    "_id": "3",
                    "금융사" : "신한카드" ,
                    "차량가격" : self.sheet.range('H8').value ,
                    "할인가격" : self.sheet.range('H9').value , 
                    "실판매가격" : self.sheet.range('U9').value ,
                    "보증금" : self.sheet.range('J14').value , 
                    "잔존가치" : round(self.sheet.range('H17').value,2) ,
                    "선수금" : self.sheet.range('J15').value ,
                    "월자동차세" : self.sheet.range('J18').value ,
                    "연간운행거리" : self.sheet.range('H20').value ,
                    "월리스료" : self.sheet.range('H23').value ,
                    "등록세" : 0 ,
                    "취득세" : self.sheet.range('V12').value ,
                    "공채" : self.sheet.range('V13').value ,
                    "탁송료" : self.sheet.range('V14').value ,
                    "기타비용" : self.sheet.range('V15').value ,
                    "취득원가" : self.sheet.range('H12').value,
                    "리스기간" : self.sheet.range('H13').value, 
                    "최대잔가" : round(self.sheet.range('AD25').value*100,2),
                    "기준금리" : round(self.sheet.range("AB36").value*100,2),
                    "고잔가" : False
                }
        return report

    def create_iter_report(self):
        leasing_iter = [2,5,6] #36, 48, 60
        res_type = [1,2] #일반잔가, 고잔가
        reports = []
        for i in leasing_iter:
            for j in res_type:
                self.sheet.range('AK49').value = j
                self.sheet.range('AN6').value = i+1 #리스기간
                if j == 2:
                    _id = '4'
                    high_res = True
                else:
                    _id = '3'
                    high_res = False

                report = {
                            "_id": _id,
                            "금융사" : "신한카드" ,
                            "차량가격" : self.sheet.range('H8').value ,
                            "할인가격" : self.sheet.range('H9').value , 
                            "실판매가격" : self.sheet.range('U9').value ,
                            "보증금" : self.sheet.range('J14').value , 
                            "잔존가치" : round(self.sheet.range('H17').value,2) ,
                            "선수금" : self.sheet.range('J15').value ,
                            "월자동차세" : self.sheet.range('J18').value ,
                            "연간운행거리" : self.sheet.range('H20').value ,
                            "월리스료" : self.sheet.range('H23').value ,
                            "등록세" : 0 ,
                            "취득세" : self.sheet.range('V12').value ,
                            "공채" : self.sheet.range('V13').value ,
                            "탁송료" : self.sheet.range('V14').value ,
                            "기타비용" : self.sheet.range('V15').value ,
                            "취득원가" : self.sheet.range('H12').value,
                            "리스기간" : self.sheet.range('H13').value, 
                            "최대잔가" : round(self.sheet.range('AD25').value*100,2),
                            "기준금리" : round(self.sheet.range("AB36").value*100,2),
                            "고잔가":  high_res
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
        
    # def __del__(self):
    #     self.wb.close()
    #     self.app.kill()

if __name__ == '__main__':
    sh = sh_calculator()
    reports = sh.main()