import glob
import xlwings as xw

class dgb_calculator():
    def __init__(self):
        self.xlsx_name = '../data/dgb.xlsm'
        # app = xw.App(visible=False)
        self.read_excel_file()
    
    def read_excel_file(self):        
        self.wb = xw.Book(self.xlsx_name)  # 파일 경로와 이름을 적절히 수정하세요
        self.sheet = self.wb.sheets['운용리스_단일']  # 시트 이름을 적절히 수정하세요
        self.sheet1 = self.wb.sheets['AG 입력시트']
        self.sheet2 = self.wb.sheets['계산_운용리스_단일']

    def fetch_calculator_parameters(self, input_data):
        self.sheet1.range('S9').value = input_data['param0'] #제휴사
        self.sheet1.range('S7').value = input_data['param1'] #브랜드명
        self.sheet.range('AS7').value = input_data['param2'] #차종
        self.sheet.range('BR18').value = True #취득세 수기 작성 여부 
        self.sheet.range('BR20').value = input_data['param3'] #탁송료 부담 여부 1.포함 2.별도
        self.sheet.range('BD23').value = input_data['param4'] #탁송료-
        self.sheet.range('AS29').value = '차량가기준' #취득원가 선택 (고정값)
        self.sheet.range('AS28').value = input_data['param5'] #리스기간 (반복 실행)
        self.sheet.range('BR22').value = False # 자동차세 포함 여부
        self.sheet.range('AN40').value = input_data['param6'] #운행거리 (반복 실행)
        self.sheet.range('AS30').value = input_data['param7'] #보증금 (세부 선택값)
        self.sheet.range('AS36').value = input_data['param8'] #잔가 (세부 선택값)
        self.sheet.range('AS33').value = input_data['param9'] #선수금 (세부 선택값)
        self.sheet.range('AS43').value = input_data['param10'] #CM 인센티브 (초기값)
        self.sheet.range('AS19').value = '대구광역시' #공채 지역 (고정값)
        self.sheet.range('BR19').value = input_data['param11'] #공채선택 
        self.sheet.range('BD21').value = input_data['param12'] #공채할인율
        self.sheet.range('BR21').value = False #기타비용 포함 여부 1.포함 2.별도 
        self.sheet.range('BD24').value = 0 #기타비용 
        self.sheet.range('AI8').value = input_data['param13'] #하이브리드 세제혜택 여부 1.미대상 2.하이브리드 3.전기차
        self.sheet2.range('I75').value = input_data['param14'] #친환경 자동차 보조금 여부 
        self.sheet.range('AV20').value = input_data['param15'] #친환경 자동차 보조금 
        self.sheet.range('AS11').value = input_data['param16'] #상세모델 
        self.sheet.range('BD18').value = input_data['param17']
        
    def create_single_report(self):
        report = {
                    "_id": "5",
                    "금융사" : "DGB캐피탈" ,
                    "차량가격" : self.sheet.range('K12').value ,
                    "할인가격" : self.sheet.range('K14').value , 
                    "실판매가격" : self.sheet.range('K17').value ,
                    "보증금" : self.sheet.range('AD19').value , 
                    "잔존가치" : round(self.sheet.range('AA23').value,2) ,
                    "선수금" : self.sheet.range('AD21').value ,
                    "월자동차세" : self.sheet.range('AA26').value ,
                    "연간운행거리" : self.sheet.range('AA18').value ,
                    "월리스료" : self.sheet.range('AA27').value ,
                    "등록세" : 0 ,
                    "취득세" : self.sheet.range('K18').value ,
                    "공채" : self.sheet.range('K19').value ,
                    "탁송료" : self.sheet.range('K20').value ,
                    "기타비용" : self.sheet.range('K21').value ,
                    "취득원가" : self.sheet.range('K24').value,
                    "리스기간" : self.sheet.range('AA17').value, 
                    "최대잔가" : round(self.sheet.range('AS38').value*100,2),
                    "기준금리" : round(self.sheet.range("AS45").value*100,2),
                    "고잔가" : False
                }
        return report

    def create_iter_report(self):
        leasing_iter = [36, 48, 60] #36, 48, 60
        reports = []
        for i in leasing_iter:
            self.sheet.range('AS28').value = i #리스기간
            self.sheet.range('AS36').value = self.sheet.range('AS38').value
            report = {
                        "_id": "5",
                        "금융사" : "DGB캐피탈" ,
                        "차량가격" : self.sheet.range('K12').value ,
                        "할인가격" : self.sheet.range('K14').value , 
                        "실판매가격" : self.sheet.range('K17').value ,
                        "보증금" : self.sheet.range('AD19').value , 
                        "잔존가치" : round(self.sheet.range('AA23').value,2) ,
                        "선수금" : self.sheet.range('AD21').value ,
                        "월자동차세" : self.sheet.range('AA26').value ,
                        "연간운행거리" : self.sheet.range('AA18').value ,
                        "월리스료" : self.sheet.range('AA27').value ,
                        "등록세" : 0 ,
                        "취득세" : self.sheet.range('K18').value ,
                        "공채" : self.sheet.range('K19').value ,
                        "탁송료" : self.sheet.range('K20').value ,
                        "기타비용" : self.sheet.range('K21').value ,
                        "취득원가" : self.sheet.range('K24').value,
                        "리스기간" : self.sheet.range('AA17').value, 
                        "최대잔가" : round(self.sheet.range('AS38').value*100,2),
                        "기준금리" : round(self.sheet.range("AS45").value*100,2),
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
    dgb = dgb_calculator()
    reports = dgb.main()