import glob
import xlwings as xw

class mz_calculator():
    def __init__(self):
        self.xlsx_name = '../data/mz.xlsx'
        # app = xw.App(visible=False)
        self.read_excel_file()
    
    def read_excel_file(self):        
        self.app = xw.App(visible=False)        
        self.wb = self.app.books.open(self.xlsx_name)
        self.app.calculation = 'manual'
        self.app.enable_events = False        

        # self.wb = xw.Book(self.xlsx_name)  # 파일 경로와 이름을 적절히 수정하세요
        self.sheet = self.wb.sheets['운용리스']  # 시트 이름을 적절히 수정하세요
    
    def fetch_calculator_parameters(self, input_data):
        self.sheet.range('AH6').value = input_data['param0'] #제휴사
        self.sheet.range('AG7').value = input_data['param1'] #브랜드명
        self.sheet.range('AG8').value = input_data['param2'] #차종
        self.sheet.range('AG9').value = input_data['param3'] #세부 모델        
        self.sheet.range('AF21').value = input_data['param4']  #탁송료 부담 여부 1.포함 2.별도
        self.sheet.range('AG21').value = input_data['param5']  #탁송료
        self.sheet.range('AG25').value = input_data['param6']  #리스기간 (반복 실행)
        # self.sheet.range('AG26').value = input_data['param7']  #운행거리 (반복 실행)
        self.sheet.range('AG26').value = 20000  #운행거리 (반복 실행)
        self.sheet.range('AG27').value = input_data['param8']  #보증금 (세부 선택값)
        # self.sheet.range('AG29').value = input_data['param9']  #잔가 (세부 선택값)
        self.sheet.range('AG29').value = 0  #잔가 (세부 선택값)
        self.sheet.range('AG28').value = input_data['param10']  #선수금 (세부 선택값)
        self.sheet.range('AG37').value = input_data['param11']  #CM 인센티브 (초기값)
        self.sheet.range('AF19').value = input_data['param12']  #공채선택 
        self.sheet.range('AH20').value = input_data['param13']  #공채할인율
        # self.sheet.range('AF18').value = input_data['param14']  #취득세 수기 작성 여부 
        self.sheet.range('AF18').value = '수기' #취득세 수기 작성 여부 
        self.sheet.range('AJ18').value = input_data['param15']  #취득세 
        self.sheet.range('AH24').value = '차량가 기준' #취득원가 선택 (고정값)
        # self.sheet.range('AF32').value = input_data['param16']  # 자동차세 포함 여부
        self.sheet.range('AF32').value = '미포함'  # 자동차세 포함 여부
        self.sheet.range('AG19').value = '대구' #공채 지역 (고정값)
        # self.sheet.range('AF22').value = input_data['param17'] #기타비용 포함 여부 1.포함 2.별도 
        self.sheet.range('AF22').value = '미포함' #기타비용 포함 여부 1.포함 2.별도 
        self.sheet.range('AG22').value = input_data['param18']  #기타비용 
        self.sheet.range('AJ12').value = input_data['param19']  # 차량 가격 
        self.sheet.range('AG13').value = input_data['param20']  # 옵션 가격 
        self.sheet.range('AG14').value = input_data['param21']  # 할인 가격 
        self.app.calculation = 'automatic'
        self.app.enable_events = True

    def create_single_report(self):
        report = {
                    "_id": "6",
                    "금융사" : "메리츠캐피탈" ,
                    "차량가격" : self.sheet.range('G9').value ,
                    "할인가격" : self.sheet.range('G11').value , 
                    "실판매가격" : self.sheet.range('T11').value ,
                    "보증금" : self.sheet.range('V17').value , 
                    "잔존가치" : round(self.sheet.range('T19').value,2) ,
                    "선수금" : self.sheet.range('V18').value ,
                    "월자동차세" : self.sheet.range('V22').value ,
                    "연간운행거리" : self.sheet.range('T20').value ,
                    "월리스료" : self.sheet.range('T23').value ,
                    "등록세" : 0 ,
                    "취득세" : self.sheet.range('I15').value ,
                    "공채" : self.sheet.range('I16').value ,
                    "탁송료" : self.sheet.range('I17').value ,
                    "기타비용" : self.sheet.range('I18').value ,
                    "취득원가" : self.sheet.range('T15').value,
                    "리스기간" : self.sheet.range('T16').value, 
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
                    "차량가격" : self.sheet.range('G9').value ,
                    "할인가격" : self.sheet.range('G11').value ,
                    "실판매가격" : self.sheet.range('T11').value ,
                    "보증금" : self.sheet.range('V17').value , 
                    "잔존가치" : round(self.sheet.range('T19').value,2) ,
                    "선수금" : self.sheet.range('V18').value ,
                    "월자동차세" : self.sheet.range('V22').value ,
                    "연간운행거리" : self.sheet.range('T20').value ,
                    "월리스료" : self.sheet.range('T23').value ,
                    "등록세" : 0 ,
                    "취득세" : self.sheet.range('I15').value ,
                    "공채" : self.sheet.range('I16').value ,
                    "탁송료" : self.sheet.range('I17').value ,
                    "기타비용" : self.sheet.range('I18').value ,
                    "취득원가" : self.sheet.range('T15').value,
                    "리스기간" : self.sheet.range('T16').value, 
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
        self.fetch_calculator_parameters(input_data)
        reports = self.create_single_report()
        return reports
    def __del__(self):
        self.wb.close()
        self.app.kill()

if __name__ == '__main__':
    mz = mz_calculator()
    reports = mz.main()