import glob
import xlwings as xw

class se_calculator():
    def __init__(self, xl_app, wb):
        self.xlsx_name = '..\data\se.xlsx'
        # app = xw.App(visible=False)
        self.app = xl_app
        self.wb = wb
        self.read_excel_file()
        self.fetch_master_data()

    def read_excel_file(self):   
        # self.app = xw.App(visible=False)        
        # self.wb = self.app.books.open(self.xlsx_name)
        self.app.calculation = 'manual'
        self.app.enable_events = False        
        self.sheet = self.wb.sheets['운용리스']  # 시트 이름을 적절히 수정하세요
  
    def fetch_master_data(self):
        self.brands = self.sheet.range('AR8', 'AR34').value
        self.dealers = self.sheet.range('BD8', 'BE26').value
        self.leasing = self.sheet.range('AX8', 'AX13').value
        self.sheet.range('AD17').value = 2 #취득원가 선택 (고정값)

        self.sheet.range('AD18').value = 2 # 리스기간 (반복 실행)
        self.sheet.range('AD21').value = 0 # 선수금 비율
        self.sheet.range('AD19').value = 2 # 운행거리 (반복 실행)
        self.sheet.range('AD22').value = 0 # 잔가 (세부 선택값) 
        self.sheet.range('AD20').value = 0 # 보증금 비율
        self.sheet.range('AD26').value = 0 # CM인센티브 비율
        self.sheet.range('AD25').value = 0 #Total inc (고정값)
        self.sheet.range('AD32').value = 1 #자동차세 포함 여부 1.별도 2.포함
        self.sheet.range('AD34').value = 2 #인지대 수납 1.차감지급 2.리스료 포함 3.수납완료 

    def fetch_calculator_parameters(self, input_data, single = False):
        #API Input [브랜드명, 모델명, 상세등급, 차량가격, 옵션가격, 할인금액, 취득세 감면대상, 공채할인, 탁송료, 부대비용]
        self.sheet.range('AD6').value = input_data['affiliates_name'] #제휴사
        self.sheet.range('AD9').value = int(input_data['brand_name']) #브랜드명
        self.sheet.range('AD10').value = int(input_data['car_name']) #차종
        self.sheet.range('AD15').value = int(input_data['delivery_yn']) #탁송료 부담 여부 1.별도 2.포함
        self.sheet.range('AE15').value = input_data['delivery_price'] #탁송료
        self.sheet.range('AD30').value = int(input_data['bond_yn']) #공채선택 1.부산승용 2.부산RV(제외) 3.별도부담
        self.sheet.range('AD31').value = float(input_data['bond_rate']) #공채할인율
        self.sheet.range('AD33').value = int(input_data['etc_yn']) #기타비용 포함 여부 1.별도 2.포함
        self.sheet.range('AE33').value = int(input_data['etc_price']) #기타비용 
        self.sheet.range('AG10').value = input_data['hybrid_yn'] #하이브리드 세제혜택 여부
        self.sheet.range('AG11').value = input_data['elec_yn'] #친환경 할인 여부 
        self.sheet.range('AD13').value = input_data['option_price'] #옵션 가격
        self.sheet.range('AD14').value = input_data['discount_price'] #할인 가격
        self.sheet.range('AI11').value = input_data['electric_subsidary'] #친환경 자동차 보조금 
        if single == True:
            self.sheet.range('AD18').value = input_data['lease_month'] #리스기간 (반복 실행)
            self.sheet.range('AD19').value = input_data['distance'] #운행거리 (반복 실행)
            self.sheet.range('AD22').value = input_data['prepayment_rate'] # 선수금 비율
            self.sheet.range('AD20').value = input_data['deposit_rate'] # 보증금 비율
            self.sheet.range('AD26').value = input_data['sales_rate'] # CM인센티브 비율
        self.app.calculation = 'automatic'
        self.app.enable_events = True

        if single == True:
            if input_data['max_res_yn'] == True:
                self.sheet.range('AD21').value = self.sheet.range('AH24').value #잔가 #최대 잔가로 재 설정
            else:
                if self.sheet.range('AM29').value > input_data['residual_rate']:
                    self.sheet.range('AD21').value = self.sheet.range('AM29').value
                else:
                    self.sheet.range('AD21').value = input_data['residual_rate'] #잔가 (세부 선택값)

        self.wb.save('../log/se.xlsx')
    def create_single_report(self):
        report = {
            "_id": "1",
            "금융사" : "KDB캐피탈" ,
            "월리스료" : self.sheet.range('H19').value ,
            "최대잔가" : round(self.sheet.range('AD21').value*100,2),
            "기준금리" : round(self.sheet.range("AD28").value*100,2),
            "고잔가" : False
        }
        return report

    def create_iter_report(self):
        # 반복 변수: 리스 기간, 잔존가치, 약정거리
        # 고정 변수: 딜러사, 선납금, 보증금, 판매수수료, 보조금
        leasing_iter = [2, 3, 4] #36, 48, 60
        reports = []

        for i in leasing_iter:
            self.sheet.range('AD18').value = i+1 #리스기간
            self.sheet.range('AD21').value = self.sheet.range('AH24').value #잔가
            report = {
                "_id": "1",
                "금융사" : "KDB캐피탈" ,
                "월리스료" : self.sheet.range('H19').value ,
                "최대잔가" : round(self.sheet.range('AD21').value*100,2),
                "기준금리" : round(self.sheet.range("AD28").value*100,2),
                "고잔가" : False
            }
            reports.append(report)
        return reports
    
    
    def main(self, input_data):
        self.fetch_calculator_parameters(input_data)
        reports = self.create_iter_report()
        return reports

    def main_single(self, input_data):
        self.fetch_calculator_parameters(input_data, True)
        reports = self.create_single_report()
        return reports

    # def __del__(self):
    #     self.wb.close()
    #     self.app.kill()

if __name__ == '__main__':
    se = se_calculator()
    se.main()



    