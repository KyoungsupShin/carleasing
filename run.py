import argparse 
import pandas as pd 
import requests 
from selenium import webdriver 
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.common.by import By 
from selenium.webdriver.chrome.options import Options 
from webdriver_manager.chrome import ChromeDriverManager
import time  

class global_objects(object):     
    chrome_options = Options()     
    chrome_options.add_argument('start-maximized')     
    # chrome_options.add_argument('--headless')     
    # chrome_options.add_argument('--no-sandbox')     
    chrome_options.add_argument('--disable-dev-shm-usage')     
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')     
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36")     
    headers = {'User-Agent' : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.93 Safari/537.36'}  
    driver_path = "./utils/chromedriver-win64/chromedriver.exe" 

def check_is_limited(driver, btn_obj):
    is_limited = driver.execute_script("""
        const button = arguments[0];
        const style = window.getComputedStyle(button, '::before');
        return style.content;
    """, btn_obj)
    
    if "재고" in is_limited:
        return "재고한정"
    else:
        return ""


def run(filter_brand = '현대'):    
    driver = webdriver.Chrome(options=global_objects.chrome_options)
    driver.get('https://www.carpan.co.kr/service/login')
    
    login_phone2 = driver.find_element(By.NAME, 'phone2') 
    login_phone2.click() #click phone2 
    login_phone2.send_keys('3104')
    
    login_phone3 = driver.find_element(By.NAME, 'phone3') 
    login_phone3.click() #click phone3
    login_phone3.send_keys('0000')
    
    login_pw = driver.find_element(By.NAME, 'pw') 
    login_pw.click() #click phone3
    login_pw.send_keys('0100000')
    
    login_btn = driver.find_element(By.CLASS_NAME, 'buttonBox') 
    login_btn.click() #click phone2 
    
    menu_all_a_tag = driver.find_elements(By.TAG_NAME, 'a')
    menu_lease_select = [i for i in menu_all_a_tag if i.text.strip() =='리스렌트']
    menu_lease_select[0].click()
    
    popup_cls_btn = driver.find_elements(By.CLASS_NAME, 'btnClose')
    if len(popup_cls_btn) == 1:
        print('팝업 있음')
        popup_cls_btn[0].click()
    else:
        print('팝업 없음')
    
    # 브랜드, 모델, 라인업, 트림, 외장, 내장
    option_btns = {}
    
    for cls in driver.find_elements(By.CLASS_NAME, 'cont'):
        btn_list = cls.find_elements(By.TAG_NAME, 'button')
    
        if len(btn_list) > 0:
            for btn in btn_list:
                if len(btn.text.strip()) > 0:
                    # print(btn.text.strip(), btn)
                    option_btns.update({btn.text.split('\n')[0] : btn})
    
    # 브랜드
    for k, v in zip(option_btns.keys(), option_btns.values()):
        print(k, v)
        if k == '브랜드':
            v.click()
            brandSel = driver.find_elements(By.CLASS_NAME, 'brandSel')
            break
            
    time.sleep(1)
    
    kr_brands_info = {}
    kr_brands = brandSel[0].find_element(By.CLASS_NAME, 'box.kr')
    kr_brands_name = [i.text for i in kr_brands.find_elements(By.TAG_NAME, 'li')]
    kr_brands_btn = [i for i in kr_brands.find_elements(By.TAG_NAME, 'button')]
    kr_brands_img_url = [i.get_attribute('src') for i in kr_brands.find_elements(By.TAG_NAME, 'img')]
    
    for n, b, i in zip(kr_brands_name, kr_brands_btn, kr_brands_img_url):
        kr_brands_info.update({n:[b, i]})
    
    
    
    merge_info_list = []
    
    for kr_brands_k, kr_brands_v in zip(kr_brands_info.keys(), kr_brands_info.values()):
        if (kr_brands_k == filter_brand):
            print(kr_brands_k, '를 선택했습니다.')
            btn = kr_brands_v[0]
            btn.click()
            time.sleep(0.5)
        
            modelList = driver.find_element(By.CLASS_NAME, 'modelSel.modelList')
            modelList_name = [i.text for i in modelList.find_elements(By.TAG_NAME, 'li')]
            modelList_btn = [i for i in modelList.find_elements(By.TAG_NAME, 'button')]
            modelList_img_url = [i.get_attribute('src') for i in modelList.find_elements(By.TAG_NAME, 'img')]
            modelList_img_url = [i for i in modelList_img_url if 'model' in i]
        
            model_info = {}
            for n, b, i in zip(modelList_name, modelList_btn, modelList_img_url):
                model_info.update({n:[b, i]})
                        
            for model_info_k, model_info_v in zip(model_info.keys(), model_info.values()):
                print(model_info_k, '를 선택했습니다.')
                btn = model_info_v[0]
                is_limited = check_is_limited(driver, btn)
                btn.click()
                time.sleep(0.5)
                popup = [i for i in driver.find_elements(By.CLASS_NAME, 'cancel') if i.text == '확인']
                
                if len(popup) > 0:
                    popup[0].click()
        
                lineupSel = driver.find_element(By.CLASS_NAME, 'lineupSel.lineupList')
                lineupSel_name = [i.text for i in lineupSel.find_elements(By.TAG_NAME, 'li')]
                lineupSel_btn = [i for i in lineupSel.find_elements(By.TAG_NAME, 'button')]
                lineup_info = {}
                for n, b in zip(lineupSel_name, lineupSel_btn):
                    lineup_info.update({n:b})
                    
                for lineup_k, lineup_v in zip(lineup_info.keys(), lineup_info.values()):
                    print(lineup_k, '를 선택했습니다.')
                    popup = [i for i in driver.find_elements(By.CLASS_NAME, 'cancel') if i.text == '확인']
                    
                    if len(popup) > 0:
                        popup[0].click()
                        
                    lineup_v.click()
                    time.sleep(0.5)
        
                    trimSel = driver.find_element(By.CLASS_NAME, 'trimSel.trimList')
                    trimSel_name = [i.text for i in trimSel.find_elements(By.TAG_NAME, 'li')]
                    trimSel_btn = [i for i in trimSel.find_elements(By.TAG_NAME, 'button')]
                    trimSel_info = {}
                        
                    for n, b in zip(trimSel_name, trimSel_btn):
                        trimSel_info.update({n:b})
        
                    # 5. 외부색상
                    for trimSel_k, trimSel_v in zip(trimSel_info.keys(), trimSel_info.values()):
                        print(trimSel_k, '를 선택했습니다.')
                        trimSel_v.click()
                        time.sleep(0.5)
                        popup = [i for i in driver.find_elements(By.CLASS_NAME, 'cancel') if i.text == '확인']
                        
                        if len(popup) > 0:
                            popup[0].click()
        
                        for k, v in zip(option_btns.keys(), option_btns.values()):
                            if k == '외장':
                                v.click()
                                time.sleep(0.5)
                                break
            
                        colorExtSel = driver.find_element(By.CLASS_NAME, 'colorExtSel.colorList')
                        colorExtSel_name = [i.text for i in colorExtSel.find_elements(By.TAG_NAME, 'li')]
                        colorExtSel_btn = [i for i in colorExtSel.find_elements(By.TAG_NAME, 'button')]
                        colorExtSel_img_color = [i.get_attribute('style') for i in colorExtSel.find_elements(By.CLASS_NAME, 'colorMain')]
                        colorExtSeltrimSel_info = {}
                        colorExtSeltrimSel_info2 = []
            
                        for n, b, i in zip(colorExtSel_name, colorExtSel_btn, colorExtSel_img_color):
                            colorExtSeltrimSel_info.update({n:[b, i]})
                            colorExtSeltrimSel_info2.append([n.strip(), i])
            
                        # 6. 내부색상
                        for colorExtSeltrimSel_info_k, colorExtSeltrimSel_info_v in zip(colorExtSeltrimSel_info.keys(), colorExtSeltrimSel_info.values()):
                            print(colorExtSeltrimSel_info_k, '를 선택했습니다.')
                            btn = colorExtSeltrimSel_info_v[0]
                            btn.click()
                            time.sleep(0.5)
                            break
            
                        for k, v in zip(option_btns.keys(), option_btns.values()):
                            if k == '내장':
                                v.click()
                                time.sleep(0.5)
                                break
            
                        colorIntSel = driver.find_element(By.CLASS_NAME, 'colorIntSel.colorList')
                        colorIntSel_name = [i.text for i in colorIntSel.find_elements(By.TAG_NAME, 'li')]
                        # colorIntSel_btn = [i for i in colorIntSel.find_elements(By.TAG_NAME, 'button')]
                        colorIntSel_img_color = [i.get_attribute('style') for i in colorIntSel.find_elements(By.CLASS_NAME, 'colorMain')]
                        colorIntSeltrimSel_info = []
            
                        for n, i in zip(colorIntSel_name, colorIntSel_img_color):
                            colorIntSeltrimSel_info.append([n.strip(), i])
                            merge_info_dict = {
                                "brand_name" : kr_brands_k,
                                "brand_img" : kr_brands_v[1],
                                "car_name" : model_info_k,
                                "car_img" : model_info_v[1],
                                "model_name" : lineup_k,
                                "model_detial" : trimSel_k.split('\n')[0],
                                "car_price" : trimSel_k.split('\n')[-1],
                                "ext_color" : colorExtSeltrimSel_info2,
                                "int_color" : colorIntSeltrimSel_info,
                                "is_limited" : is_limited
                            }
            
                        merge_info_list.append(merge_info_dict)
                        print('\n\n' + '#' * 30)
                        print(merge_info_dict)
                        print('#' * 30 + '\n\n')
        
                        #재설정
                        for k, v in zip(option_btns.keys(), option_btns.values()):
                            if k == '트림':
                                v.click()
                                time.sleep(0.5)
                                break
                    
                    for k, v in zip(option_btns.keys(), option_btns.values()):
                        if k == '라인업':
                            v.click()
                            time.sleep(0.5)
                            break
                    
                for k, v in zip(option_btns.keys(), option_btns.values()):
                    if k == '모델':
                        v.click()
                        time.sleep(0.5)
                        break
        
            for k, v in zip(option_btns.keys(), option_btns.values()):
                if k == '브랜드':
                    v.click()
                    time.sleep(0.5)
                    break
    
    
    df = pd.DataFrame(merge_info_list)
    brand_list = df['brand_name'].unique()
    for bl in brand_list:
        df[df['brand_name'] == bl].to_csv(f'./data/{bl}.csv', encoding = 'cp949')   

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Run the Flask application.')
    parser.add_argument('--filter_brand')
    args = parser.parse_args()    
    run(args.filter_brand)
