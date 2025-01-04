import sys
import datetime
sys.path.append('../utils')
from se import * 
from nh import *
from sh import * 
from dgb import * 
from mz import *
from bnk import *
from pdfgenerator import *  
from flask import Flask, jsonify, request, g
import argparse 
import xlwings as xw
import threading
import os 

app = Flask(__name__)    

# 전역 Excel App 객체와 락
xl_app = None
xl_app_lock = threading.Lock()

def initialize_excel_app():
    global xl_app, wb
    try:
        if xl_app is None or not xl_app.api:  # xl_app이 없거나 API가 유효하지 않은 경우
            xl_app = xw.App(visible=False)
            wb = xl_app.books.open(args.excel)
            print('[app] Excel application initialized:', xl_app)
    except Exception as e:
        print(f"Error initializing Excel application: {e}")
        xl_app = None
        wb = None

@app.before_first_request
def setup_excel_app():
    with xl_app_lock:
        initialize_excel_app()

@app.route('/api/get_se_report_post', methods=['POST'])
def get_se_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            se = se_calculator(xl_app, wb)
            report = se.main(input_data)
            print(report)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            del se 
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_se_single_report_post', methods=['POST'])
def get_se_single_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            se = se_calculator(xl_app, wb)
            report = se.main_single(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            del se
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_nh_report_post', methods=['POST'])
def get_nh_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            nh = nh_calculator(xl_app, wb)
            report = nh.main(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            del nh
            xl_app = None
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_nh_single_report_post', methods=['POST'])
def get_nh_single_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            nh = nh_calculator(xl_app, wb)
            report = nh.main_single(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            del nh
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_sh_report_post', methods=['POST'])
def get_sh_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            sh = sh_calculator(xl_app, wb)
            report = sh.main(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_sh_single_report_post', methods=['POST'])
def get_sh_single_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            sh = sh_calculator(xl_app, wb)
            report = sh.main_single(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_mz_report_post', methods=['POST'])
def get_mz_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            mz = mz_calculator(xl_app, wb)
            report = mz.main(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_mz_single_report_post', methods=['POST'])
def get_mz_single_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            mz = mz_calculator(xl_app, wb)
            report = mz.main_single(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            return jsonify({'error': str(e)}), 500


@app.route('/api/get_dgb_report_post', methods=['POST'])
def get_dgb_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            dgb = dgb_calculator(xl_app, wb)
            report = dgb.main(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_dgb_single_report_post', methods=['POST'])
def get_dgb_single_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            dgb = dgb_calculator(xl_app, wb)
            report = dgb.main_single(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_bnk_report_post', methods=['POST'])
def get_bnk_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500
    
            print('\n [1 time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            bnk = bnk_calculator(xl_app, wb)
            report = bnk.main(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            wb.save('../log/bnk.xlsm')
            xl_app = None
            return jsonify({'error': str(e)}), 500

@app.route('/api/get_bnk_single_report_post', methods=['POST'])
def get_bnk_single_report_post():
    global xl_app, wb
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        try:
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500

            print('\n [Requested time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            bnk = bnk_calculator(xl_app, wb)
            report = bnk.main_single(input_data)
            print('\n [Completed time]: ', datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
            return jsonify(report)

        except Exception as e:
            print(f"Error during Excel processing: {e}")
            xl_app = None
            return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Run the Flask application.')
    parser.add_argument('--port', type=int, default=8501, help='Port to run the Flask app on')
    parser.add_argument('--excel', type=str, default = '../data/bnk.xlsm', help='excel path')
    args = parser.parse_args()
    app.run(host='0.0.0.0', port=args.port, debug=True, threaded=False)