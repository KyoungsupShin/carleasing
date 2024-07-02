from flask import Flask, request, jsonify
import xlwings as xw
import os
import threading

app = Flask(__name__)

# 전역 Excel App 객체와 락
xl_app = None
xl_app_lock = threading.Lock()

def initialize_excel_app():
    """Excel 애플리케이션을 초기화하고 전역 변수에 저장"""
    global xl_app
    try:
        if xl_app is None or not xl_app.api:  # xl_app이 없거나 API가 유효하지 않은 경우
            xl_app = xw.App(visible=True)
            print('[app] Excel application initialized:', xl_app)
    except Exception as e:
        print(f"Error initializing Excel application: {e}")
        xl_app = None

def get_excel_workbook():
    """현재 엑셀 애플리케이션에서 워크북을 연다."""
    global xl_app
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        file_path = os.path.join(base_dir, '../data/bnk.xlsm')
        wb = xl_app.books.open(file_path)
        return wb
    except Exception as e:
        print(f"Error opening workbook: {e}")
        return None

@app.before_first_request
def setup_excel_app():
    """애플리케이션이 처음 시작될 때 Excel 애플리케이션 초기화"""
    with xl_app_lock:
        initialize_excel_app()

@app.route('/api/get_bnk_single_report_post', methods=['POST'])
def get_bnk_single_report_post():
    global xl_app
    input_data = request.get_json()
    
    with xl_app_lock:
        if xl_app is None or not xl_app.api:  # Excel 애플리케이션이 유효하지 않은 경우 다시 초기화
            initialize_excel_app()
            if xl_app is None:
                return jsonify({'error': 'Failed to initialize Excel application'}), 500
        
        try:
            wb = get_excel_workbook()
            if wb is None:
                return jsonify({'error': 'Failed to open workbook'}), 500

            print('!@!@!@!@![app]!@!@!@!@!@ ', xl_app)
            print([n.name for n in xl_app.books])
            
            # 필요한 엑셀 작업 수행
            # ...

            return jsonify({'message': 'Success'})
        except Exception as e:
            print(f"Error during Excel processing: {e}")
            # 여기서 xl_app을 None으로 설정하여 다음 요청에서 다시 초기화되도록 합니다.
            xl_app = None
            return jsonify({'error': str(e)}), 500
        finally:
            if wb is not None:
                try:
                    wb.close()
                except Exception as e:
                    print(f"Error closing workbook: {e}")

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=8506, debug=True, threaded=False)
