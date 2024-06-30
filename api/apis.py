import sys
sys.path.append('../utils')
from se import * 
from nh import *
from sh import * 
from dgb import * 
from mz import *
from bnk import * 
from flask import Flask, jsonify, request
import argparse 
app = Flask(__name__)    

# @app.route('/api/get_se_report', methods=['GET'])
# def get_se_report():
#     se = se_calculator()
#     report = se.main()
#     del se 
#     return jsonify(report)

@app.route('/api/get_se_report_post', methods=['POST'])
def get_se_report_post():
    input_data = request.get_json() 
    print(input_data)
    se = se_calculator()
    report = se.main(input_data)
    # print(report)
    del se 
    return jsonify(report)

@app.route('/api/get_se_single_report_post', methods=['POST'])
def get_se_single_report_post():
    input_data = request.get_json() 
    print(input_data)
    se = se_calculator()
    report = se.main_single(input_data)
    print(report)
    del se 
    return jsonify(report)

@app.route('/api/get_nh_report_post', methods=['POST'])
def get_nh_report_post():
    input_data = request.get_json() 
    print(input_data)
    nh = nh_calculator()
    report = nh.main(input_data)
    print(report)
    del nh
    return jsonify(report)


@app.route('/api/get_sh_report_post', methods=['POST'])
def get_sh_report_post():
    input_data = request.get_json() 
    print(input_data)
    sh = sh_calculator()
    report = sh.main(input_data)
    print(report)
    del sh
    return jsonify(report)

@app.route('/api/get_dgb_report_post', methods=['POST'])
def get_dgb_report_post():
    input_data = request.get_json() 
    print(input_data)
    dgb = dgb_calculator()
    report = dgb.main(input_data)
    print(report)
    # del dgb
    return jsonify(report)

@app.route('/api/get_mz_report_post', methods=['POST'])
def get_mz_report_post():
    input_data = request.get_json() 
    print(input_data)
    mz = mz_calculator()
    report = mz.main(input_data)
    print(report)
    # del mz
    return jsonify(report)

@app.route('/api/get_bnk_report_post', methods=['POST'])
def get_bnk_report_post():
    input_data = request.get_json() 
    print(input_data)
    bnk = bnk_calculator()
    report = bnk.main(input_data)
    print(report)
    # del mz
    return jsonify(report)



# if __name__ == '__main__':
#     app.run(host='0.0.0.0', port=8501, debug=True)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Run the Flask application.')
    parser.add_argument('--port', type=int, default=8501, help='Port to run the Flask app on')
    args = parser.parse_args()
    app.run(host='0.0.0.0', port=args.port, debug=True)