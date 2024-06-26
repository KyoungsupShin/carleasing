import sys
sys.path.append('../utils')
from se import * 
from flask import Flask, jsonify, request

app = Flask(__name__)    

@app.route('/api/get_se_report', methods=['GET'])
def get_se_report():
    se = se_calculator()
    report = se.main()
    del se 
    return jsonify(report)

@app.route('/api/get_se_report_post', methods=['POST'])
def get_se_report_post():
    input_data = request.get_json() 
    print(input_data)
    se = se_calculator()
    report = se.main(input_data)
    print(report)
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


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8501, debug=True)

