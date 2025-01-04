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
from kakao import * 
from flask import Flask, jsonify, request, g
import argparse 
import xlwings as xw
import threading
import os 


app = Flask(__name__)    

@app.route('/api/pdfgenerate', methods=['POST'])
def pdfgenerate():
    input_data = request.get_json()
    print(input_data)
    pg = pdf_gen(input_data)
    pg.main()
    share_url = pg.share_url
    # del pg
    return jsonify({'share_url': share_url})

@app.route('/api/request_consultant', methods=['POST'])
def request_consultant():
    input_data = request.get_json()
    print(input_data)
    # km = kakaomsg()
    text = input_data['pdfurl'] + '\n' + input_data['contents'] + '\n'
    asyncio.run(send_pdf_url(text))
    
    
    # km.test_info()
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Run the Flask application.')
    parser.add_argument('--port', type=int, default=8501, help='Port to run the Flask app on')
    parser.add_argument('--excel', type=str, default = '../data/bnk.xlsm', help='excel path')
    args = parser.parse_args()
    app.run(host='0.0.0.0', port=args.port, debug=True, threaded=False)