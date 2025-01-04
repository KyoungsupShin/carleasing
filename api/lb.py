from flask import Flask, request, jsonify
import requests
import argparse 

app = Flask(__name__)

class LoadBalancer:
    def __init__(self, servers):
        self.servers = servers
        self.current = 0

    def get_server(self):
        server = self.servers[self.current]
        self.current = (self.current + 1) % len(self.servers)
        return server

se_load_balancer = LoadBalancer(['http://20.39.190.103:8501', 'http://20.39.190.103:8502'])
se_single_load_balancer = LoadBalancer(['http://20.39.190.103:8511', 'http://20.39.190.103:8512'])
nh_load_balancer = LoadBalancer(['http://20.39.190.103:8503', 'http://20.39.190.103:8504'])
nh_single_load_balancer = LoadBalancer(['http://20.39.190.103:8513', 'http://20.39.190.103:8514'])

mz_load_balancer = LoadBalancer(['http://20.39.185.89:8501', 'http://20.39.185.89:8502'])
mz_single_load_balancer = LoadBalancer(['http://20.39.185.89:8511', 'http://20.39.185.89:8512'])
dgb_load_balancer = LoadBalancer(['http://20.39.185.89:8503', 'http://20.39.185.89:8504'])
dgb_single_load_balancer = LoadBalancer(['http://20.39.185.89:8513', 'http://20.39.185.89:8514'])

bnk_load_balancer = LoadBalancer(['http://40.82.152.60:8501', 'http://40.82.152.60:8502'])
bnk_single_load_balancer = LoadBalancer(['http://40.82.152.60:8511', 'http://40.82.152.60:8512'])


@app.route('/se-lb', methods=['POST'])
def se_proxy():
    server_url = se_load_balancer.get_server() + '/api/get_se_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code

@app.route('/se-single-lb', methods=['POST'])
def se_single_proxy():
    server_url = se_single_load_balancer.get_server() + '/api/get_se_single_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code

@app.route('/nh-lb', methods=['POST'])
def nh_proxy():
    server_url = nh_load_balancer.get_server() + '/api/get_nh_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code

@app.route('/nh-single-lb', methods=['POST'])
def nh_single_proxy():
    server_url = nh_single_load_balancer.get_server() + '/api/get_nh_single_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code


@app.route('/mz-lb', methods=['POST'])
def mz_proxy():
    server_url = mz_load_balancer.get_server() + '/api/get_mz_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code

@app.route('/mz-single-lb', methods=['POST'])
def mz_single_proxy():
    server_url = mz_single_load_balancer.get_server() + '/api/get_mz_single_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code


@app.route('/dgb-lb', methods=['POST'])
def dgb_proxy():
    server_url = dgb_load_balancer.get_server() + '/api/get_dgb_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code

@app.route('/dgb-single-lb', methods=['POST'])
def dgb_single_proxy():
    server_url = dgb_single_load_balancer.get_server() + '/api/get_dgb_single_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code


@app.route('/bnk-lb', methods=['POST'])
def bnk_proxy():
    server_url = bnk_load_balancer.get_server() + '/api/get_bnk_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code

@app.route('/bnk-single-lb', methods=['POST'])
def bnk_single_proxy():
    server_url = bnk_single_load_balancer.get_server() + '/api/get_bnk_single_report_post'
    print(server_url)
    print(request.get_json()) 
    headers = {'Content-Type': 'application/json'}
    response = requests.post(server_url, json=request.get_json(), headers=headers)
    
    print(response.json())
    return jsonify(response.json()), response.status_code
    
if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Run the Flask application.')
    parser.add_argument('--port', type=int, default=8600, help='Port to run the Flask app on')
    args = parser.parse_args()
    app.run(host='0.0.0.0', port=args.port)
