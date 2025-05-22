from flask import Flask, request, jsonify
from datetime import datetime
import os
app = Flask(__name__)
@app.route('/receive_mfa', methods=['POST'])
def receive_mfa():
    data = request.get_json()
    if not data or 'code' not in data:
        return jsonify({"error": "mfa_code not provided"}), 400
    mfa_code = data['code']
    timestamp = datetime.now().isoformat()
    with open("mfa_code.txt", "a") as f:
        f.write(f"{mfa_code},{timestamp}\n")
    return jsonify({"status": "MFA code received", "timestamp": timestamp}), 200
@app.route('/latest_mfa', methods=['GET'])
def latest_mfa():
    try:
        if not os.path.exists("mfa_code.txt"):
            return jsonify({"mfa_code": "", "timestamp": ""}), 200
        with open("mfa_code.txt", "r") as f:
            lines = [line.strip() for line in f if line.strip()]
        if not lines:
            return jsonify({"mfa_code": "", "timestamp": ""}), 200
        latest_line = lines[-1]
        latest_code, latest_timestamp = latest_line.split(",", 1)
        return jsonify({"mfa_code": latest_code, "timestamp": latest_timestamp}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8502)