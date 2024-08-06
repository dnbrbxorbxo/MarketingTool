import base64
import random
import re
import subprocess
import time
from email.utils import formataddr

import pandas as pd
import socks
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_from_directory
import os
import smtplib
from flask import Flask, render_template, request, jsonify
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header

import logging
from email.mime.image import MIMEImage

from bs4 import BeautifulSoup
import requests
from bs4 import BeautifulSoup
import re




app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['SESSION_TYPE'] = 'filesystem'

# 디버깅을 위한 로그 설정
logging.basicConfig(level=logging.DEBUG)


@app.route('/')
def home():
    return redirect(url_for('main'))
# Define the 'main' endpoint
@app.route('/main')
def main():
    return render_template('main.html')

@app.route('/download_sample')
def download_sample():
    return send_from_directory(directory='static', filename='샘플엑셀.xlsx', as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'status': 'error', 'message': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'status': 'error', 'message': 'No selected file'}), 400
    if file and file.filename.endswith('.xlsx'):
        try:
            # Read the Excel file
            df = pd.read_excel(file)

            if 'Emails' not in df.columns:
                return jsonify({'status': 'error', 'message': 'No Emails column in the file'}), 400

            # Extract emails
            emails = df['Emails'].dropna().unique().tolist()
            return jsonify({'status': 'success', 'emails': emails})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500
    return jsonify({'status': 'error', 'message': 'Invalid file format'}), 400

def parse_proxies(url):
    text = """223.130.140.89:1080	SOCKS5	HIA	KR	223.130.140.89 (NAVER BUSINESS PLATFORM ASIA PACIFIC PTE. LTD.)	4.499	
100% (2) -	06-aug-2024 10:48 (5 hours ago)
118.67.129.21:1080	SOCKS5	HIA	KR	118.67.129.21 (NAVER BUSINESS PLATFORM ASIA PACIFIC PTE. LTD.)	3.363	
50% (1) -	05-aug-2024 22:27 (17 hours ago)
43.133.81.188:20357	SOCKS5	HIA	KR Seoul	43.133.81.188 (Tencent Building, Kejizhongyi Avenue)	10.491	
50% (1) -	05-aug-2024 05:41 (1 days ago)
43.155.145.159:20357	SOCKS5	HIA	KR Seoul	43.155.145.159 (Tencent Building, Kejizhongyi Avenue)	9.969	
100% (2) -	05-aug-2024 05:39 (1 days ago)
43.155.165.26:20357	SOCKS5	HIA	KR Seoul	43.155.165.26 (Tencent Building, Kejizhongyi Avenue)	7.947	
new -	05-aug-2024 03:44 (1 days ago)
150.109.247.82:20357	SOCKS5	HIA	KR Seoul	150.109.247.82 (Tencent Building, Kejizhongyi Avenue)	8.439	
new -	05-aug-2024 03:43 (1 days ago)
150.109.237.41:20357	SOCKS5	HIA	KR Seoul	150.109.237.41 (Tencent Building, Kejizhongyi Avenue)	10.614	
50% (1) -	05-aug-2024 03:18 (1 days ago)
223.130.143.116:1080	SOCKS5	HIA	KR	223.130.143.116 (NAVER BUSINESS PLATFORM ASIA PACIFIC PTE. LTD.)	1.924	
25% (1) -	03-aug-2024 20:01 (2 days ago)
8.220.245.74:1080	SOCKS5	HIA	KR Seoul	8.220.245.74 (Alibaba US Technology Co., Ltd.)	3.639	
new	03-aug-2024 14:58 (3 days ago)
115.23.87.171:54787	SOCKS5	HIA	KR Naju (Jeollanam-do)	115.23.87.171 (Korea Telecom)	18.174	
3% (7) -	03-aug-2024 09:44 (3 days ago)
223.130.153.191:1080	SOCKS5	HIA	KR	223.130.153.191 (NAVER BUSINESS PLATFORM ASIA PACIFIC PTE. LTD.)	1.711	
33% (1) -	03-aug-2024 03:41 (3 days ago)
146.56.101.184:21681	SOCKS5	HIA	KR Chuncheon (Gangwon-do)	146.56.101.184 (ORACLE-BMC-31898)	20.99	
5% (5) -	02-aug-2024 18:56 (3 days ago)
223.130.152.179:1080	SOCKS5	HIA	KR	223.130.152.179 (NAVER BUSINESS PLATFORM ASIA PACIFIC PTE. LTD.)	1.87	
new -	02-aug-2024 16:40 (3 days ago)
223.130.160.52:1080	SOCKS5	HIA	KR	223.130.160.52 (NAVER BUSINESS PLATFORM ASIA PACIFIC PTE. LTD.)	1.909	
50% (1) +	01-aug-2024 17:51 (4 days ago)
220.90.95.249:52124	SOCKS5	HIA	KR Hwasun-gun (Jeollanam-do)	220.90.95.249 (Korea Telecom)	3.381	
13% (1) -	01-aug-2024 13:34 (5 days ago)
129.154.59.13:43238	SOCKS5	HIA	KR Chuncheon (Gangwon-do)	129.154.59.13 (ORACLE-BMC-31898)	24.194	
5% (3) -	22-jul-2024 16:44 (14 days ago)
43.155.174.22:33333	SOCKS5	HIA	KR Seoul	43.155.174.22 (Tencent Building, Kejizhongyi Avenue)	5.309	
11% (1) +	22-jul-2024 12:50 (15 days ago)
125.141.133.49:5566	SOCKS5	HIA	KR Gwanak-gu (Seoul) !!!	125.141.133.49 (Korea Telecom)
    """

    # Regular expression pattern to match IP:port, type
    pattern = re.compile(r'(\d+\.\d+\.\d+\.\d+):(\d+)\s+(\w+(\s+\(\w+\))?)')

    # Extract matches
    matches = pattern.findall(text)

    # Create a list of tuples (IP, Port, Type)
    proxy_list = [(ip, port, proxy_type) for ip, port, proxy_type, _ in matches]

    return proxy_list
@app.route('/send_email', methods=['POST'])
def send_email():
    url = 'https://spys.one/free-proxy-list/KR/'

    proxies = parse_proxies(url)
    print(proxies)
    if not proxies:
        return jsonify({'status': 'error', 'message': 'No proxies available'}), 500


    data = request.json
    email_list = data.get('MailReceive')
    SMTP_Type = data.get("SMTP_Type")
    MailSenderNM = data.get("MailSenderNM")

    random_number = random.choice([1, 2, 3 ,4 , 5])
    MailContent = data.get('MailContent'+str(random_number))
    MailTitle = data.get('MailTitle'+str(random_number))

    if SMTP_Type == "NAVER" :
        # SMTP 설정
        SMTP_SERVER = 'smtp.naver.com'  # SMTP 서버 주소
        SMTP_PORT = 465  # SMTP 포트
    else:
        # SMTP 설정
        SMTP_SERVER = 'smtp.gmail.com'  # SMTP 서버 주소
        SMTP_PORT = 465  # SMTP 포트

    # 유효한 계정만 리스트에 포함
    smtp_accounts = [
        (data.get(f'SMTP_USER{i}'), data.get(f'SMTP_PASSWORD{i}'))
        for i in range(10)
        if data.get(f'SMTP_USER{i}') is not None and data.get(f'SMTP_PASSWORD{i}') is not None
    ]

    print(smtp_accounts)

    # 수신자를 최대 100명까지 가져오기
    recipient_list = email_list[:100]

    errors = []
    sent = False
    retries = 0
    account_index = 0

    for smtp_index, (smtp_user, smtp_password) in enumerate(smtp_accounts):
        if smtp_user and smtp_password:
            logging.debug(f"Using SMTP account: {smtp_user}")

            msg = MIMEMultipart()
            msg['Subject'] = MailTitle

            # HTML 본문을 파싱하여 base64 이미지를 찾아서 첨부
            soup = BeautifulSoup(MailContent, 'html.parser')
            cid_count = 0

            for img in soup.find_all('img'):
                if 'src' in img.attrs and img.attrs['src'].startswith('data:image'):
                    cid_count += 1
                    img_type = img.attrs['src'].split(';')[0].split('/')[1]
                    img_data = re.sub('^data:image/.+;base64,', '', img.attrs['src'])
                    img_data = base64.b64decode(img_data)
                    image_name = f'image{cid_count}.{img_type}'

                    # 이미지 MIME 객체 생성 및 첨부
                    mime_img = MIMEImage(img_data, _subtype=img_type)
                    mime_img.add_header('Content-Disposition', 'attachment', filename=image_name)
                    mime_img.add_header('Content-ID', f'<{image_name}>')
                    msg.attach(mime_img)

                    # HTML 본문 내 이미지 src를 cid로 변경
                    img.attrs['src'] = f'cid:{image_name}'

            cleaned_html = str(soup)
            msg.attach(MIMEText(cleaned_html, 'html'))

            for proxy_index, proxy in enumerate(proxies):
                try:
                    # Choose a proxy
                    proxy_ip, proxy_port, proxy_type = proxy
                    print(f"Using proxy: {proxy_ip}:{proxy_port} ({proxy_type})")

                    # Set up the proxy with PySocks
                    if proxy_type.upper() == 'SOCKS5':
                        socks.setdefaultproxy(socks.PROXY_TYPE_SOCKS5, proxy_ip, int(proxy_port))
                    else:
                        socks.setdefaultproxy(socks.PROXY_TYPE_HTTP, proxy_ip, int(proxy_port))

                    socks.wrapmodule(smtplib)

                    msg['From'] = formataddr((Header(MailSenderNM, 'utf-8').encode(), smtp_user))
                    msg['To'] = ', '.join(recipient_list)

                    # Attempt to send email
                    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=30)
                    server.login(smtp_user, smtp_password)
                    server.sendmail(smtp_user, recipient_list, msg.as_string())
                    server.quit()
                    sent = True
                    break  # Break proxy loop on successful send

                except Exception as e:
                    error_message = f"SMTP 계정 {smtp_index + 1}, 프록시 {proxy_index + 1} 오류 발생: {e}"
                    logging.error(error_message)
                    errors.append(error_message)

                finally:
                    # Release the proxy after sending the email or failure
                    socks.set_default_proxy(None)

            if sent:
                break  # Break SMTP loop on successful send

    if not sent:
        return jsonify({'status': 'error', 'message': error_message}), 500
    else:
        return jsonify({'status': 'success', 'message': '이메일이 성공적으로 전송되었습니다.'})


if __name__ == '__main__':
    app.run(debug=True)
