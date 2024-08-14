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


@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    email_list = data.get('MailReceive')
    SMTP_Type = data.get("SMTP_Type")
    MailSenderNM = data.get("MailSenderNM")

    random_number = random.choice([1, 2, 3, 4, 5])
    MailContent = data.get('MailContent' + str(random_number))
    MailTitle = data.get('MailTitle' + str(random_number))

    if SMTP_Type == "NAVER":
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
        while retries < 10 and not sent:
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

                try:
                    # 프록시 설정

                    # 프록시 설정
                    socks.setdefaultproxy(
                        socks.SOCKS5,
                        "proxy.oculus-proxy.com",
                        31115,
                        True,
                        "ba13396172373555b0b863c3af19140f54480fc8d80f2c828c7fb52b7751ee3d1075bb7f5ac0aae2dcecb5e3bbf5779e",
                        "oree8kphuc2b"
                    )
                    socks.wrapmodule(smtplib)

                    msg['From'] = formataddr((Header(MailSenderNM, 'utf-8').encode(), smtp_user))
                    msg['To'] = ', '.join(recipient_list)

                    # 이메일 발송 시도
                    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, timeout=30)
                    server.login(smtp_user, smtp_password)
                    server.sendmail(smtp_user, recipient_list, msg.as_string())
                    server.quit()
                    sent = True
                    break  # 성공적으로 이메일이 발송되면 루프를 종료합니다.

                except Exception as e:
                    retries += 1
                    error_message = f"SMTP 계정 {smtp_index + 1}, 시도 {retries}번, 오류 발생: {e}"
                    logging.error(error_message)
                    errors.append(error_message)

                    if retries >= 10:
                        break  # 10번 시도 후 실패 시 루프 종료

    if not sent:
        return jsonify({'status': 'error', 'message': error_message}), 500
    else:
        return jsonify({'status': 'success', 'message': '이메일이 성공적으로 전송되었습니다.'})


if __name__ == '__main__':
    app.run(debug=True)
