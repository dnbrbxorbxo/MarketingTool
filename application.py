import base64
import re

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify, send_from_directory
import os
import smtplib
from flask import Flask, render_template, request, jsonify
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
from email.mime.image import MIMEImage

from bs4 import BeautifulSoup
app = Flask(__name__)
app.secret_key = os.urandom(24)
app.config['SESSION_TYPE'] = 'filesystem'

# 디버깅을 위한 로그 설정
logging.basicConfig(level=logging.DEBUG)

# SMTP 설정
SMTP_SERVER = 'smtp.naver.com'  # SMTP 서버 주소
SMTP_PORT = 465  # SMTP 포트

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
            emails = df['Emails'].dropna().tolist()
            return jsonify({'status': 'success', 'emails': emails})
        except Exception as e:
            return jsonify({'status': 'error', 'message': str(e)}), 500
    return jsonify({'status': 'error', 'message': 'Invalid file format'}), 400


@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    MailTitle = data.get('MailTitle')
    MailContent = data.get('MailContent')
    email_list = data.get('MailReceive')

    smtp_accounts = [
        (data.get('SMTP_USER'), data.get('SMTP_PASSWORD')),
        (data.get('SMTP_USER1'), data.get('SMTP_PASSWORD1')),
        (data.get('SMTP_USER2'), data.get('SMTP_PASSWORD2')),
        (data.get('SMTP_USER3'), data.get('SMTP_PASSWORD3')),
        (data.get('SMTP_USER4'), data.get('SMTP_PASSWORD4')),
        (data.get('SMTP_USER5'), data.get('SMTP_PASSWORD5')),
        (data.get('SMTP_USER6'), data.get('SMTP_PASSWORD6')),
        (data.get('SMTP_USER7'), data.get('SMTP_PASSWORD7')),
        (data.get('SMTP_USER8'), data.get('SMTP_PASSWORD8')),
        (data.get('SMTP_USER9'), data.get('SMTP_PASSWORD9')),
    ]

    # 수신자를 100명씩 나눔
    def chunk_recipients(lst, n):
        for i in range(0, len(lst), n):
            yield lst[i:i + n]

    recipient_chunks = list(chunk_recipients(email_list, 100))

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

    errors = []
    account_index = 0
    email_count = 0

    for chunk in recipient_chunks:
        sent = False
        retries = 0

        while not sent and retries < len(smtp_accounts):
            smtp_user, smtp_password = smtp_accounts[account_index]
            if smtp_user and smtp_password:
                try:
                    msg['From'] = smtp_user
                    msg['To'] = ', '.join(chunk)

                    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
                    server.login(smtp_user, smtp_password)
                    server.sendmail(smtp_user, chunk, msg.as_string())
                    server.quit()
                    sent = True
                    email_count += len(chunk)

                    if email_count >= 30:
                        account_index = (account_index + 1) % len(smtp_accounts)
                        email_count = 0

                except Exception as e:
                    error_message = f"계정 {account_index + 1}로 이메일 전송 중 오류 발생: {e}"
                    logging.error(error_message)
                    errors.append(error_message)
                    account_index = (account_index + 1) % len(smtp_accounts)
                    retries += 1

        if not sent:
            errors.append(f"수신자 청크에 이메일 전송 실패: {chunk}")

    if errors:
        return jsonify({'status': 'error', 'message': errors}), 500
    else:
        return jsonify({'status': 'success', 'message': '이메일이 성공적으로 전송되었습니다.'})


if __name__ == '__main__':
    app.run(debug=True)
