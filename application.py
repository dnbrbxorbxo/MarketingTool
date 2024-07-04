import base64
import re
from flask import Flask, render_template, request, redirect, url_for, flash, session, jsonify
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


@app.route('/send_email', methods=['POST'])
def send_email():
    data = request.json
    MailTitle = data.get('MailTitle')
    MailContent = data.get('MailContent')

    SMTP_USER = data.get('SMTP_USER')
    SMTP_PASSWORD = data.get('SMTP_PASSWORD')
    MailReceive = data.get('MailReceive1')

    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = MailReceive
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
        server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, MailReceive, msg.as_string())
        server.quit()

        return jsonify({'status': 'success', 'message': 'Email sent successfully'})

    except Exception as e:
        logging.error(f"Error sending email: {e}")
        return jsonify({'status': 'error', 'message': str(e)})


if __name__ == '__main__':
    app.run(debug=True)
