import requests
import json
import time
import pandas as pd
from datetime import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv

load_dotenv()

GMAIL_USER = os.environ['GMAIL_USER']
GMAIL_PASSWORD = os.environ['GMAIL_PASSWORD']
EMAIL_TO = os.environ['EMAIL_TO']

API_LOGIN = os.environ['API_LOGIN']
API_PASSWORD = os.environ['API_PASSWORD']
API_HOST = os.environ['API_HOST']
API_PORT = os.environ['API_PORT']
API_IP = os.environ['API_IP']

API_BASE_URL = f'http://{API_HOST}:{API_PORT}'

SESSION_ID = None

def authorize_user():
    url = f'{API_BASE_URL}/api/login/'
    payload = {
        'login': API_LOGIN,
        'password': API_PASSWORD,
        'ip': API_IP
    }
    headers = {'Content-Type': 'application/json'}
    response = requests.post(url, data=json.dumps(payload), headers=headers)
    return response.json()

response = authorize_user()
if response.get('status') == 200 and response.get('session_id'):
    SESSION_ID = response['session_id']
    print('<b>Вы успешно залогинились!</b>')
else:
    print('Ошибка авторизации')

def init_session():
    global SESSION_ID
    response = authorize_user()
    if response.get('status') == 200 and response.get('session_id'):
        SESSION_ID = response['session_id']
        print('Вы успешно залогинились!')
        return True
    else:
        print('Ошибка авторизации')
        return False

def active_calls_get():
    url = f'{API_BASE_URL}/api/active_calls_get/'
    payload = {
        'session_id': SESSION_ID,
        'data': {
            "tp_ids": [3576],
            "fields": ["number"]
        }
    }
    headers = {'Content-Type': 'application/json'}

    response = requests.post(url, data=json.dumps(payload), headers=headers)

    newdata = response.json()
    total_calls = len(newdata['data'])
    connected_calls = newdata['connected_calls']

    return total_calls, connected_calls

def save_to_excel(rows, filepath):
    df_data = pd.DataFrame(rows, columns=['time', 'total_calls', 'connected', 'percent'])
    
    result_rows = []
    
    for hour, group in df_data.groupby(df_data['time'].str[:2]):
        result_rows.append(group.values.tolist())
        avg_total = round(group['total_calls'].mean())
        avg_connected = round(group['connected'].mean())
        avg_percent = round(group['percent'].mean(), 2)
        result_rows.append([[f"{hour}:00 AVG", avg_total, avg_connected, avg_percent]])

    final_rows = []
    for block in result_rows:
        for row in block:
            final_rows.append(row)

    df_final = pd.DataFrame(final_rows, columns=['time', 'total_calls', 'connected', 'percent %'])
    df_final.to_excel(filepath, index=False)
    print(f"Сохранено в {filepath}")

def send_email_report(rows, hour):
    df = pd.DataFrame(rows, columns=['time', 'total_calls', 'connected', 'percent %'])
    
    # HTML таблица
    html_table = df.to_html(index=False, border=1)
    html = f"""
    <html><body>
        <h3>Отчёт по активным звонкам (TP 3576) за {hour}:00</h3>
        {html_table}
    </body></html>
    """

    msg = MIMEMultipart('alternative')
    msg['Subject'] = f"Active Calls Report {datetime.now().strftime('%Y-%m-%d')} {hour}:00"
    msg['From'] = GMAIL_USER
    msg['To'] = EMAIL_TO
    msg.attach(MIMEText(html, 'html'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, EMAIL_TO, msg.as_string())
    
    print(f"[{hour}:00] Отчёт отправлен на {EMAIL_TO}")

def run_monitor():
    init_session()

    filepath = f"active_calls_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
    rows = []
    last_hour = None

    while True:
        now = datetime.now()
        current_time = now.strftime('%H:%M')
        current_hour = now.strftime('%H')

        try:
            total_calls, connected_calls = active_calls_get()
            percent = round((connected_calls / total_calls * 100), 2) if total_calls > 0 else 0
            print(f"[{current_time}] Всего: {total_calls}, Подключённых: {connected_calls}, Процент: {percent}%")
            rows.append([current_time, total_calls, connected_calls, percent])
            save_to_excel(rows, filepath)

            # Отправка раз в час
            if last_hour is None:
                last_hour = current_hour
            
            if current_time != last_hour:
                hour_rows = [r for r in rows if r[0].startswith(last_hour)]
                send_email_report(hour_rows, last_hour)
                last_hour = current_hour

        except Exception as e:
            print(f"[{current_time}] Ошибка: {e}")

        time.sleep(5)

run_monitor()



	
