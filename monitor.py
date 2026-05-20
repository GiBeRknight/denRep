import argparse
import logging
import os
import signal
import smtplib
import sys
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests
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

POLL_INTERVAL_SEC = 30
REQUEST_TIMEOUT_SEC = 10
SAVE_INTERVAL_SEC = 60
MAX_RETRIES = 3
TP_IDS = [3576]

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
)
log = logging.getLogger('monitor')


class ApiClient:
    def __init__(self):
        self.session_id = None
        self.http = requests.Session()

    def login(self):
        url = f'{API_BASE_URL}/api/login/'
        payload = {'login': API_LOGIN, 'password': API_PASSWORD, 'ip': API_IP}
        data = self._post(url, payload)
        if data and data.get('status') == 200 and data.get('session_id'):
            self.session_id = data['session_id']
            log.info('Logged in successfully')
            return True
        log.error('Login failed: %s', data)
        return False

    def active_calls(self):
        if self.session_id is None and not self.login():
            return None

        url = f'{API_BASE_URL}/api/active_calls_get/'
        payload = {
            'session_id': self.session_id,
            'data': {'tp_ids': TP_IDS, 'fields': ['number']},
        }
        data = self._post(url, payload)

        # Session likely expired — re-login once and retry
        if not data or data.get('status') in (401, 403) or 'data' not in data:
            log.warning('Session invalid or response malformed, re-authorizing')
            self.session_id = None
            if not self.login():
                return None
            payload['session_id'] = self.session_id
            data = self._post(url, payload)
            if not data or 'data' not in data:
                return None

        total = len(data['data'])
        connected = data.get('connected_calls', 0)
        return total, connected

    def _post(self, url, payload):
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                r = self.http.post(url, json=payload, timeout=REQUEST_TIMEOUT_SEC)
                r.raise_for_status()
                return r.json()
            except (requests.RequestException, ValueError) as e:
                log.warning('POST %s failed (attempt %d/%d): %s',
                            url, attempt, MAX_RETRIES, e)
                if attempt < MAX_RETRIES:
                    time.sleep(2 ** attempt)
        return None


def save_to_excel(rows, filepath):
    if not rows:
        return
    df = pd.DataFrame(rows, columns=['time', 'total_calls', 'connected', 'percent'])

    final_rows = []
    for hour, group in df.groupby(df['time'].str[:2]):
        final_rows.extend(group.values.tolist())
        final_rows.append([
            f'{hour}:00 AVG',
            round(group['total_calls'].mean()),
            round(group['connected'].mean()),
            round(group['percent'].mean(), 2),
        ])

    df_final = pd.DataFrame(final_rows,
                            columns=['time', 'total_calls', 'connected', 'percent %'])
    # Atomic write so a crash mid-write doesn't corrupt the xlsx
    tmp_path = filepath + '.tmp'
    df_final.to_excel(tmp_path, index=False)
    os.replace(tmp_path, filepath)


def send_email_report(rows, hour, test=False):
    if not rows:
        log.info('[%s:00] No rows for this hour, skip email', hour)
        return

    df = pd.DataFrame(rows, columns=['time', 'total_calls', 'connected', 'percent %'])
    html_table = df.to_html(index=False, border=1)
    title_prefix = '[TEST] ' if test else ''
    html = f"""
    <html><body>
        <h3>{title_prefix}Отчёт по активным звонкам (TP {TP_IDS[0]}) за {hour}:00</h3>
        {html_table}
    </body></html>
    """

    subject_prefix = '[TEST] ' if test else ''
    msg = MIMEMultipart('alternative')
    msg['Subject'] = (f"{subject_prefix}Active Calls Report "
                      f"{datetime.now().strftime('%Y-%m-%d')} {hour}:00")
    msg['From'] = GMAIL_USER
    msg['To'] = EMAIL_TO
    msg.attach(MIMEText(html, 'html'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=REQUEST_TIMEOUT_SEC) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, EMAIL_TO, msg.as_string())

    log.info('[%s:00] %sReport sent to %s', hour, subject_prefix, EMAIL_TO)


_stop = False


def _handle_signal(signum, _frame):
    global _stop
    log.info('Received signal %s, shutting down', signum)
    _stop = True


def _interruptible_sleep(seconds):
    for _ in range(seconds):
        if _stop:
            return
        time.sleep(1)


def run_monitor():
    signal.signal(signal.SIGINT, _handle_signal)
    signal.signal(signal.SIGTERM, _handle_signal)

    api = ApiClient()
    if not api.login():
        log.error('Could not log in; exiting')
        sys.exit(1)

    filepath = f'active_calls_{datetime.now().strftime("%Y-%m-%d")}.xlsx'
    rows = []
    last_save_ts = 0.0
    last_hour = datetime.now().strftime('%H')

    while not _stop:
        now = datetime.now()
        current_time = now.strftime('%H:%M')
        current_hour = now.strftime('%H')

        result = api.active_calls()
        if result is not None:
            total_calls, connected_calls = result
            percent = round((connected_calls / total_calls * 100), 2) if total_calls > 0 else 0
            log.info('[%s] total=%d connected=%d percent=%s%%',
                     current_time, total_calls, connected_calls, percent)
            rows.append([current_time, total_calls, connected_calls, percent])

            now_ts = time.time()
            if now_ts - last_save_ts >= SAVE_INTERVAL_SEC:
                try:
                    save_to_excel(rows, filepath)
                except Exception as e:
                    log.exception('save_to_excel failed: %s', e)
                last_save_ts = now_ts

            if current_hour != last_hour:
                hour_rows = [r for r in rows if r[0].startswith(last_hour)]
                try:
                    send_email_report(hour_rows, last_hour)
                except Exception as e:
                    log.exception('send_email_report failed: %s', e)
                last_hour = current_hour

        _interruptible_sleep(POLL_INTERVAL_SEC)

    try:
        save_to_excel(rows, filepath)
    except Exception as e:
        log.exception('Final save_to_excel failed: %s', e)
    log.info('Stopped cleanly. Data saved to %s', filepath)


def run_test():
    """Single-shot: log in, fetch one snapshot, send a test email, exit.

    Use for verifying that .env credentials, API access and Gmail SMTP all
    work end-to-end before leaving the monitor running.
    """
    api = ApiClient()
    if not api.login():
        log.error('API login failed')
        sys.exit(1)

    result = api.active_calls()
    if result is None:
        log.error('Could not fetch active_calls from API')
        sys.exit(1)

    total, connected = result
    percent = round((connected / total * 100), 2) if total > 0 else 0
    now = datetime.now()
    log.info('[%s] total=%d connected=%d percent=%s%% (test)',
             now.strftime('%H:%M'), total, connected, percent)

    rows = [[now.strftime('%H:%M'), total, connected, percent]]
    try:
        send_email_report(rows, now.strftime('%H'), test=True)
    except Exception as e:
        log.exception('Test email failed: %s', e)
        sys.exit(1)

    log.info('Test OK — credentials, API and SMTP all work.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Active calls monitor')
    parser.add_argument('--test', action='store_true',
                        help='Fetch one snapshot, send a test email, and exit')
    args = parser.parse_args()

    if args.test:
        run_test()
    else:
        run_monitor()
