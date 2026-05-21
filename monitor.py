import argparse
import logging
import os
import signal
import smtplib
import sys
import time
from datetime import datetime, timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import pandas as pd
import requests
from dotenv import load_dotenv

load_dotenv()

# ============================================================
# Configuration — edit these to tune behaviour
# ============================================================

# Polling / network
POLL_INTERVAL_SEC = 60           # how often to query the API
REQUEST_TIMEOUT_SEC = 10
SAVE_INTERVAL_SEC = 300           # how often to rewrite the xlsx on disk
MAX_RETRIES = 3

# Telephony filter
TP_IDS = [3576]
BRAND_NAME = 'TELES'             # shown in the email header

# Work window — monitor is active while
# WORK_HOUR_START <= hour < WORK_HOUR_END   (24h clock)
WORK_HOUR_START = 7              # 07:00 inclusive
WORK_HOUR_END = 19               # 19:00 exclusive  (=> active 07:00 — 18:59)

# Alerting
ALERT_THRESHOLD = 100            # alert when total_calls < this
ALERT_RECOVERY_COUNT = 3         # consecutive readings >= threshold required
                                 # before another alert can fire

# Daily summary — sent once per day at this hour (24h clock).
# Should be >= WORK_HOUR_END so the day's data is complete.
DAILY_REPORT_HOUR = 21

# ============================================================
# Credentials (from .env)
# ============================================================

GMAIL_USER = os.environ['GMAIL_USER']
GMAIL_PASSWORD = os.environ['GMAIL_PASSWORD']
EMAIL_TO = os.environ['EMAIL_TO']

API_LOGIN = os.environ['API_LOGIN']
API_PASSWORD = os.environ['API_PASSWORD']
API_HOST = os.environ['API_HOST']
API_PORT = os.environ['API_PORT']
API_IP = os.environ['API_IP']

API_BASE_URL = f'http://{API_HOST}:{API_PORT}'

# ============================================================

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
    # Atomic write so a crash mid-write doesn't corrupt the xlsx.
    # Keep the .xlsx suffix on the tmp file — older pandas validates
    # the extension even when an engine is passed explicitly.
    base, ext = os.path.splitext(filepath)
    tmp_path = f'{base}.tmp{ext}'
    df_final.to_excel(tmp_path, index=False, engine='openpyxl')
    os.replace(tmp_path, filepath)


# ============================================================
# Email rendering
# ============================================================

EMAIL_COLUMNS = ['time', 'total_calls', 'connected', 'percent %']

THEMES = {
    'normal': {'accent': '#546e7a', 'tag': ''},          # muted blue-grey
    'test':   {'accent': '#7e6b8f', 'tag': '[TEST] '},   # muted plum
    'alert':  {'accent': '#a86464', 'tag': '[ALERT] '},  # muted brick red
}


def render_html_report(rows, title, theme='normal', columns=None, intro_html=''):
    """Self-contained HTML report. All styles inline — Gmail strips <style> blocks."""
    accent = THEMES[theme]['accent']
    df = pd.DataFrame(rows, columns=columns or EMAIL_COLUMNS)

    th_style = (
        f'padding:10px 14px; background:{accent}; color:#ffffff;'
        f' text-align:left; font-weight:600; border:1px solid {accent};'
        ' font-size:13px; letter-spacing:.3px;'
    )
    td_style_base = (
        'padding:10px 14px; border:1px solid #e6e6e6;'
        ' font-variant-numeric: tabular-nums; color:#222;'
    )

    header_cells = ''.join(f'<th style="{th_style}">{c}</th>' for c in df.columns)

    body_rows = []
    for i, (_, row) in enumerate(df.iterrows()):
        is_avg = isinstance(row.iloc[0], str) and 'AVG' in row.iloc[0]
        if is_avg:
            bg = '#eef1f4'
            extra = f' font-weight:700; border-top:2px solid {accent};'
        else:
            bg = '#fafafa' if i % 2 else '#ffffff'
            extra = ''
        cells = ''.join(
            f'<td style="{td_style_base} background:{bg};{extra}">{v}</td>'
            for v in row
        )
        body_rows.append(f'<tr>{cells}</tr>')

    intro_block = (
        f'<div style="padding:16px 24px; color:#444; font-size:14px; '
        f'border-bottom:1px solid #eee;">{intro_html}</div>'
        if intro_html else ''
    )

    return f"""
<html><body style="margin:0; padding:24px; font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif; background:#f4f5f7;">
  <div style="max-width:720px; margin:0 auto; background:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,.08);">
    <div style="padding:20px 24px; background:{accent}; color:#ffffff;">
      <div style="font-size:18px; font-weight:600; line-height:1.3;">{title}</div>
      <div style="margin-top:4px; font-size:13px; opacity:.85;">
        {datetime.now().strftime('%Y-%m-%d %H:%M')} &middot; {BRAND_NAME}
      </div>
    </div>
    {intro_block}
    <table style="border-collapse:collapse; width:100%; font-size:14px;">
      <thead><tr>{header_cells}</tr></thead>
      <tbody>{''.join(body_rows)}</tbody>
    </table>
  </div>
</body></html>
""".strip()


def _send_email(subject, html_body):
    msg = MIMEMultipart('alternative')
    msg['Subject'] = subject
    msg['From'] = GMAIL_USER
    msg['To'] = EMAIL_TO
    msg.attach(MIMEText(html_body, 'html'))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465, timeout=REQUEST_TIMEOUT_SEC) as server:
        server.login(GMAIL_USER, GMAIL_PASSWORD)
        server.sendmail(GMAIL_USER, EMAIL_TO, msg.as_string())


def send_email_report(rows, hour, theme='normal'):
    if not rows:
        log.info('[%s:00] No rows for this hour, skip email', hour)
        return
    tag = THEMES[theme]['tag']
    title = (f"{tag}{BRAND_NAME} Active Calls Report — "
             f"{datetime.now().strftime('%Y-%m-%d')} {hour}:00")

    body_rows = list(rows)
    if len(rows) > 1:
        df = pd.DataFrame(rows, columns=EMAIL_COLUMNS)
        body_rows.append([
            f'{hour}:00 AVG',
            round(df['total_calls'].mean()),
            round(df['connected'].mean()),
            round(df['percent %'].mean(), 2),
        ])

    html = render_html_report(body_rows, title, theme=theme)
    _send_email(title, html)
    log.info('[%s:00] %sReport sent to %s', hour, tag, EMAIL_TO)


DAILY_COLUMNS = ['#', 'from', 'to', 'readings', 'min total']


def send_daily_report(incidents, date_str, test=False):
    """End-of-day summary: when total_calls dipped below ALERT_THRESHOLD,
    in which time intervals and how many readings each dip lasted."""
    tag = '[TEST] ' if test else ''
    count = len(incidents)
    title = (f"{tag}{BRAND_NAME} Daily Report {date_str} "
             f"— {count} dip" + ('s' if count != 1 else ''))

    if not incidents:
        intro = (f'<p style="margin:0;">No readings below threshold '
                 f'(<b>{ALERT_THRESHOLD}</b>) today.</p>')
        # Single placeholder row so the table doesn't look broken
        rows = [['—', '—', '—', 0, '—']]
    else:
        intro = (f'<p style="margin:0;">Active calls fell below threshold '
                 f'(<b>{ALERT_THRESHOLD}</b>) <b>{count}</b> '
                 f'time' + ('s' if count != 1 else '') +
                 f' today.</p>')
        rows = [
            [i,
             inc['start'],
             inc['end'],
             inc['readings'],
             inc['min_total']]
            for i, inc in enumerate(incidents, 1)
        ]

    html = render_html_report(
        rows, title,
        theme='alert' if incidents and not test else 'normal',
        columns=DAILY_COLUMNS,
        intro_html=intro,
    )
    _send_email(title, html)
    log.info('%sDaily report sent (%d incidents) to %s', tag, count, EMAIL_TO)


def send_alert(rows, current_total, test=False):
    tag = ('[TEST]' + THEMES['alert']['tag']) if test else THEMES['alert']['tag']
    title = f"{tag}{BRAND_NAME} active calls: {current_total}"
    html = render_html_report(rows, title, theme='alert')
    _send_email(title, html)
    log.warning('%sAlert sent: total_calls=%d (threshold %d)',
                tag, current_total, ALERT_THRESHOLD)


# ============================================================
# Work-hours helpers
# ============================================================

def in_work_hours(now):
    return WORK_HOUR_START <= now.hour < WORK_HOUR_END


def next_work_start(now):
    candidate = now.replace(hour=WORK_HOUR_START, minute=0, second=0, microsecond=0)
    if candidate <= now:
        candidate += timedelta(days=1)
    return candidate


# ============================================================
# Main loop
# ============================================================

_stop = False


def _handle_signal(signum, _frame):
    global _stop
    log.info('Received signal %s, shutting down', signum)
    _stop = True


def _interruptible_sleep(seconds):
    for _ in range(int(seconds)):
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

    rows = []
    incidents = []
    current_incident = None
    filepath = None
    work_day_date = None
    last_save_ts = 0.0
    last_hour = None
    alert_active = False
    consecutive_ok = 0
    work_day_finalized = False

    while not _stop:
        now = datetime.now()

        # =====================================================
        #  Outside work hours: finalize day, then wait for the
        #  daily-summary slot at DAILY_REPORT_HOUR, then reset.
        # =====================================================
        if not in_work_hours(now):
            # ---- Finalize the work day exactly once ----
            if not work_day_finalized and rows:
                if last_hour is not None:
                    hour_rows = [r for r in rows if r[0].startswith(last_hour)]
                    try:
                        send_email_report(hour_rows, last_hour)
                    except Exception as e:
                        log.exception('send_email_report failed: %s', e)
                try:
                    save_to_excel(rows, filepath)
                except Exception as e:
                    log.exception('save_to_excel failed: %s', e)
                # If a dip was still open when the work day ended, close it
                if current_incident is not None:
                    incidents.append(current_incident)
                    current_incident = None
                work_day_finalized = True

            # ---- Send the daily summary at/after DAILY_REPORT_HOUR ----
            ready_for_daily = (
                work_day_finalized
                and work_day_date is not None
                and now.hour >= DAILY_REPORT_HOUR
            )
            if ready_for_daily:
                try:
                    send_daily_report(incidents, work_day_date)
                except Exception as e:
                    log.exception('send_daily_report failed: %s', e)
                # Reset for the next work day
                rows = []
                incidents = []
                current_incident = None
                filepath = None
                work_day_date = None
                last_hour = None
                last_save_ts = 0.0
                alert_active = False
                consecutive_ok = 0
                work_day_finalized = False

            # ---- Decide when to wake up ----
            if work_day_finalized:
                # Sleep until DAILY_REPORT_HOUR today
                wake = now.replace(hour=DAILY_REPORT_HOUR,
                                   minute=0, second=0, microsecond=0)
                if wake <= now:
                    # Already past it but somehow not finalized — go to morning
                    wake = next_work_start(now)
            else:
                # Nothing to do — sleep until next work-day start
                wake = next_work_start(now)

            sleep_sec = max(1, (wake - now).total_seconds())
            log.info('Sleeping until %s (%.0f sec)',
                     wake.strftime('%Y-%m-%d %H:%M'), sleep_sec)
            _interruptible_sleep(sleep_sec)
            continue

        # =====================================================
        #  Inside work hours
        # =====================================================
        if filepath is None:
            work_day_date = now.strftime('%Y-%m-%d')
            filepath = f'active_calls_{work_day_date}.xlsx'
            last_hour = now.strftime('%H')

        current_time = now.strftime('%H:%M')
        current_hour = now.strftime('%H')

        result = api.active_calls()
        if result is not None:
            total_calls, connected_calls = result
            percent = round((connected_calls / total_calls * 100), 2) if total_calls > 0 else 0
            log.info('[%s] total=%d connected=%d percent=%s%%',
                     current_time, total_calls, connected_calls, percent)
            rows.append([current_time, total_calls, connected_calls, percent])

            # ---- Alert state machine + incident tracking ----
            if total_calls < ALERT_THRESHOLD:
                consecutive_ok = 0
                if not alert_active:
                    current_incident = {
                        'start': current_time,
                        'end': current_time,
                        'readings': 1,
                        'min_total': total_calls,
                    }
                    hour_rows = [r for r in rows if r[0].startswith(current_hour)]
                    try:
                        send_alert(hour_rows, total_calls)
                    except Exception as e:
                        log.exception('send_alert failed: %s', e)
                    alert_active = True
                else:
                    if current_incident is not None:
                        current_incident['end'] = current_time
                        current_incident['readings'] += 1
                        current_incident['min_total'] = min(
                            current_incident['min_total'], total_calls
                        )
                    log.info('Still below threshold (%d), alert suppressed', total_calls)
            else:
                if alert_active:
                    consecutive_ok += 1
                    log.info('Above threshold (%d/%d OK readings)',
                             consecutive_ok, ALERT_RECOVERY_COUNT)
                    if consecutive_ok >= ALERT_RECOVERY_COUNT:
                        log.info('Alert cleared after %d OK readings', consecutive_ok)
                        alert_active = False
                        consecutive_ok = 0
                        if current_incident is not None:
                            incidents.append(current_incident)
                            current_incident = None

            # ---- Periodic xlsx save ----
            now_ts = time.time()
            if now_ts - last_save_ts >= SAVE_INTERVAL_SEC:
                try:
                    save_to_excel(rows, filepath)
                except Exception as e:
                    log.exception('save_to_excel failed: %s', e)
                last_save_ts = now_ts

            # ---- Hourly email at top of every hour ----
            if current_hour != last_hour:
                hour_rows = [r for r in rows if r[0].startswith(last_hour)]
                try:
                    send_email_report(hour_rows, last_hour)
                except Exception as e:
                    log.exception('send_email_report failed: %s', e)
                last_hour = current_hour

        _interruptible_sleep(POLL_INTERVAL_SEC)

    try:
        if rows and filepath:
            save_to_excel(rows, filepath)
    except Exception as e:
        log.exception('Final save_to_excel failed: %s', e)
    log.info('Stopped cleanly.')


def _fetch_one_snapshot():
    """Log in, fetch one snapshot, return (rows, now). Exit with 1 on failure."""
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
    log.info('[%s] total=%d connected=%d percent=%s%%',
             now.strftime('%H:%M'), total, connected, percent)
    return [[now.strftime('%H:%M'), total, connected, percent]], now, total


def run_test():
    """Send a test hourly-style report with the current snapshot. Ignores work hours."""
    rows, now, _ = _fetch_one_snapshot()
    try:
        send_email_report(rows, now.strftime('%H'), theme='test')
    except Exception as e:
        log.exception('Test email failed: %s', e)
        sys.exit(1)
    log.info('Test OK — credentials, API and SMTP all work.')


def run_test_alert():
    """Send a test alert email with the current snapshot. Ignores work hours
    and the threshold — the alert goes out regardless of total_calls value."""
    rows, _, total = _fetch_one_snapshot()
    try:
        send_alert(rows, total, test=True)
    except Exception as e:
        log.exception('Test alert email failed: %s', e)
        sys.exit(1)
    log.info('Test alert sent.')


def run_test_daily():
    """Send a fake daily report so you can see how it looks."""
    fake_incidents = [
        {'start': '07:00', 'end': '07:01', 'readings': 3, 'min_total': 78},
        {'start': '13:45', 'end': '13:45', 'readings': 1, 'min_total': 92},
        {'start': '16:30', 'end': '16:33', 'readings': 7, 'min_total': 61},
    ]
    today = datetime.now().strftime('%Y-%m-%d')
    try:
        send_daily_report(fake_incidents, today, test=True)
    except Exception as e:
        log.exception('Test daily email failed: %s', e)
        sys.exit(1)
    log.info('Test daily report sent.')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Active calls monitor')
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument('--test', action='store_true',
                      help='Fetch one snapshot, send a test hourly report, exit')
    mode.add_argument('--test-alert', action='store_true',
                      help='Fetch one snapshot, send a test alert email, exit')
    mode.add_argument('--test-daily', action='store_true',
                      help='Send a fake daily summary email and exit')
    args = parser.parse_args()

    if args.test:
        run_test()
    elif args.test_alert:
        run_test_alert()
    elif args.test_daily:
        run_test_daily()
    else:
        run_monitor()
