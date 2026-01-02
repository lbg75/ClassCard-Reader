# 이 파일은 엑셀에서 학생 정보를 읽고 웹으로 실시간 인원수를 보여줍니다.
# 초보자용으로 최대한 쉬운 말로 주석을 달았습니다.
# - 엑셀 파일을 직접 읽거나(저장된 파일)
# - Excel이 열려 있을 때는 COM을 통해 메모리상의 최신 값을 읽습니다.
# - 변경이 있으면 웹 클라이언트에 실시간으로 보냅니다.

import os
import time
import threading
import pandas as pd
from flask import Flask, render_template
from flask_socketio import SocketIO
import re

EXCEL_PATH = 'entries.xlsx'  # 저장된 엑셀 파일 이름
POLL_INTERVAL = 2  # 몇 초마다 변경을 확인할지 (초 단위)

app = Flask(__name__)
socketio = SocketIO(app, async_mode='eventlet')

# Windows에서 Excel이 열려 있을 때, 그 안의 최신 값을 읽기 위해 pywin32를 사용합니다.
try:
    import win32com.client as win32
    HAVE_COM = True
except Exception:
    # pywin32가 없거나 Windows가 아닐 수 있습니다. 그럼 파일을 직접 읽습니다.
    HAVE_COM = False


def read_from_com(path):
    """Excel이 열려 있을 때 메모리 값(저장되지 않은 변경 포함)을 읽어옵니다.
    성공하면 pandas DataFrame, 실패하면 None을 반환합니다.
    (초보자용: Excel 창에서 바로 읽어오는 방법입니다.)"""
    if not HAVE_COM:
        return None
    try:
        excel = win32.Dispatch('Excel.Application')
        abs_path = os.path.abspath(path).lower()
        # find open workbook matching the path
        for wb in excel.Workbooks:
            try:
                if wb.FullName.lower() == abs_path:
                    ws = wb.Worksheets(1)
                    used = ws.UsedRange
                    values = used.Value
                    if not values:
                        return None
                    # values may be tuple of tuples; first row = headers
                    rows = list(values)
                    if len(rows) < 2:
                        return None
                    headers = [str(h) for h in rows[0]]
                    data = rows[1:]
                    df = pd.DataFrame(list(data), columns=headers)
                    return df
            except Exception:
                continue
    except Exception:
        return None
    return None


def read_and_count():
    """현재 엑셀 데이터를 읽고 1/2/3학년별 인원수를 세어서 반환합니다.
    우선 COM(실시간)을 시도하고 실패하면 저장된 파일을 읽습니다."""
    # 1) Try COM (captures unsaved changes)
    df = read_from_com(EXCEL_PATH)
    if df is None:
        # 2) Fallback to reading the saved file
        try:
            df = pd.read_excel(EXCEL_PATH, engine='openpyxl')
        except Exception:
            # file not present or unreadable
            return {1: 0, 2: 0, 3: 0}

    # 'student_id' 칼럼이 있는지 확인합니다. 없으면 빈 결과를 돌려줍니다.
    if 'student_id' not in df.columns:
        return {1: 0, 2: 0, 3: 0}

    # student_id 문자열에서 학년(1/2/3)을 뽑는 함수입니다.
    # 여러가지 형식('12345', 'I25: 13645554' 등)을 다루도록 안전하게 만듭니다.
    def extract_grade(val):
        s = '' if pd.isna(val) else str(val)
        # 1) Prefer digits after a colon, e.g. 'I25: 13645554'
        m = re.search(r":\s*(\d+)", s)
        if m:
            digits = m.group(1)
            try:
                return int(digits[0])
            except Exception:
                return None

        # 2) Otherwise prefer the longest digit-run in the string
        all_digits = re.findall(r"(\d+)", s)
        if all_digits:
            longest = max(all_digits, key=len)
            try:
                return int(longest[0])
            except Exception:
                return None

        # 3) Fallback: first character as digit
        try:
            return int(s[0])
        except Exception:
            return None

    # 새 칼럼 'grade'를 만듭니다.
    df['grade'] = df['student_id'].apply(extract_grade)
    if df['grade'].isnull().all():
        return {1: 0, 2: 0, 3: 0}

    counts = df['grade'].value_counts().to_dict()
    return {g: int(counts.get(g, 0)) for g in [1, 2, 3]}


def background_thread():
    # 백그라운드에서 주기적으로 엑셀을 확인해 변경이 있으면 클라이언트에 보냅니다.
    last = {}
    while True:
        counts = read_and_count()
        if counts != last:
            socketio.emit('update_counts', counts)
            last = counts
        time.sleep(POLL_INTERVAL)


@app.route('/')
def index():
    return render_template('index.html')


@socketio.on('connect')
def handle_connect():
    # 클라이언트가 접속하면 현재 인원수를 바로 보냅니다.
    counts = read_and_count()
    socketio.emit('update_counts', counts)


if __name__ == '__main__':
    thread = threading.Thread(target=background_thread)
    thread.daemon = True
    thread.start()
    port = int(os.environ.get('PORT', '5000'))
    socketio.run(app, host='0.0.0.0', port=port)
