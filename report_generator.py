import os
import subprocess
from datetime import datetime
import time
import ctypes
import win32com.client
import psutil


def active_excel(cpu_threshold=7.0):
    """Check if Excel is actively processing and consuming CPU.
    """
    for proc in psutil.process_iter(['name', 'cpu_percent']):
        if proc.info['name'] and 'EXCEL' in proc.info['name'].upper():
            try:
                cpu = proc.cpu_percent(interval=1)
                print(f"Uso de CPU do Excel: {cpu}%")
                if cpu > cpu_threshold:
                    return True
            except psutil.NoSuchProcess:
                continue
    return False


def wait_excel_stop_updates():
    """Waiting for Excel to finish updating spreadsheets.
    """
    print("⏳ Waiting for Excel...")
    inactive_counter = 0
    last_move = time.time()
    while inactive_counter < 5:
        if active_excel():
            print("Excel processing...")
            inactive_counter = 0
        else:
            print("Excel inactive...")
            inactive_counter += 1
        # Cursor movement to prevent screensaver or sleep mode
        if time.time() - last_move > 240:
            ctypes.windll.user32.mouse_event(0x0001, 0, 1, 0, 0)
            ctypes.windll.user32.mouse_event(0x0001, 0, -1, 0, 0)
            last_move = time.time()
        time.sleep(5)
    print("✅ Excel finished updating.")


DAY = datetime.now().strftime('%d')
YEAR = datetime.now().year
MONTH = datetime.now().strftime('%m')

arq = {  # 'file location':['Spreadsheet pages to export'],
    'file location': [2, 3, 4, 5, 6, 7],
    'file location': [1, 2],
    'tLocal do arquivo': [1],
    'file location': [1],
    'file location': [1],
    'file location': [1, 2, 3],
    'file location': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
    'file location': [1, 2, 3, 4]
}

os.system("taskkill /f /im excel.exe")

excel = win32com.client.DispatchEx("Excel.Application")
excel.Visible = False

workbooks = []

for path in arq.keys():
    print(f"📂 Trying to open: {path}")
    ATTEMPT = 1
    SUCCESS = False
    while ATTEMPT <= 3 and not SUCCESS:
        try:
            if "Heavy A" in path or "Heavy B" in path:  # These spreadsheets take longer to open
                time.sleep(15)
            else:
                time.sleep(6)
            wb = excel.Workbooks.Open(path)
            workbooks.append(wb)
            print(f"✅ Success open: {path} (attempt {ATTEMPT})")
            SUCCESS = True
        except Exception as e:
            print(f"❌ Attempt {ATTEMPT} failed to open {path}: {e}")
            ATTEMPT += 1
            time.sleep(10)
    if not SUCCESS:
        print(f"⚠️ Couldn't open {path} after 3 tries.")
# Spreadsheets get auto updated when opened, we just have to wait for them to finish
wait_excel_stop_updates()
# Closing files after update
for wb in workbooks:
    try:
        wb.Close(False)
    except:
        pass
workbooks.clear()
excel.Quit()
del excel
time.sleep(5)
TODAY = datetime.now().strftime('%Y_%m_%d')
TODAY_FOLDER = f'exported pdf folder path\\{TODAY}'
for file, index in arq.items():
    WB_PATH = file
    ws_index_list = index
    nome = file.split('/')[-1][:-5]
    DIR_END = f'{TODAY_FOLDER}\\{nome}'
    FDATE = f'_{YEAR}_{MONTH}_{DAY}'
    PATH_TO_PDF = DIR_END+FDATE+'.pdf'
    excel = win32com.client.Dispatch("Excel.Application")
    try:
        excel.Visible = False
        print(f"📄 Exporting {nome}...")
        wb = excel.Workbooks.Open(WB_PATH)
        wb.Worksheets(ws_index_list).Select()
        wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
        time.sleep(5)
    except Exception as e:
        print(f"❌ Fail to export {nome}: {e}")
        time.sleep(3)
    else:
        print('Success.')
    finally:
        try:
            wb.Close(False)
            time.sleep(3)
        except Exception as e:
            print('Fail.')
            print(e)
            os.system("taskkill /f /im excel.exe")
            time.sleep(3)
graphic_path = 'path to other script\\graphic.py'
subprocess.run(['python', graphic_path], check=True)
