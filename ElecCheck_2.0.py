import os
import re
import math
import threading
import pandas as pd
import openpyxl
import pyautogui
import tkinter as tk
import tkinter.font as tkFont
import time
from tkinter import filedialog, messagebox

import ttkbootstrap as tb
from ttkbootstrap.constants import *

# ========== Selenium (Edge, 로컬 드라이버) ==========
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
from selenium.common.exceptions import TimeoutException

# ================================================
# 환경 설정
# ================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EDGE_DRIVER_PATH = os.path.join(BASE_DIR, "msedgedriver.exe")  # ← py랑 같은 폴더

desired_width = 800
desired_height = 1024

edge_options = EdgeOptions()
"""edge_options.add_argument("--headless") """ # 필요 없으면 주석
edge_options.add_argument("--new-tab")
edge_options.add_argument(f"--window-size={desired_width},{desired_height}")
edge_options.add_argument("--force-device-scale-factor=1")

driver = None  # 전역 driver
tree = None    # 전역 Treeview

# ================================================
# 드라이버 생성
# ================================================
def create_driver():
    if not os.path.exists(EDGE_DRIVER_PATH):
        raise FileNotFoundError(f"Edge 드라이버(msedgedriver.exe)를 찾을 수 없습니다: {EDGE_DRIVER_PATH}")
    service = EdgeService(executable_path=EDGE_DRIVER_PATH)
    d = webdriver.Edge(service=service, options=edge_options)
    return d

# ================================================
# 실행 버튼 누르면 쓰레드로 웹 작업
# ================================================
def run_web():
    answer = messagebox.askyesno("실행", "한전 원격지침 정보를 받아옵니다. 실행하겠습니까?", parent=root)
    if not answer:
        return
    t = threading.Thread(target=web_task)
    t.start()

# ================================================
# 실제 웹에서 긁어오는 작업
# ================================================
def web_task():
    global driver

    update_progress()

    # TreeView에 있던 초기값들 저장
    initial_values = []
    for item in tree.get_children():
        initial_values.append(tree.item(item, "values"))

    # 비교 전 테이블 비우기
    for item in tree.get_children():
        tree.delete(item)

    progress_log_thread("웹드라이버 실행")

    try:
        driver = create_driver()
    except Exception as e:
        progress_log_thread(f"웹드라이버 오류: {e}")
        reset_progress()
        return

    try:
        # 로그인
        driver.get("https://pp.kepco.co.kr/intro.do")
        progress_log_thread("로그인")

        driver.find_element(By.ID, "RSA_USER_ID").send_keys("pentjj")
        driver.find_element(By.ID, "RSA_USER_PWD").send_keys("Hyun@9539")
        ActionChains(driver).send_keys(Keys.RETURN).perform()

        # 첫 번째 고객번호 로그인 후 검침 화면으로
        progress_log_thread("검침화면 이동")
        driver.get("https://pp.kepco.co.kr/auth/register_after.do?CUSTNO=0526314773")
        driver.get("https://pp.kepco.co.kr/cc/cc0101.do?menu_id=O010207")

        # 이 고객번호는 계기 여러 개라 셀렉트 해야 하는 것들
        values_to_select = ["0526314773+06", "0526314773***", "0526314773+01", "0526314773+02","0526314773+03", "0526314773+04", "0526314773+05", "0526314773+07", "0526314773+08"]
        sheet_name_1 = ["설화명곡", "월배기지", "서부정류장", "반월당", "신천", "방촌", "안심", "숙천", "금락"]

        # 1호선 변전소
        for i in range(9):
            progress_log_thread(f"{sheet_name_1[i]} 페이지 로딩중...")

            # 테이블 로딩
            WebDriverWait(driver, 10**6).until(EC.text_to_be_present_in_element((By.ID, "jqgh_grid_VAR_NGT"), "지상"))
            time.sleep(1.0)
            
            # 계기 선택
            select_element = WebDriverWait(driver, 10**6).until(EC.presence_of_element_located((By.ID, "SEL_METER_ID")))
            Select(select_element).select_by_value(values_to_select[i])
            time.sleep(1.0)
            
            # 로딩 끝날 때까지
            WebDriverWait(driver, 10**6).until(EC.presence_of_element_located((By.XPATH,'//div[@id="backgroundLayer" and @class="loadingwrap" and @style="display: none;"]')))
            time.sleep(1.0)
            # 조회 버튼 클릭
            
            WebDriverWait(driver, 10**6).until(EC.presence_of_element_located((By.XPATH, '//*[@id="txt"]/div[2]/p/span[1]/a/img'))).click()            
            time.sleep(1.0)

            # 로딩 끝날 때까지
            WebDriverWait(driver, 10**6).until(EC.presence_of_element_located((By.XPATH,'//div[@id="backgroundLayer" and @class="loadingwrap" and @style="display: none;"]')))
            time.sleep(1.0)

            # 표 텍스트 파싱
            income_values = parse_table_to_values(driver)
            update_progress()

            # tree에 있던 initial_values와 비교
            insert_compare_rows(initial_values, i, income_values)

        # 2호선 변전소
        custnum_line2 = ["0530087761", "0530142327", "0530094940", "0530094888", "0530094851", "0530166621", "0530160011", "0530160020", "0530160039", "0537184143"]
        sheet_name_2 = ["문양기지", "대실", "성서산단", "죽전", "반고개", "대구은행", "만촌", "대공원", "사월", "영남대"]

        for offset, (custno, name) in enumerate(zip(custnum_line2, sheet_name_2), start=9):
            progress_log_thread(f"{name} 페이지 로딩중...")

            driver.get(f"https://pp.kepco.co.kr/auth/register_after.do?CUSTNO={custno}")
            driver.get("https://pp.kepco.co.kr/cc/cc0101.do?menu_id=O010207")

            WebDriverWait(driver, 30).until(EC.text_to_be_present_in_element((By.ID, "jqgh_grid_VAR_NGT"), "지상"))
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH,'//div[@id="backgroundLayer" and @class="loadingwrap" and @style="display: none;"]')))

            income_values = parse_table_to_values(driver)
            update_progress()

            insert_compare_rows(initial_values, offset, income_values)

        progress_log_thread("완료")
        reset_progress()
        pyautogui.alert("완료")

    finally:
        if driver:
            driver.quit()


# ================================================
# 테이블 파싱 → 우리가 비교에 쓸 8개 값으로 정리
# ================================================
def parse_table_to_values(driver):
    element = driver.find_element(By.ID, "gview_grid")
    all_text = element.text

    # "진상" 이후부터만 쓰는 기존 로직 유지
    target_string = "진상"
    start_index = all_text.find(target_string)
    if start_index != -1:
        data_text = all_text[start_index + len(target_string):]
        next_line_index = data_text.find('\n') + 1
        data_text = data_text[next_line_index:]
    else:
        data_text = all_text

    # 줄 → 칸 분리
    data_rows = data_text.split('\n')
    data_columns = [row.split() for row in data_rows if row.strip()]
    df = pd.DataFrame(data_columns)

    # ===== 여기부터 경고 없애는 부분 =====
    # 1) 콤마 제거 (문자열 칼럼만)
    for col in df.columns:
        if df[col].dtype == object:
            # 콤마만 제거
            df[col] = df[col].str.replace(',', '', regex=False)

    # 2) 숫자로 변환 가능한 건 숫자로
    for col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
    # ===== 여기까지 =====

    # 안전하게 값 꺼내는 헬퍼
    def get_val(r, c):
        try:
            val = df.iloc[r, c]
            # NaN이면 빈 문자열로
            if pd.isna(val):
                return ''
            return val
        except Exception:
            return ''

    # 원래 패턴 유지 (8개)
    income_values = [
        '',
        get_val(0, 3),  # 9
        get_val(0, 4),  # 10
        get_val(0, 5),  # 11
        get_val(0, 8),  # 12
        '',
        get_val(0, 6),  # 14
        get_val(0, 7),  # 15
    ]
    return income_values

# ================================================
# 기존값 vs 원격값 비교해서 TreeView에 두 줄씩 넣기
# ================================================
def insert_compare_rows(initial_values, idx, income_values):
    # initial_values[idx] 가 엑셀에 있던 그 값
    # income_values         가 원격사이트에서 긁어온 값
    tolerance = 1e-9

    try:
        cond = (
            float(initial_values[idx][1]) == float(income_values[1]) and
            float(initial_values[idx][2]) == float(income_values[2]) and
            float(initial_values[idx][3]) == float(income_values[3]) and
            float(initial_values[idx][6]) == float(income_values[6]) and
            float(initial_values[idx][7]) == float(income_values[7]) and
            math.isclose(
                float(initial_values[idx][4]) + float(initial_values[idx][5]),
                float(income_values[4]),
                rel_tol=tolerance
            )
        )
    except Exception:
        cond = False

    if cond:
        tag = f'row_ok_{idx}'
        tree.tag_configure(tag, background='#B7DBF7', foreground='#0E2E44')
    else:
        tag = f'row_bad_{idx}'
        tree.tag_configure(tag, background='#F7C8B0', foreground='#5A1E00')

    tree.insert('', tk.END, values=initial_values[idx], tags=(tag,))
    tree.insert('', tk.END, values=income_values, tags=(tag,))


# ================================================
# TreeView 생성 (엑셀 내용 보여줄 곳)
# ================================================
def create_table(root):
    global tree
    frame = tb.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)

    tree = tb.Treeview(
        frame,
        columns=('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'),
        show='headings',
        bootstyle=PRIMARY
    )
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tb.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    tree.configure(yscrollcommand=scrollbar.set)

    tree.heading('A', text='변전소', anchor='center')
    tree.heading('B', text='9', anchor='center')
    tree.heading('C', text='10', anchor='center')
    tree.heading('D', text='11', anchor='center')
    tree.heading('E', text='12', anchor='center')
    tree.heading('F', text='13', anchor='center')
    tree.heading('G', text='14', anchor='center')
    tree.heading('H', text='15', anchor='center')

    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        tree.column(col, width=70, stretch=True, anchor='center')

# ================================================
# 엑셀 열기 → 시트 선택 → 지정셀 읽어서 TreeView에 넣기
# ================================================
def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet_names = workbook.sheetnames

        root.attributes('-disabled', True)

        sheet_window = tb.Toplevel(root)
        sheet_window.title("시트 선택")
        sheet_window.geometry("200x220")

        def on_close():
            root.attributes('-disabled', False)
            sheet_window.destroy()

        sheet_window.protocol("WM_DELETE_WINDOW", on_close)

        listbox = tk.Listbox(sheet_window)
        listbox.pack(fill=tk.BOTH, expand=True)

        for name in sheet_names:
            listbox.insert(tk.END, name)

        def load_sheet():
            sel = listbox.curselection()
            if sel:
                sname = listbox.get(sel[0])
                on_sheet_select(sname, workbook)
                on_close()
                enable_run_button()

        tb.Button(sheet_window, text="시트 열기", command=load_sheet, bootstyle=SUCCESS).pack(pady=5)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")

# ================================================
# 시트 선택 후 셀 범위별로 값 읽어서 트리에 넣기
# (원래 네가 올린 셀 목록 그대로 가져옴)
# ================================================
def on_sheet_select(sheet_name, workbook):
    sheet = workbook[sheet_name]

    # 기존 내용 비우기
    for item in tree.get_children():
        tree.delete(item)

    cell_ranges = [
        ['B4', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12'], #설화명곡
        ['F4', 'G6', 'G7', 'G8', 'G9', 'G10', 'G11', 'G12'], #월배기지
        ['J4', 'K6', 'K7', 'K8', 'K9', 'K10', 'K11', 'K12'], #서부정류장
        
        ['B17', 'C19', 'C20', 'C21', 'C22', 'C23', 'C24', 'C25'], #반월당
        ['F17', 'G19', 'G20', 'G21', 'G22', 'G23', 'G24', 'G25'], #신천
        ['J17', 'K19', 'K20', 'K21', 'K22', 'K23', 'K24', 'K25'], #방촌
        
        ['B30', 'C32', 'C33', 'C34', 'C35', 'C36', 'C37', 'C38'], #안심
        ['F30', 'G32', 'G33', 'G34', 'G35', 'G36', 'G37', 'G38'], #숙천
        ['J30', 'K32', 'K33', 'K34', 'K35', 'K36', 'K37', 'K38'], #금락
        
        ['B51', 'C53', 'C54', 'C55', 'C56', 'C57', 'C58', 'C59'], #문양기지
        ['F51', 'G53', 'G54', 'G55', 'G56', 'G57', 'G58', 'G59'], #대실
        ['J51', 'K53', 'K54', 'K55', 'K56', 'K57', 'K58', 'K59'], #성서산단
        
        ['B64', 'C66', 'C67', 'C68', 'C69', 'C70', 'C71', 'C72'], #죽전
        ['F64', 'G66', 'G67', 'G68', 'G69', 'G70', 'G71', 'G72'], #반고개
        ['J64', 'K66', 'K67', 'K68', 'K69', 'K70', 'K71', 'K72'], #대구은행
        
        ['B77', 'C79', 'C80', 'C81', 'C82', 'C83', 'C84', 'C85'], #만촌
        ['F77', 'G79', 'G80', 'G81', 'G82', 'G83', 'G84', 'G85'], #대공원
        ['J77', 'K79', 'K80', 'K81', 'K82', 'K83', 'K84', 'K85'], #사월
        
        ['B90', 'C92', 'C93', 'C94', 'C95', 'C96', 'C97', 'C98'], #만촌        
        
       
    ]
    cell_values = [
        '설화명곡', '월배기지', '서부정류장', '반월당', '신천', '방촌', '안심', '숙천', '금락',
        '문양기지', '대실', '성서산단', '죽전', '반고개', '대구은행', '만촌', '대공원', '사월', '영남대'
    ]

    for cell_range, first_col in zip(cell_ranges, cell_values):
        row_vals = []
        for idx, cell_addr in enumerate(cell_range):
            if idx == 0:
                row_vals.append(first_col)
            else:
                row_vals.append(sheet[cell_addr].value)
        tree.insert('', tk.END, values=row_vals)

    root.geometry("700x450")


# ================================================
# 진행로그 / 프로그레스바
# ================================================
def progress_log(value):
    progresslog.config(text=f"{value}")

def progress_log_thread(value):
    threading.Thread(target=progress_log, args=(value,)).start()

def update_progress():
    progress.step(5)

def reset_progress():
    progress['value'] = 0

def enable_run_button():
    menubar.entryconfig("실행", state="normal")


# ================================================
# 메인 윈도우
# ================================================
root = tb.Window(themename="minty")
root.title("한전 원격검침 비교")
root.geometry("650x150")

# 나눔고딕 기본 적용
default_font = tkFont.nametofont("TkDefaultFont")
default_font.configure(family="NanumGothic", size=10)
tkFont.nametofont("TkTextFont").configure(family="NanumGothic", size=10)
tkFont.nametofont("TkFixedFont").configure(family="NanumGothic", size=10)

status_frame = tb.Frame(root)
status_frame.pack(side=tk.BOTTOM, fill=tk.X, anchor="center")

progress = tb.Progressbar(status_frame, orient='horizontal', mode='determinate', bootstyle=SUCCESS)
progress.pack(fill=tk.BOTH, anchor="center")

progresslog = tb.Label(status_frame, text="파일 - 열기 - 검침엑셀파일 선택 - 비교시트 열기 - 실행 누르기", justify='left')
progresslog.pack(anchor='w')

# 메뉴
menubar = tk.Menu(root)
file_menu = tk.Menu(menubar, tearoff=0)
file_menu.add_command(label="열기", command=open_file)
file_menu.add_separator()
file_menu.add_command(label="종료", command=root.quit)
menubar.add_cascade(label="파일", menu=file_menu)
menubar.add_cascade(label="실행", command=run_web, state="disabled")
help_menu = tk.Menu(menubar, tearoff=0)
help_menu.add_command(label="도움말")
menubar.add_cascade(label="About", menu=help_menu)
root.config(menu=menubar)

# 테이블 만들기
create_table(root)

root.mainloop()
