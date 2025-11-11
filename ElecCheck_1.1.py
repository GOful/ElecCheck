import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.keys import Keys
import tkinter as tk
import pyautogui
import threading
import pandas as pd
import re 
import math
import time
from webdriver_manager.core.os_manager import ChromeType

desired_width = 800
desired_height = 1024
options = Options()
options.add_argument("--headless")
options.add_argument('--new-tab')
options.add_argument(f'--window-size={desired_width},{desired_height}')
options.add_experimental_option("detach", True)

def run_web():
    answer = messagebox.askyesno("실행", "한전 원격지침 정보를 받아옵니다. 실행하겠습니까?", parent=root)
    if not answer:
        return
    t = threading.Thread(target=web_task)
    t.start()

def web_task():
    update_progress()  
    initial_values=[]
    for item in tree.get_children():        
        initial_values.append(tree.item(item, 'values'))
    for item in tree.get_children():
        tree.delete(item)   
    progress_log_thread('웹드라이버 실행중')
    global driver
    try:
        # 캐시 강제 초기화 + 최신 버전 사용
        service = Service(
            ChromeDriverManager(
                version='114.0.5735.90',  # 명시적 버전 지정
                chrome_type=ChromeType.GOOGLE,
                cache_valid_range=0  # 캐시 무효화
            ).install()
        )
        
        # 헤드리스 모드 정확한 창 설정
        options.add_argument(f"--window-size={desired_width},{desired_height}")
        options.add_argument("--force-device-scale-factor=1")
        
        global driver
        driver = webdriver.Chrome(service=service, options=options)   
    
    
    driver = webdriver.Chrome(service=service, options=options)
    driver.get("https://pp.kepco.co.kr/intro.do")    
    progress_log_thread('로그인중')
    driver.find_element(By.XPATH, '//*[@id="RSA_USER_ID"]').send_keys("1")
    driver.find_element(By.XPATH, '//*[@id="RSA_USER_PWD"]').send_keys("2")
    webdriver.ActionChains(driver).send_keys(Keys.RETURN).perform()    
    driver.get("https://pp.kepco.co.kr/auth/register_after.do?CUSTNO=0526314773")
    driver.get("https://pp.kepco.co.kr/cc/cc0101.do?menu_id=O010207")
    values_to_select = ["0526314773+06", "0526314773***", "0526314773+01", "0526314773+02", "0526314773+07", "0526314773+04", "0526314773+05"]
    dfs = []
    income_values = []
    sheet_name=["설화명곡", "월배기지", "서부정류장", "반월당", "신천", "방촌", "안심"]
    for i in range(7): 
        progress_log_thread(f'{sheet_name[i]} 페이지 로딩중...')
        WebDriverWait(driver, 1000).until(EC.text_to_be_present_in_element((By.ID, 'jqgh_grid_VAR_NGT'), '지상'))
        select_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "SEL_METER_ID")))
        select = Select(select_element)
        select.select_by_value(values_to_select[i])
        WebDriverWait(driver, 1000).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txt"]/div[2]/p/span[1]/a/img'))).click()
        WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, '//div[@id="backgroundLayer" and @class="loadingwrap" and @style="display: none;"]')))
        target_string = "진상"
        element = driver.find_element(By.XPATH, '//*[@id="gview_grid"]')
        all_text = element.text
        start_index = all_text.find(target_string)
        if start_index != -1:
            data_text = all_text[start_index + len(target_string):]
            next_line_index = data_text.find('\n') + 1
            data_text = data_text[next_line_index:]
        else:
            data_text = all_text
        data_rows = data_text.split('\n')
        data_columns = [row.split() for row in data_rows]
        df = pd.DataFrame(data_columns)
        dfs.append(df)
        update_progress()     
        df = df.applymap(lambda x: re.sub(r',', '', x) if isinstance(x, str) else x)
        df = df.apply(pd.to_numeric, errors='ignore')  # 숫자로 변환합니다.
        income_values = ['', df.iloc[0, 3], df.iloc[0, 4], df.iloc[0, 5], df.iloc[0, 8], '', df.iloc[0, 6], df.iloc[0, 7]]
        print(income_values)
        tolerance = 1e-9 
        if (float(initial_values[i][1]) == float(income_values[1])) and \
        (float(initial_values[i][2]) == float(income_values[2])) and \
        (float(initial_values[i][3]) == float(income_values[3])) and \
        (float(initial_values[i][6]) == float(income_values[6])) and \
        (float(initial_values[i][7]) == float(income_values[7])) and \
        (math.isclose(float(initial_values[i][4]) + float(initial_values[i][5]), float(income_values[4]), rel_tol=tolerance)):  
            tree.tag_configure(f'column_tag{i}', background='blue', foreground='white')
            tree.insert('', tk.END, values=initial_values[i], tags=(f'column_tag{i}',)) 
            tree.insert('', tk.END, values=income_values, tags=(f'column_tag{i}',)) 
        else:
            tree.tag_configure(f'column_tag{i}', background='red', foreground='yellow')
            tree.insert('', tk.END, values=initial_values[i], tags=(f'column_tag{i}',)) 
            tree.insert('', tk.END, values=income_values, tags=(f'column_tag{i}',))

    custnum_line2 = ["0530087761", "0530142327", "0530094940", "0530094888", "0530094851", "0530166621", "0530160011", "0530160020", "0530160039", "0537184143" ]
    dfs2 = []
    income_values = []
    sheet_name2=["문양기지", "대실", "성서산단", "죽전", "반고개", "대구은행", "만촌", "대공원", "사월", "영남대"]
    for i, j in zip(range(7, 17), range(10)):
        progress_log_thread(f'{sheet_name2[j]} 페이지 로딩중')
        driver.get(f"https://pp.kepco.co.kr/auth/register_after.do?CUSTNO={custnum_line2[j]}")        
        driver.get("https://pp.kepco.co.kr/cc/cc0101.do?menu_id=O010207")
        WebDriverWait(driver, 1000).until(EC.text_to_be_present_in_element((By.ID, 'jqgh_grid_VAR_NGT'), '지상'))
        WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, '//div[@id="backgroundLayer" and @class="loadingwrap" and @style="display: none;"]')))
        target_string = "진상"
        element = driver.find_element(By.XPATH, '//*[@id="gview_grid"]')
        all_text = element.text
        start_index = all_text.find(target_string)
        if start_index != -1:
            data_text = all_text[start_index + len(target_string):]
            next_line_index = data_text.find('\n') + 1
            data_text = data_text[next_line_index:]
        else:
            data_text = all_text
        data_rows = data_text.split('\n')
        data_columns = [row.split() for row in data_rows]
        df = pd.DataFrame(data_columns)
        dfs2.append(df)
        update_progress()
        df = df.applymap(lambda x: re.sub(r',', '', x) if isinstance(x, str) else x)
        df = df.apply(pd.to_numeric, errors='ignore')  # 숫자로 변환합니다.
        income_values = ['', df.iloc[0, 3], df.iloc[0, 4], df.iloc[0, 5], df.iloc[0, 8], '', df.iloc[0, 6], df.iloc[0, 7]]
        print(income_values)        
        tolerance = 1e-9 
        if (float(initial_values[i][1]) == float(income_values[1])) and \
        (float(initial_values[i][2]) == float(income_values[2])) and \
        (float(initial_values[i][3]) == float(income_values[3])) and \
        (float(initial_values[i][6]) == float(income_values[6])) and \
        (float(initial_values[i][7]) == float(income_values[7])) and \
        (math.isclose(float(initial_values[i][4]) + float(initial_values[i][5]), float(income_values[4]), rel_tol=tolerance)):  
            tree.tag_configure(f'column_tag{i}', background='blue', foreground='white')
            tree.insert('', tk.END, values=initial_values[i], tags=(f'column_tag{i}',)) 
            tree.insert('', tk.END, values=income_values, tags=(f'column_tag{i}',)) 
        else:
            tree.tag_configure(f'column_tag{i}', background='red', foreground='yellow')
            tree.insert('', tk.END, values=initial_values[i], tags=(f'column_tag{i}',)) 
            tree.insert('', tk.END, values=income_values, tags=(f'column_tag{i}',))
    progress_log_thread('완료')
    reset_progress()
    pyautogui.alert('완료')
    driver.quit()

def create_table(root):
    global tree    
    frame = tk.Frame(root)
    frame.pack(fill=tk.BOTH, expand=True)
    tree = ttk.Treeview(frame, columns=('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'), show='headings')
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
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
    for column in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
        tree.column(column, width=70, stretch=True, anchor='center')

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if not file_path:
        return
    try:
        workbook = openpyxl.load_workbook(file_path, read_only=True)
        sheet_names = workbook.sheetnames

        root.attributes('-disabled', True)

        # Get the position and size of the root window
        root.update_idletasks()  # Ensure window size is updated
        root_width = root.winfo_width()
        root_height = root.winfo_height()
        root_x = root.winfo_rootx()
        root_y = root.winfo_rooty()

        # Calculate the position for the new window
        new_x = root_x + root_width // 2 - 100  # Adjust as needed
        new_y = root_y + root_height // 2 - 100  # Adjust as needed

        # Show sheet names in a new window with a listbox
        sheet_window = tk.Toplevel(root)
        sheet_window.title("Select Sheet")

        def on_close():
            root.attributes('-disabled', False)
            sheet_window.destroy()

        sheet_window.protocol("WM_DELETE_WINDOW", on_close)

        sheet_window.geometry(f"200x200+{new_x}+{new_y}")  # Set window size and position

        # Create a frame for the listbox and scrollbar
        list_frame = tk.Frame(sheet_window)
        list_frame.pack(fill=tk.BOTH, expand=True)

        # Create the listbox with a scrollbar
        listbox = tk.Listbox(list_frame)
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)

        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        for sheet_name in sheet_names:
            listbox.insert(tk.END, sheet_name)

        def load_sheet():
            selection = listbox.curselection()
            if selection:
                sheet_name = listbox.get(selection[0])
                on_sheet_select(sheet_name, workbook)
                on_close()
                enable_run_button()                
                
                

        button = tk.Button(sheet_window, text="시트 열기", command=load_sheet)
        button.pack(pady=5)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")
        
running = True

def progress_log(value):    
    progresslog.config(text=f"{value}....")
    
def progress_log_thread(value):
    global running
    running = True
    thread = threading.Thread(target=progress_log, args=(value,))
    thread.start()
    # progress_log() 함수가 실행되면서 작업을 처리하는 동안에는 running을 True로 유지한다.
    # progress_log_thread() 함수 호출이 끝나면 running을 False로 변경하여 progress_log() 함수가 더 이상 실행되지 않도록 한다.
    running = False



def on_sheet_select(sheet_name, workbook):
    sheet = workbook[sheet_name]


    
    # Clear existing data in the Treeview
    for item in tree.get_children():
        tree.delete(item)
    
    cell_ranges = [
        ['B4', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12'],
        ['F4', 'G6', 'G7', 'G8', 'G9', 'G10', 'G11', 'G12'],
        ['J4', 'K6', 'K7', 'K8', 'K9', 'K10', 'K11', 'K12'],

        ['B17', 'C19', 'C20', 'C21', 'C22', 'C23', 'C24', 'C25'],
        ['F17', 'G19', 'G20', 'G21', 'G22', 'G23', 'G24', 'G25'],
        ['J17', 'K19', 'K20', 'K21', 'K22', 'K23', 'K24', 'K25'],

        ['B30', 'C32', 'C33', 'C34', 'C35', 'C36', 'C37', 'C38'],


        ['B44', 'C46', 'C47', 'C48', 'C49', 'C50', 'C51', 'C52'],
        ['F44', 'G46', 'G47', 'G48', 'G49', 'G50', 'G51', 'G52'],
        ['J44', 'K46', 'K47', 'K48', 'K49', 'K50', 'K51', 'K52'],

        ['B57', 'C59', 'C60', 'C61', 'C62', 'C63', 'C64', 'C65'],
        ['F57', 'G59', 'G60', 'G61', 'G62', 'G63', 'G64', 'G65'],
        ['J57', 'K59', 'K60', 'K61', 'K62', 'K63', 'K64', 'K65'],

        ['B70', 'C72', 'C73', 'C74', 'C75', 'C76', 'C77', 'C78'],
        ['F70', 'G72', 'G73', 'G74', 'G75', 'G76', 'G77', 'G78'],
        ['J70', 'K72', 'K73', 'K74', 'K75', 'K76', 'K77', 'K78'],

        ['B83', 'C85', 'C86', 'C87', 'C88', 'C89', 'C90', 'C91'],
    ]
    cell_values = ['설화명곡', '월배기지', '서부정류장', '반월당', '신천', '방촌', '안심', '문양기지', '대실', '성서산단', '죽전', '반고개', '대구은행', '만촌', '대공원', '사월', '영남대']
    stored_values = [] 
    # Insert data into Treeview
    for cell_range, cell_value in zip(cell_ranges, cell_values):
        cell_values_list = [cell_value if idx == 0 else sheet[cell].value for idx, cell in enumerate(cell_range)]
        stored_values.append(cell_values_list)
        tree.insert('', tk.END, values=cell_values_list)
    
    root.geometry("600x400")

def update_progress():
    progress.step(5)

def reset_progress():
    progress['value'] = 0

def enable_run_button():
    menubar.entryconfig("실행", state="normal")    

# Create the main window
root = tk.Tk()
root.title("한전 원격검침 비교")
root.geometry("600x100")
status_frame = tk.Frame(root)
status_frame.pack(side=tk.BOTTOM, fill=tk.X,anchor="center")
progress = ttk.Progressbar(status_frame, orient='horizontal', mode='determinate')
progress.pack(fill=tk.BOTH, anchor="center")


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

progresslog = tk.Label(status_frame, text="파일 - 열기 - 검침엑셀파일 선택 - 비교시트 열기 - 실행 누르기", justify='left')
progresslog.pack(anchor='w')


# Create the menu

# Create the table
create_table(root)
root.mainloop()
driver.quit()
