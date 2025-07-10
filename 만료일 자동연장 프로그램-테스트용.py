import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, Menu
import pyodbc
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from dateutil.relativedelta import relativedelta
import re
import json
import os
import threading
# --- 업데이트 기능에 필요한 라이브러리 ---
import requests
import sys
import subprocess

class ExpiryExtensionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("만료일 자동연장 프로그램")
        self.root.geometry("600x550")
        
        # ========================================================
        # ======== 업데이트 기능 관련 설정 (신규) START ========
        # ========================================================
        # 현재 프로그램 버전 (새 버전 배포 시 이 번호를 올려주세요)
        self.CURRENT_VERSION = "1.0" 
        # 고객님의 GitHub 저장소 주소 (예: "username/repository-name")
        self.GITHUB_REPO = "YOUR_USERNAME/YOUR_REPOSITORY_NAME" 
        # ========================================================
        # ========= 업데이트 기능 관련 설정 (신규) END =========
        # ========================================================
        
        # --- 기본 설정값 ---
        self.default_server = "116.122.36.35,1433"
        self.default_db = "UBIPLUS"
        self.default_user = "ubi_db_admin"
        self.default_password = "d8*PyaEdA%6gP$GrGx#SNgJidQGzb"
        
        # --- 구글 시트 설정 ---
        self.sheet_id = "1uZWnq_qPXH9JHIMiRlcaeMZ9bubkC5KKBXwsZhswSfQ"
        self.sheet_name = "자동 연장하기"
        
        self.config_file = "config.json"

        # --- 화면 전환을 위한 프레임 설정 ---
        container = tk.Frame(root)
        container.pack(side="top", fill="both", expand=True)
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        self.frames = {}
        for F in (ExtensionPage, CalculatorPage):
            page_name = F.__name__
            frame = F(parent=container, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky="nsew")

        # --- 초기 UI 설정 ---
        self.setup_menu()
        self.show_frame("ExtensionPage") 

        # --- 단축키 및 프로토콜 설정 ---
        self.root.bind("<F5>", self.process_sms)
        self.root.bind("<F4>", self.clear_and_paste)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # --- 프로그램 시작 시 업데이트 확인 ---
        self.root.title(f"만료일 자동연장 프로그램 v{self.CURRENT_VERSION}")
        threading.Thread(target=self.check_for_updates, daemon=True).start()

    # ========================================================
    # ======== 업데이트 기능 관련 함수 (신규) START ========
    # ========================================================
    def check_for_updates(self):
        """GitHub API를 통해 최신 릴리즈 정보를 확인하고 업데이트를 제안합니다."""
        try:
            # GitHub 저장소 주소를 정확히 입력해주세요.
            if self.GITHUB_REPO == "YOUR_USERNAME/YOUR_REPOSITORY_NAME":
                print("업데이트: GitHub 저장소 정보가 설정되지 않았습니다.")
                return

            api_url = f"https://api.github.com/repos/{self.GITHUB_REPO}/releases/latest"
            response = requests.get(api_url)
            response.raise_for_status() # 오류가 있으면 예외 발생
            
            latest_release = response.json()
            latest_version = latest_release["tag_name"].lstrip('v') # 'v1.1' -> '1.1'
            
            # 버전 비교 (간단한 문자열 비교, 필요시 더 정교한 비교 로직 사용 가능)
            if latest_version > self.CURRENT_VERSION:
                if messagebox.askyesno("업데이트 알림", f"새로운 버전(v{latest_version})이 있습니다. 지금 업데이트하시겠습니까?"):
                    asset = latest_release["assets"][0]
                    download_url = asset["browser_download_url"]
                    filename = asset["name"]
                    
                    # 다운로드 및 업데이트는 GUI가 멈추지 않도록 새 스레드에서 실행
                    threading.Thread(target=self.download_and_apply_update, args=(download_url, filename), daemon=True).start()
    
        except requests.exceptions.RequestException as e:
            print(f"업데이트 확인 중 네트워크 오류 발생: {e}")
        except Exception as e:
            print(f"업데이트 확인 중 오류 발생: {e}")

    def download_and_apply_update(self, url, new_filename):
        """새 버전 파일을 다운로드하고, 교체 스크립트를 실행합니다."""
        try:
            # 메인 페이지의 상태 라벨에 진행 상황 표시
            status_label = self.frames["ExtensionPage"].status_label
            status_label.config(text=f"업데이트 다운로드 중... 0%", foreground="blue")
            self.root.update_idletasks()

            response = requests.get(url, stream=True)
            response.raise_for_status()

            # 다운로드할 파일 경로 설정 (기존 파일명에 '_new' 추가)
            base_name, ext = os.path.splitext(os.path.basename(sys.executable))
            download_path = f"{base_name}_new{ext}"
            
            total_size = int(response.headers.get('content-length', 0))
            bytes_downloaded = 0
            
            with open(download_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
                    bytes_downloaded += len(chunk)
                    progress = int((bytes_downloaded / total_size) * 100) if total_size > 0 else 0
                    status_label.config(text=f"업데이트 다운로드 중... {progress}%")
                    self.root.update_idletasks()

            status_label.config(text="다운로드 완료. 프로그램을 재시작합니다.", foreground="green")

            # 현재 실행 파일의 전체 경로
            current_exe_path = sys.executable
            # 업데이트 스크립트(updater.bat) 생성
            updater_script_path = "updater.bat"
            with open(updater_script_path, "w") as f:
                f.write(f"@echo off\n")
                f.write(f"echo 프로그램을 업데이트하는 중입니다. 잠시만 기다려주세요...\n")
                f.write(f"timeout /t 2 /nobreak > nul\n") # 기존 프로그램이 완전히 종료될 때까지 대기
                f.write(f"move /y \"{download_path}\" \"{current_exe_path}\"\n") # 새 파일로 덮어쓰기
                f.write(f"start \"\" \"{current_exe_path}\"\n") # 새 버전 실행
                f.write(f"del \"{updater_script_path}\"\n") # 업데이터 자신 삭제

            # .bat 파일 실행하고 현재 프로그램 종료
            subprocess.Popen(updater_script_path, shell=True)
            sys.exit()

        except Exception as e:
            messagebox.showerror("업데이트 실패", f"업데이트 중 오류가 발생했습니다: {e}")
            status_label.config(text="업데이트 실패", foreground="red")
    # ========================================================
    # ========= 업데이트 기능 관련 함수 (신규) END =========
    # ========================================================
    
    def show_frame(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()

    def setup_menu(self):
        menubar = Menu(self.root)
        self.root.config(menu=menubar)

        menu_main = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="MENU", menu=menu_main)
        
        menu_main.add_command(label="연장하기", command=lambda: self.show_frame("ExtensionPage"))
        menu_main.add_command(label="연장 계산하기", command=lambda: self.show_frame("CalculatorPage"))
    
    def process_sms(self, event=None):
        self.frames["ExtensionPage"].process_sms(event)

    def clear_and_paste(self, event=None):
        self.frames["ExtensionPage"].clear_and_paste(event)

    def load_config(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    return config.get('last_writer', "송민규")
        except (json.JSONDecodeError, FileNotFoundError):
            pass
        return "송민규"

    def save_config(self):
        writer = self.frames["ExtensionPage"].writer_combobox.get()
        config = {'last_writer': writer}
        with open(self.config_file, 'w') as f:
            json.dump(config, f)
            
    def on_closing(self):
        self.save_config()
        self.root.destroy()

# --- '연장 계산하기' 페이지 ---
class CalculatorPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        
        label = ttk.Label(self, text="여기는 '연장 계산하기' 화면입니다.\n이곳에 새로운 기능을 추가할 수 있습니다.", font=("맑은 고딕", 12))
        label.pack(side="top", fill="both", expand=True, padx=20, pady=20)

# --- '연장하기' 페이지 ---
class ExtensionPage(tk.Frame):
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.setup_ui()

    def setup_ui(self):
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        db_frame = ttk.LabelFrame(main_frame, text="데이터베이스 설정", padding="10")
        db_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(db_frame, text="서버:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.server_entry = ttk.Entry(db_frame, width=50)
        self.server_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=2)
        self.server_entry.insert(0, self.controller.default_server)
        
        ttk.Label(db_frame, text="DB:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.db_entry = ttk.Entry(db_frame, width=50)
        self.db_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=2)
        self.db_entry.insert(0, self.controller.default_db)
        
        ttk.Label(db_frame, text="DB계정:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.user_entry = ttk.Entry(db_frame, width=50)
        self.user_entry.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=2)
        self.user_entry.insert(0, self.controller.default_user)
        
        ttk.Label(db_frame, text="DB계정 암호:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.password_entry = ttk.Entry(db_frame, width=50, show="*")
        self.password_entry.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=2)
        self.password_entry.insert(0, self.controller.default_password)
        
        memo_frame = ttk.LabelFrame(main_frame, text="상담 내역 설정", padding="10")
        memo_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        writer_list = ["변은경", "홍서하", "우정영", "송민규"]
        last_writer = self.controller.load_config()
        
        ttk.Label(memo_frame, text="작성자:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.writer_combobox = ttk.Combobox(memo_frame, values=writer_list, width=18, state="readonly")
        self.writer_combobox.grid(row=0, column=1, sticky=tk.W, pady=2)
        
        if last_writer in writer_list:
            self.writer_combobox.set(last_writer)
        else:
            self.writer_combobox.current(0)

        self.memo_record_enabled = tk.BooleanVar(value=True)
        self.memo_record_checkbox = ttk.Checkbutton(memo_frame, text="상담기록", variable=self.memo_record_enabled)
        self.memo_record_checkbox.grid(row=0, column=2, sticky=tk.W, pady=2, padx=(20, 0))

        self.sales_record_enabled = tk.BooleanVar(value=True)
        self.sales_record_checkbox = ttk.Checkbutton(memo_frame, text="판매기록", variable=self.sales_record_enabled)
        self.sales_record_checkbox.grid(row=0, column=3, sticky=tk.W, pady=2, padx=(10, 0))
        
        sms_frame = ttk.LabelFrame(main_frame, text="문자 정보", padding="10")
        sms_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        ttk.Label(sms_frame, text="문자 내용을 붙여넣기 하세요:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.sms_text = scrolledtext.ScrolledText(sms_frame, width=70, height=10)
        self.sms_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=2)
        
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=10)
        
        self.check_driver_button = ttk.Button(button_frame, text="드라이버 확인", command=self.check_drivers)
        self.check_driver_button.pack(side=tk.LEFT, padx=5)
        
        self.test_connection_button = ttk.Button(button_frame, text="연결 테스트", command=self.test_connection)
        self.test_connection_button.pack(side=tk.LEFT, padx=5)

        self.clear_paste_button = ttk.Button(button_frame, text="지우고 붙여넣기 (F4)", command=self.clear_and_paste)
        self.clear_paste_button.pack(side=tk.LEFT, padx=5)
        
        self.process_button = ttk.Button(button_frame, text="처리하기 (F5)", command=self.process_sms)
        self.process_button.pack(side=tk.LEFT, padx=5)
        
        self.status_label = ttk.Label(main_frame, text="준비됨", foreground="green")
        self.status_label.grid(row=4, column=0, columnspan=2, pady=5)
        
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)
        db_frame.columnconfigure(1, weight=1)
        sms_frame.columnconfigure(0, weight=1)
        sms_frame.rowconfigure(1, weight=1)

    # (ExtensionPage의 나머지 함수들은 이전과 동일하게 유지됩니다)
    def clear_and_paste(self, event=None):
        try:
            clipboard_content = self.controller.root.clipboard_get()
            self.sms_text.delete("1.0", tk.END)
            self.sms_text.insert("1.0", clipboard_content)
        except tk.TclError:
            print("클립보드에서 텍스트를 가져올 수 없습니다.")
    
    def update_status(self, message, color="black"):
        self.status_label.config(text=message, foreground=color)
        self.controller.root.update_idletasks()
    
    def check_drivers(self):
        try:
            all_drivers = pyodbc.drivers()
            sql_drivers = [driver for driver in all_drivers if 'SQL Server' in driver]
            if sql_drivers:
                messagebox.showinfo("SQL Server 드라이버", f"설치된 SQL Server 드라이버:\n\n{'\n'.join(sql_drivers)}")
            else:
                messagebox.showwarning("드라이버 없음", f"SQL Server 드라이버가 설치되지 않았습니다.\n\n설치된 모든 드라이버:\n{'\n'.join(pyodbc.drivers())}")
        except Exception as e:
            messagebox.showerror("오류", f"드라이버 확인 중 오류 발생: {str(e)}")
    
    def test_connection(self):
        def test_thread():
            try:
                self.test_connection_button.config(state="disabled")
                self.update_status("연결 테스트 중...", "blue")
                conn = self.connect_to_database()
                cursor = conn.cursor()
                cursor.execute("SELECT 1")
                if cursor.fetchone():
                    self.update_status("연결 테스트 성공", "green")
                    messagebox.showinfo("성공", "데이터베이스 연결이 성공했습니다!")
                else:
                    self.update_status("연결 테스트 실패", "red")
                    messagebox.showerror("실패", "연결은 되었지만 쿼리 실행에 실패했습니다.")
                conn.close()
            except Exception as e:
                self.update_status(f"연결 실패: {str(e)}", "red")
                messagebox.showerror("연결 실패", f"데이터베이스 연결에 실패했습니다:\n\n{str(e)}")
            finally:
                self.test_connection_button.config(state="normal")
        threading.Thread(target=test_thread, daemon=True).start()
    
    def connect_to_database(self):
        server = self.server_entry.get()
        database = self.db_entry.get()
        username = self.user_entry.get()
        password = self.password_entry.get()
        drivers = ["ODBC Driver 17 for SQL Server", "ODBC Driver 18 for SQL Server", "ODBC Driver 13 for SQL Server", "ODBC Driver 11 for SQL Server", "SQL Server Native Client 11.0", "SQL Server Native Client 10.0", "SQL Server"]
        for driver in drivers:
            try:
                conn = pyodbc.connect(f"DRIVER={{{driver}}};SERVER={server};DATABASE={database};UID={username};PWD={password}", autocommit=False)
                print(f"연결 성공: {driver}")
                return conn
            except pyodbc.Error as e:
                last_error = str(e)
        raise Exception(f"데이터베이스 연결 실패: {last_error}")
    
    def connect_to_google_sheets(self):
        try:
            creds_path = "ubiplus-n8n-05e734977179.json"
            if not os.path.exists(creds_path):
                    raise Exception(f"Google Sheets 인증 파일({creds_path})이 필요합니다.")
            
            credentials = Credentials.from_service_account_file(
                creds_path,
                scopes=['https://www.googleapis.com/auth/spreadsheets']
            )
            gc = gspread.authorize(credentials)
            return gc.open_by_key(self.controller.sheet_id).worksheet(self.controller.sheet_name)
        except Exception as e:
            raise Exception(f"구글 시트 연결 실패: {str(e)}")
    
    def find_matching_row(self, sheet, sms_content):
        try:
            normalized_sms = "".join(sms_content.split())
            all_values = sheet.get_all_values()
            for i, row in enumerate(all_values[1:], start=2):
                if len(row) >= 1 and row[0]:
                    if "".join(row[0].split()) == normalized_sms:
                        return i, row
            return None, None
        except Exception as e:
            raise Exception(f"구글 시트 데이터 검색 실패: {str(e)}")

    def create_new_sales_record(self, cursor, company_idx, writer_name, employee_idx):
        cursor.execute("""
            SELECT TOP 1 * FROM tb_sales
            WHERE company_idx = ? AND (contents LIKE '%월납%' OR contents LIKE '%1개월%' OR contents LIKE '%6개월%' OR contents LIKE '%12개월%')
            ORDER BY insert_date DESC
        """, (company_idx,))
        sales_columns = [column[0] for column in cursor.description]
        source_sale_row = cursor.fetchone()
        if not source_sale_row:
            print(f"경고: company_idx={company_idx}에 대한 기준 판매 건을 찾을 수 없어, 판매 기록을 생성하지 않습니다.")
            return
        source_sale_dict = dict(zip(sales_columns, source_sale_row))
        today_str_yyyymmdd = datetime.now().strftime('%Y%m%d')
        cursor.execute("SELECT MAX(sale_no) FROM tb_sales WHERE sale_no LIKE ?", (f"{today_str_yyyymmdd}%",))
        last_sale_no = cursor.fetchone()[0]
        new_seq = int(last_sale_no[-4:]) + 1 if last_sale_no else 1
        new_sale_data = source_sale_dict.copy()
        new_sale_data.update({
            'sale_no': f"{today_str_yyyymmdd}{new_seq:04d}",
            'sale_date': datetime.now().strftime('%Y-%m-%d'),
            'state': "미수", 'etax_bill_send_yn': '', 'etax_bill_no': None,
            'refer_no': '', 'settlement_price': 0, 'insert_date': datetime.now(),
            'update_date': datetime.now(), 'employee_idx': employee_idx,
            'insert_emp_idx': employee_idx, 'update_emp_idx': employee_idx,
            'insert_emp_name': writer_name, 'update_emp_name': writer_name
        })
        insert_columns = [col for col in sales_columns if col.lower() != 'idx']
        insert_values = [new_sale_data.get(col) for col in insert_columns]
        sql_insert_sales = f"INSERT INTO tb_sales ({', '.join(insert_columns)}) OUTPUT INSERTED.idx VALUES ({', '.join(['?'] * len(insert_columns))})"
        new_sale_idx = cursor.execute(sql_insert_sales, insert_values).fetchone()[0]
        if not new_sale_idx:
            raise Exception("새로운 판매 IDX를 가져오지 못했습니다.")
        print(f"새로운 sale_idx 생성됨: {new_sale_idx}")
        cursor.execute("SELECT * FROM tb_sales_detail WHERE sale_idx = ?", (source_sale_dict['idx'],))
        detail_columns = [column[0] for column in cursor.description]
        source_details = cursor.fetchall()
        if not source_details: return
        insert_detail_columns = [col for col in detail_columns if col.lower() != 'idx']
        sql_insert_detail = f"INSERT INTO tb_sales_detail ({', '.join(insert_detail_columns)}) VALUES ({', '.join(['?'] * len(insert_detail_columns))})"
        for detail_row in source_details:
            new_detail_data = dict(zip(detail_columns, detail_row))
            new_detail_data.update({'sale_idx': new_sale_idx, 'insert_date': datetime.now(), 'update_date': datetime.now(), 'insert_emp_name': writer_name, 'update_emp_name': writer_name})
            insert_detail_values = [new_detail_data.get(col) for col in insert_detail_columns]
            cursor.execute(sql_insert_detail, insert_detail_values)
        print(f"{len(source_details)}건의 tb_sales_detail 기록 추가 완료.")

    def update_expiry_and_sales(self, conn, cdkeys, usage_months, server_months):
        try:
            cursor = conn.cursor()
            writer_name = self.writer_combobox.get()
            if not writer_name: raise Exception("작성자를 선택해주세요.")
            cursor.execute("SELECT idx FROM tb_Employee WHERE name = ?", (writer_name,))
            emp_result = cursor.fetchone()
            if not emp_result: raise Exception(f"'{writer_name}' 이름을 가진 직원이 없습니다.")
            employee_idx = emp_result[0]
            cursor.execute("SELECT user_idx FROM tb_purchase WHERE cdkey = ?", (cdkeys[0],))
            first_result = cursor.fetchone()
            if not first_result: raise Exception(f"첫 번째 인증키 {cdkeys[0]}를 찾을 수 없어, 기록을 생성할 수 없습니다.")
            representative_user_idx = first_result[0]
            memo_log_by_key = []
            for cdkey in cdkeys:
                print(f"--- 인증키: {cdkey} 만료일 연장 시작 ---")
                cursor.execute("SELECT purchase_date, server_date FROM tb_purchase WHERE cdkey = ?", (cdkey,))
                result = cursor.fetchone()
                if not result:
                    print(f"경고: 인증키 {cdkey}를 찾을 수 없습니다. 건너뜁니다.")
                    continue
                current_purchase_date, current_server_date = result
                key_specific_memo = [f"[{cdkey}]"]
                update_parts, params = [], []
                if usage_months:
                    base_date = current_purchase_date or datetime.now()
                    new_date = base_date + relativedelta(months=int(usage_months))
                    if base_date.date() != new_date.date(): key_specific_memo.append(f"사용만료일: {base_date.strftime('%Y-%m-%d')} >> {new_date.strftime('%Y-%m-%d')}")
                    else: key_specific_memo.append("사용만료일: (변경 없음)")
                    update_parts.append("purchase_date = ?"); params.append(new_date)
                if server_months:
                    base_date = current_server_date or datetime.now()
                    new_date = base_date + relativedelta(months=int(server_months))
                    if base_date.date() != new_date.date(): key_specific_memo.append(f"서버만료일: {base_date.strftime('%Y-%m-%d')} >> {new_date.strftime('%Y-%m-%d')}")
                    else: key_specific_memo.append("서버만료일: (변경 없음)")
                    update_parts.append("server_date = ?"); params.append(new_date)
                if len(key_specific_memo) > 1: memo_log_by_key.append("\n".join(key_specific_memo))
                if update_parts:
                    params.append(cdkey)
                    cursor.execute(f"UPDATE tb_purchase SET {', '.join(update_parts)} WHERE cdkey = ?", params)
            
            if self.sales_record_enabled.get():
                print("--- 판매기록 생성 시작 ---")
                self.create_new_sales_record(cursor, representative_user_idx, writer_name, employee_idx)
            else:
                print("--- 판매기록 건너뜀 (체크 해제됨) ---")

            if self.memo_record_enabled.get():
                if memo_log_by_key:
                    print("--- 상담기록 생성 시작 ---")
                    final_memo_text = "\n\n".join(memo_log_by_key)
                    try:
                        cursor.execute("""INSERT INTO tb_companymemo (company_idx, write_date, writer, memo, insert_date, insert_emp_name, update_date, update_emp_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                                    (representative_user_idx, datetime.now().strftime('%Y-%m-%d'), writer_name, final_memo_text, datetime.now(), "관리자", datetime.now(), "관리자"))
                        print(f"상담 내역 기록 완료: {writer_name}")
                    except Exception as memo_error:
                        print(f"경고: 상담 내역 기록에 실패했습니다: {memo_error}")
                else:
                    print("--- 기록할 상담 내용이 없어 상담기록 건너뜀 ---")
            else:
                print("--- 상담기록 건너뜀 (체크 해제됨) ---")

            conn.commit()
            return True, f"{len(cdkeys)}개의 인증키에 대한 만료일 연장 및 관련 기록 처리가 완료되었습니다."
        except Exception as e:
            conn.rollback(); raise Exception(f"데이터베이스 작업 실패 (롤백됨): {str(e)}")
    
    def update_google_sheet_result(self, sheet, row_num, success, error_message=""):
        try:
            update_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            update_values = ["성공", update_time, ""] if success else ["실패", update_time, str(error_message)]
            sheet.update(f'E{row_num}:G{row_num}', [update_values])
        except Exception as e:
            print(f"구글 시트 업데이트 실패: {str(e)}")
    
    def process_sms(self, event=None):
        def process_thread():
            self.process_button.config(state="disabled")
            self.update_status("처리 중...", "blue")
            sheet, row_num = None, None
            try:
                sms_content = self.sms_text.get("1.0", tk.END).strip()
                if not sms_content:
                    messagebox.showwarning("경고", "문자 내용을 입력해주세요.")
                    self.update_status("준비됨", "green")
                    return
                self.update_status("구글 시트 연결 중...", "blue")
                sheet = self.connect_to_google_sheets()
                self.update_status("문자 내용 매칭 중...", "blue")
                row_num, row_data = self.find_matching_row(sheet, sms_content)
                if not row_num:
                    messagebox.showwarning("데이터 없음", "입력된 문자와 일치하는 데이터를 구글 시트에서 찾을 수 없습니다.")
                    self.update_status("일치하는 데이터 없음", "orange")
                    return
                cdkeys_str = row_data[1].strip() if len(row_data) > 1 else ""
                if not cdkeys_str: raise Exception("구글 시트에 인증키(cdkey)가 없습니다.")
                cdkeys = [key.strip() for key in re.split(r'[,\s\n]+', cdkeys_str) if key.strip()]
                usage_months = row_data[2].strip() if len(row_data) > 2 else ""
                server_months = row_data[3].strip() if len(row_data) > 3 else ""
                if not usage_months and not server_months:
                    messagebox.showinfo("정보", "연장할 개월 수가 없어 처리하지 않습니다.")
                    self.update_status("연장 개월 수 없음", "orange")
                    return
                self.update_status("데이터베이스 연결 중...", "blue")
                conn = self.connect_to_database()
                self.update_status("만료일 업데이트 및 기록 중...", "blue")
                success, message = self.update_expiry_and_sales(conn, cdkeys, usage_months, server_months)
                conn.close()
                self.update_google_sheet_result(sheet, row_num, True)
                self.update_status("처리 완료", "green")
                messagebox.showinfo("성공", message)
            except Exception as e:
                error_msg = str(e)
                self.update_status(f"오류: {error_msg}", "red")
                messagebox.showerror("오류", error_msg)
                if sheet and row_num:
                    self.update_google_sheet_result(sheet, row_num, False, error_msg)
            finally:
                self.process_button.config(state="normal")
        threading.Thread(target=process_thread, daemon=True).start()

def main():
    # main 함수는 변경사항 없습니다.
    try:
        from dateutil.relativedelta import relativedelta
    except ImportError:
        messagebox.showerror("라이브러리 필요", "정확한 날짜 계산을 위해 'python-dateutil' 라이브러리가 필요합니다.\n\n터미널(cmd)에서 'pip install python-dateutil'을 실행해주세요.")
        return
    root = tk.Tk()
    app = ExpiryExtensionApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
