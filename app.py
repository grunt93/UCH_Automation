import time
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from collections import defaultdict
from typing import Set, Dict, Any, List, Tuple
import json

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# ===============================================
#                【設定常數區】
# ===============================================

# 配置檔案名稱 (用於儲存 COURSE_FACTORS)
CONFIG_FILE = "course_factors_config.json"
LOGIN_URL = "https://std.uch.edu.tw/Std_Xerox/Login_Index.aspx" 
TARGET_URL = "https://std.uch.edu.tw/Std_Xerox/Miss_ct.aspx" 
TABLE_ID = "ctl00_ContentPlaceHolder1_gw_absent"
ABSENCE_TYPES = ['事假', '病假', '遲到', '曠課']

# 預設的課程應計節次因子 (只有第一次運行找不到配置檔時才會使用)
DEFAULT_COURSE_FACTORS: Dict[str, int] = {}

# --- 資料持久化函數 ---

def get_app_path():
    """獲取程式運行的基礎路徑，兼容打包後的環境"""
    if getattr(sys, 'frozen', False):
        # 如果是打包後的 exe/app，使用執行檔所在的目錄
        return os.path.dirname(sys.executable)
    # 如果是直接運行 python 腳本，使用腳本所在的目錄
    return os.path.dirname(os.path.abspath(__file__))

def get_config_filepath():
    """獲取配置檔案的完整路徑"""
    # 讓配置檔和執行檔在同一目錄
    return os.path.join(get_app_path(), CONFIG_FILE)

def load_factors_from_file() -> Dict[str, int]:
    """從檔案載入課程因子，失敗則使用預設值"""
    filepath = get_config_filepath()
    if os.path.exists(filepath):
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 確保載入的鍵值對都是正確類型
                return {str(k): int(v) for k, v in data.items()}
        except Exception:
            # 載入失敗，使用預設值並發出警告
            print(f"警告：載入配置檔失敗，使用預設因子。")
            return DEFAULT_COURSE_FACTORS
    # 檔案不存在，使用預設值
    return DEFAULT_COURSE_FACTORS

def save_factors_to_file(factors: Dict[str, int]):
    """將課程因子儲存到檔案"""
    filepath = get_config_filepath()
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            # ensure_ascii=False 確保中文能正確寫入 JSON
            json.dump(factors, f, ensure_ascii=False, indent=4)
        print(f"課程因子已成功儲存到: {filepath}")
    except Exception as e:
        messagebox.showerror("儲存錯誤", f"無法儲存課程因子到檔案: {e}")
        print(f"儲存錯誤: {e}")


# --- 編輯視窗類別 ---

class EditFactorsWindow(tk.Toplevel):
    def __init__(self, master, current_factors: Dict[str, int], update_callback):
        super().__init__(master)
        self.title("編輯課程因子 (課程名稱: 應計節次)")
        self.geometry("450x400")
        self.transient(master) # 設置為模態對話框
        self.grab_set() 
        self.protocol("WM_DELETE_WINDOW", self.on_close) # 處理關閉按鈕
        
        # 儲存傳入的因子副本，在視窗內部編輯
        self.current_factors = current_factors 
        self.update_callback = update_callback
        
        self.factor_tree = None
        self.create_widgets()
        self.populate_tree()

    def create_widgets(self):
        # Frame for Treeview
        tree_frame = ttk.Frame(self, padding="10")
        tree_frame.pack(fill='both', expand=True)

        # Treeview (顯示課程因子)
        columns = ('課程名稱', '應計節次')
        self.factor_tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        self.factor_tree.heading('課程名稱', text='課程名稱')
        self.factor_tree.column('課程名稱', width=250, anchor='w')
        self.factor_tree.heading('應計節次', text='應計節次')
        self.factor_tree.column('應計節次', width=100, anchor='center')

        # Scrollbar
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.factor_tree.yview)
        vsb.pack(side='right', fill='y')
        self.factor_tree.configure(yscrollcommand=vsb.set)
        
        self.factor_tree.pack(fill='both', expand=True)
        
        # --- 按鈕框架 ---
        button_frame = ttk.Frame(self, padding="10")
        button_frame.pack(fill='x')
        
        ttk.Button(button_frame, text="新增課程", command=self.add_factor).pack(side='left', padx=5)
        ttk.Button(button_frame, text="修改節次", command=self.edit_factor).pack(side='left', padx=5)
        ttk.Button(button_frame, text="移除課程", command=self.remove_factor).pack(side='left', padx=5)
        ttk.Button(button_frame, text="儲存並關閉", command=self.save_and_close).pack(side='right', padx=5)
        ttk.Button(button_frame, text="取消", command=self.on_close).pack(side='right', padx=5)

    def populate_tree(self):
        """用 current_factors 的內容填充 Treeview"""
        # 清空舊數據
        for item in self.factor_tree.get_children():
            self.factor_tree.delete(item)
            
        # 填充新數據 (按課程名稱排序顯示)
        for name, factor in sorted(self.current_factors.items()):
            self.factor_tree.insert('', tk.END, values=(name, factor))

    def add_factor(self):
        """新增課程因子"""
        new_name = simpledialog.askstring("新增課程", "請輸入新的課程名稱:", parent=self)
        if new_name:
            new_name = new_name.strip()
            if new_name in self.current_factors:
                messagebox.showwarning("警告", f"課程【{new_name}】已存在，請使用修改節次功能。", parent=self)
                return
                
            new_factor_str = simpledialog.askstring("新增課程", f"請輸入【{new_name}】的應計節次 (數字):", parent=self)
            try:
                if new_factor_str is None: return
                new_factor = int(new_factor_str)
                if new_factor <= 0:
                    raise ValueError
                self.current_factors[new_name] = new_factor
                self.populate_tree()
            except (TypeError, ValueError):
                messagebox.showerror("錯誤", "應計節次必須是一個正整數。", parent=self)
                
    def edit_factor(self):
        """修改選定課程的節次"""
        selected_item = self.factor_tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "請先選擇要修改的課程。", parent=self)
            return

        item = selected_item[0]
        current_name = self.factor_tree.item(item, 'values')[0]
        
        new_factor_str = simpledialog.askstring("修改節次", f"請輸入【{current_name}】新的應計節次 (目前為 {self.current_factors[current_name]}):", parent=self)
        
        if new_factor_str is not None:
            try:
                new_factor = int(new_factor_str)
                if new_factor <= 0:
                    raise ValueError
                self.current_factors[current_name] = new_factor
                self.populate_tree()
            except (TypeError, ValueError):
                messagebox.showerror("錯誤", "應計節次必須是一個正整數。", parent=self)

    def remove_factor(self):
        """移除選定的課程"""
        selected_item = self.factor_tree.selection()
        if not selected_item:
            messagebox.showwarning("警告", "請先選擇要移除的課程。", parent=self)
            return
            
        item = selected_item[0]
        course_name = self.factor_tree.item(item, 'values')[0]

        if messagebox.askyesno("確認移除", f"確定要移除課程【{course_name}】嗎?", parent=self):
            del self.current_factors[course_name]
            self.populate_tree()

    def save_and_close(self):
        """儲存並關閉視窗"""
        # 儲存到檔案
        save_factors_to_file(self.current_factors)
        # 通知主程式更新課程因子
        self.update_callback(self.current_factors)
        self.destroy()

    def on_close(self):
        """處理關閉和取消"""
        # 如果使用者點擊 X 或取消，則直接關閉不儲存
        self.destroy()

# --- 主程式類別 ---

class MissingAttendanceApp:
    def __init__(self, master):
        self.master = master
        master.title("學務系統缺曠課查詢工具")
        
        # 【修改】程式啟動時，載入課程因子
        self.COURSE_FACTORS = load_factors_from_file()
        
        self.driver = None 
        self.create_widgets(master)
        self.set_status("準備就緒。請輸入學號和密碼。")

    def update_factors(self, new_factors: Dict[str, int]):
        """從編輯視窗接收並更新課程因子 (供主程式使用)"""
        self.COURSE_FACTORS = new_factors

    def open_edit_factors_window(self):
        """開啟編輯課程因子視窗"""
        # 傳遞當前因子的副本 (使用 .copy())，避免在編輯時直接影響主程式運行
        # 只有在用戶點擊「儲存並關閉」時，才會通過 update_factors 更新主程式的 self.COURSE_FACTORS
        EditFactorsWindow(self.master, self.COURSE_FACTORS.copy(), self.update_factors)


    def create_widgets(self, master):
        # --- 登入資訊框架 ---
        input_frame = ttk.Frame(master, padding="10")
        input_frame.pack(fill='x')

        # 帳號
        ttk.Label(input_frame, text="學號/帳號:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.account_entry = ttk.Entry(input_frame, width=30)
        self.account_entry.grid(row=0, column=1, padx=5, pady=5)

        # 密碼
        ttk.Label(input_frame, text="密碼:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.password_entry = ttk.Entry(input_frame, width=30, show='*')
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)
        
        # --- 按鈕框架 (包含查詢和編輯) ---
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)

        # 查詢按鈕
        self.run_button = ttk.Button(button_frame, text="開始查詢並計算", command=self.run_scraper)
        self.run_button.pack(side='left', padx=10)

        # 【新增】編輯課程因子按鈕
        ttk.Button(button_frame, text="編輯課程因子", command=self.open_edit_factors_window).pack(side='left', padx=10)

        # --- 狀態訊息 ---
        self.status_label = ttk.Label(master, text="", foreground="blue", padding="10")
        self.status_label.pack(fill='x')
        
        # --- 結果顯示框架 (Treeview) ---
        result_frame = ttk.Frame(master, padding="10")
        result_frame.pack(fill='both', expand=True)
        
        # 定義 Treeview (表格)
        columns = ['課程名稱'] + ABSENCE_TYPES + ['總缺課數量', '總天數']
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings')
        
        # 設置欄位標題與寬度
        self.tree.heading('課程名稱', text='課程名稱', anchor='w')
        self.tree.column('課程名稱', width=200, anchor='w')
        for col in ABSENCE_TYPES:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=60, anchor='center')
        self.tree.heading('總缺課數量', text='總節次')
        self.tree.column('總缺課數量', width=70, anchor='center')
        self.tree.heading('總天數', text='總天數')
        self.tree.column('總天數', width=70, anchor='center')
        
        # 添加滾動條
        vsb = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        vsb.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(fill='both', expand=True)

    def set_status(self, message, is_error=False):
        """更新狀態欄的訊息和顏色"""
        self.status_label.config(text=message)
        self.status_label.config(foreground="red" if is_error else "blue")
        self.master.update_idletasks() # 強制更新介面

    def get_driver_path(self, driver_name="chromedriver.exe"):
        """獲取 ChromeDriver 的路徑 (用於打包兼容性)"""
        # 使用共用函數確保路徑正確
        base_path = get_app_path()
        return os.path.join(base_path, driver_name)

    def scrape_and_calculate(self, account: str, password: str) -> List[List[str]]:
        """
        核心爬蟲和計算邏輯
        返回整理好的表格數據 (List[List[str]])
        """
        self.set_status("1/7 正在初始化瀏覽器...")
        
        try:
            # 嘗試初始化驅動
            try:
                self.driver = webdriver.Chrome()
            except WebDriverException:
                # 嘗試使用 PyInstaller 兼容路徑 
                driver_path = self.get_driver_path()
                self.driver = webdriver.Chrome(executable_path=driver_path)
            
            self.driver.get(LOGIN_URL)
            self.set_status(f"2/7 已訪問登入頁面: {LOGIN_URL}")
            time.sleep(1) 

            # 2. 執行登入操作
            account_input = self.driver.find_element(By.NAME, "account")
            password_input = self.driver.find_element(By.NAME, "account_pass")
            sign_in_button = self.driver.find_element(By.NAME, "SignIn")
            
            account_input.send_keys(account)
            password_input.send_keys(password) 
            self.set_status("3/7 帳號密碼已填寫，正在登入...")
            
            sign_in_button.click()
            
            time.sleep(3)
            
            # 3. 直接跳轉到目標頁面 (Miss_ct.aspx)
            self.driver.get(TARGET_URL)
            self.set_status(f"4/7 登入成功，已跳轉到缺曠記錄頁面: {TARGET_URL}")
            
            # 4. 擷取表格資訊
            self.set_status(f"5/7 正在抓取表格數據...")
            
            # 等待表格元素出現 (最多 10 秒)
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.ID, TABLE_ID))
            )
            
            table = self.driver.find_element(By.ID, TABLE_ID)
            rows = table.find_elements(By.TAG_NAME, "tr")
            
            raw_data: List[Tuple[str, str]] = [] 
            
            if len(rows) > 1:
                for row in rows[1:]:
                    cols = row.find_elements(By.TAG_NAME, "td")
                    if cols:
                        # 抓取：[課程名稱(索引2), 狀態(索引3)]
                        course_name = cols[2].text
                        absence_status = cols[3].text
                        raw_data.append((course_name, absence_status))

            # 5. 統計數據
            self.set_status("6/7 正在計算總結數據...")
            
            # 結構: {課程名稱: {狀態: 數量, 總缺課數量: 數量}}
            summary_data = defaultdict(lambda: defaultdict(float)) 
            
            for course_name, status in raw_data: 
                if status in ABSENCE_TYPES:
                    summary_data[course_name][status] += 1
                    summary_data[course_name]['總缺課數量'] += 1
                    
            # 6. 整理最終輸出列表
            
            recorded_courses: Set[str] = set(summary_data.keys())
            # 【重要】使用 self.COURSE_FACTORS (已載入或更新)
            factor_courses: Set[str] = set(self.COURSE_FACTORS.keys())
            
            # 創建聯集清單
            final_course_list: List[str] = list(factor_courses)
            for course in sorted(list(recorded_courses - factor_courses)):
                final_course_list.append(course)
                
            output_rows = []
            
            for course_name in final_course_list:
                counts = summary_data[course_name] 

                total_absent = counts.get('總缺課數量', 0)
                # 【重要】從 self.COURSE_FACTORS 獲取因子
                factor = self.COURSE_FACTORS.get(course_name)
                calculated_days_str = "" 

                # 計算總天數
                if factor:
                    if total_absent > 0:
                        calculated_days = total_absent / factor
                        calculated_days_str = f"{calculated_days:.2f}"
                    else:
                        calculated_days_str = "0.00"
                else:
                     calculated_days_str = "N/A"
                     if total_absent > 0:
                         self.set_status(f"⚠️ 警告: 課程【{course_name}】缺少應計節次，總天數無法計算 (N/A)。", is_error=True)

                row: List[str] = [course_name]
                for status in ABSENCE_TYPES:
                    row.append(str(int(counts.get(status, 0)))) 
                row.append(str(int(total_absent))) 
                row.append(calculated_days_str) 
                
                output_rows.append(row)
            
            self.set_status("7/7 資料抓取與計算完成！")
            return output_rows

        except (TimeoutException, NoSuchElementException) as e:
            self.set_status(f"錯誤：抓取頁面元素或登入超時。請檢查帳密或網路。錯誤: {e.__class__.__name__}", is_error=True)
            return []
        except WebDriverException as e:
            self.set_status(f"錯誤：瀏覽器驅動程式問題。請確保 Chrome 和 ChromeDriver 版本匹配。錯誤: {e.__class__.__name__}", is_error=True)
            return []
        except Exception as e:
            self.set_status(f"發生未預期的錯誤: {e}", is_error=True)
            return []
        finally:
            if self.driver:
                self.driver.quit()
                self.driver = None

    def run_scraper(self):
        """點擊按鈕時執行的函數"""
        
        # 清除舊的表格數據
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        account = self.account_entry.get().strip()
        password = self.password_entry.get()
        
        if not account or not password:
            messagebox.showerror("錯誤", "請輸入學號和密碼！")
            return
            
        self.run_button.config(state=tk.DISABLED, text="查詢中...")
        self.set_status("開始運行爬蟲程式...")
        
        # 執行核心邏輯
        data = self.scrape_and_calculate(account, password)
        
        # 顯示結果到 Treeview
        if data:
            self.set_status(f"查詢完成。總計找到 {len(data)} 門課程記錄。", is_error=False)
            for row in data:
                # 插入數據到 Treeview
                self.tree.insert('', tk.END, values=row)
        else:
            self.set_status("查詢失敗或未找到任何缺曠記錄。", is_error=True)

        self.run_button.config(state=tk.NORMAL, text="開始查詢並計算")


if __name__ == "__main__":
    # 創建主視窗
    root = tk.Tk()
    app = MissingAttendanceApp(root)
    # 設置視窗大小
    root.geometry("700x500") 
    # 啟動主循環
    root.mainloop()