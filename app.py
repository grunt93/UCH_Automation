import time
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from collections import defaultdict
from typing import Set, Dict, Any, List, Tuple

# 引入 Selenium 相關模組
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# ===============================================
#                【設定常數區】
# ===============================================

# 網站 URL
LOGIN_URL = "https://std.uch.edu.tw/Std_Xerox/Login_Index.aspx" 
TARGET_URL = "https://std.uch.edu.tw/Std_Xerox/Miss_ct.aspx" 
TABLE_ID = "ctl00_ContentPlaceHolder1_gw_absent"

# 課程應計節次因子 (用戶手動維護的資料)
COURSE_FACTORS: Dict[str, int] = {
    "程式設計與應用(三)": 4,
    "資料庫系統與實習": 2,
    "網頁資料庫程式開發實作": 4,
    "廣域網路與實習": 4,
    "性別與文化": 2,
    "RHCE紅帽Linux系統自動化": 4,
}
# 假別類型
ABSENCE_TYPES = ['事假', '病假', '遲到', '曠課']

class MissingAttendanceApp:
    def __init__(self, master):
        self.master = master
        master.title("學務系統缺曠課查詢工具")
        
        # 儲存瀏覽器實例
        self.driver = None 
        
        # 創建 GUI 介面元素
        self.create_widgets(master)
        
        # 設置狀態訊息
        self.set_status("準備就緒。請輸入學號和密碼。")

    def create_widgets(self, master):
        # --- 登入資訊框架 ---
        input_frame = ttk.Frame(master, padding="10")
        input_frame.pack(fill='x')

        # 帳號
        ttk.Label(input_frame, text="學號/帳號:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.account_entry = ttk.Entry(input_frame, width=30)
        self.account_entry.insert(0, "D11213201") # 範例學號
        self.account_entry.grid(row=0, column=1, padx=5, pady=5)

        # 密碼
        ttk.Label(input_frame, text="密碼:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.password_entry = ttk.Entry(input_frame, width=30, show='*')
        self.password_entry.insert(0, "Gg0976682163") # 範例密碼
        self.password_entry.grid(row=1, column=1, padx=5, pady=5)

        # 運行按鈕
        self.run_button = ttk.Button(input_frame, text="開始查詢並計算", command=self.run_scraper)
        self.run_button.grid(row=2, column=0, columnspan=2, pady=10)

        # --- 狀態訊息 ---
        self.status_label = ttk.Label(master, text="", foreground="blue", padding="10")
        self.status_label.pack(fill='x')
        
        # --- 結果顯示框架 ---
        result_frame = ttk.Frame(master, padding="10")
        result_frame.pack(fill='both', expand=True)
        
        # 定義 Treeview (表格)
        columns = ['課程名稱'] + ABSENCE_TYPES + ['總缺課數量', '總天數']
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings')
        
        # 設置欄位標題與寬度
        self.tree.heading('課程名稱', text='課程名稱', anchor='w')
        self.tree.column('課程名稱', width=200, anchor='w')
        
        # 其他數值欄位
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
        # 檢查是否運行在 PyInstaller 打包環境中
        if getattr(sys, 'frozen', False):
            # PyInstaller 設置的臨時路徑
            base_path = sys._MEIPASS
        else:
            # 一般 Python 腳本運行時的路徑
            base_path = os.path.dirname(__file__)

        # 組合驅動程式的完整路徑
        return os.path.join(base_path, driver_name)

    def scrape_and_calculate(self, account: str, password: str) -> List[List[str]]:
        """
        核心爬蟲和計算邏輯
        返回整理好的表格數據 (List[List[str]])
        """
        self.set_status("1/7 正在初始化瀏覽器...")
        
        try:
            # 嘗試使用自動管理驅動，如果失敗，則使用手動路徑
            try:
                self.driver = webdriver.Chrome()
            except WebDriverException:
                # 嘗試使用 PyInstaller 兼容路徑 (假設備份驅動在同目錄)
                driver_path = self.get_driver_path()
                self.driver = webdriver.Chrome(executable_path=driver_path)
            
            self.driver.get(LOGIN_URL)
            self.set_status(f"2/7 已訪問登入頁面: {LOGIN_URL}")
            time.sleep(1) # 稍等

            # 2. 執行登入操作
            account_input = self.driver.find_element(By.NAME, "account")
            password_input = self.driver.find_element(By.NAME, "account_pass")
            sign_in_button = self.driver.find_element(By.NAME, "SignIn")
            
            account_input.send_keys(account)
            password_input.send_keys(password) 
            self.set_status("3/7 帳號密碼已填寫，正在登入...")
            
            sign_in_button.click()
            
            # 給予伺服器一點時間處理登入請求並跳轉 (例如 3 秒)
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
                    
            # 6. 整理最終輸出列表 (聯集邏輯)
            
            # 取得所有在網頁上有記錄的課程名稱
            recorded_courses: Set[str] = set(summary_data.keys())
            # 取得所有在 COURSE_FACTORS 中定義的課程名稱
            factor_courses: Set[str] = set(COURSE_FACTORS.keys())
            
            # 創建聯集：先放 COURSE_FACTORS 中的課程，再放其他有記錄的課程
            final_course_list: List[str] = list(factor_courses)
            for course in sorted(list(recorded_courses - factor_courses)):
                final_course_list.append(course)
                
            output_rows = []
            
            for course_name in final_course_list:
                counts = summary_data[course_name] 

                total_absent = counts.get('總缺課數量', 0)
                factor = COURSE_FACTORS.get(course_name)
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
                     # 警告訊息仍然顯示在狀態欄
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