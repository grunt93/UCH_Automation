# 爬蟲核心邏輯

import os 
import time
from collections import defaultdict
from typing import Set, Dict, List, Tuple

# 引入 Selenium 相關模組
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# 引入常數和路徑函數
import config_data 

# ===============================================
#                【爬蟲核心函數】
# ===============================================

def get_driver_path(driver_name="chromedriver.exe"):
    """獲取 ChromeDriver 的路徑 (用於打包兼容性)"""
    # 調用 config_data 中的通用路徑函數
    base_path = config_data.get_app_path()
    return os.path.join(base_path, driver_name)

def scrape_and_calculate(
    account: str, 
    password: str, 
    course_factors: Dict[str, int],
    set_status_callback # 傳入 GUI 的狀態更新函式
) -> List[List[str]]:
    """
    核心爬蟲和計算邏輯
    返回整理好的表格數據 (List[List[str]])
    """
    
    # ***新增常數: 假單列印頁面 URL***
    XEROX_URL = "https://std.uch.edu.tw/Std_Xerox/Xerox.aspx"
    
    driver = None
    set_status_callback("1/9 正在初始化瀏覽器...")
    
    try:
        # 嘗試初始化驅動
        try:
            driver = webdriver.Chrome()
        except WebDriverException:
            # 嘗試使用 PyInstaller 兼容路徑 
            driver_path = get_driver_path()
            driver = webdriver.Chrome(executable_path=driver_path)
        
        driver.get(config_data.LOGIN_URL)
        set_status_callback(f"2/9 已訪問登入頁面: {config_data.LOGIN_URL}")
        time.sleep(1) 

        # 2. 執行登入操作
        account_input = driver.find_element(By.NAME, "account")
        password_input = driver.find_element(By.NAME, "account_pass")
        sign_in_button = driver.find_element(By.NAME, "SignIn")
        
        account_input.send_keys(account)
        password_input.send_keys(password) 
        set_status_callback("3/9 帳號密碼已填寫，正在登入...")
        
        sign_in_button.click()
        time.sleep(3)
        
        # ==========================================================
        # 步驟 A: 抓取缺曠課記錄 (原 TARGET_URL)
        # ==========================================================
        
        # 4. 跳轉到缺曠記錄頁面
        driver.get(config_data.TARGET_URL)
        set_status_callback(f"4/9 登入成功，已跳轉到缺曠記錄頁面: {config_data.TARGET_URL}")
        
        # 5. 擷取缺曠課表格資訊
        set_status_callback(f"5/9 正在抓取缺曠課表格數據...")
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, config_data.TABLE_ID))
        )
        
        table = driver.find_element(By.ID, config_data.TABLE_ID)
        rows = table.find_elements(By.TAG_NAME, "tr")
        
        # raw_data 結構: (course_name, absence_status, week_number, section)
        raw_data: List[Tuple[str, str, str, str]] = [] 
        if len(rows) > 1:
            for row in rows[1:]:
                cols = row.find_elements(By.TAG_NAME, "td")
                if cols and len(cols) >= 5: 
                    week_number = cols[0].text
                    course_name = cols[2].text
                    absence_status = cols[3].text
                    section = cols[4].text
                    
                    raw_data.append((course_name, absence_status, week_number, section))
        
        # 輸出原始缺曠課數據到終端機
        print("\n" + "="*70)
        print("【原始缺曠課記錄 (Miss_ct.aspx) - 終端機輸出】")
        print("格式: (課程名稱, 缺曠狀態, 週別, 節次)")
        print("-"*70)
        if raw_data:
            for course_name, status, week, section in raw_data:
                print(f"({course_name}, {status}, 週{week}, 節次{section})")
        else:
            print("無缺曠課記錄。")
        print("="*70 + "\n")


        # ==========================================================
        # 步驟 B: 抓取假單記錄 (新頁面: Xerox.aspx)
        # ==========================================================

        # 6. 跳轉到假單列印頁面
        driver.get(XEROX_URL)
        set_status_callback(f"6/9 已跳轉到假單列印頁面: {XEROX_URL}")
        
        # 7. 擷取假單表格資訊
        set_status_callback(f"7/9 正在抓取假單表格數據...")
        
        # 使用相同的 TABLE_ID, 假單回傳資料.txt 中 ID 確實是 ctl00_ContentPlaceHolder1_gw_absent
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, config_data.TABLE_ID)) 
        )
        
        xerox_table = driver.find_element(By.ID, config_data.TABLE_ID)
        xerox_rows = xerox_table.find_elements(By.TAG_NAME, "tr")
        
        xerox_data: List[str] = [] 
        if len(xerox_rows) > 1:
            for row in xerox_rows[1:]:
                cols = row.find_elements(By.TAG_NAME, "td")
                if cols and len(cols) >= 2:
                    # 根據 列印假單回傳資料.txt，週別是第二個欄位 (索引 1)
                    week_number = cols[1].text.strip()
                    xerox_data.append(week_number)

        # 輸出假單週別到終端機
        print("\n" + "="*70)
        print("【假單記錄週別 (Xerox.aspx) - 終端機輸出】")
        print("格式: 週別")
        print("-"*70)
        if xerox_data:
            # 為了避免多個週別重複顯示，這裡可以考慮只顯示不重複的週別
            unique_weeks = sorted(list(set(xerox_data)), key=lambda x: int(x.split(' ')[0]))
            for week in unique_weeks:
                print(f"假單記錄於: 週{week}")
        else:
            print("無假單記錄。")
        print("="*70 + "\n")
        
        # ==========================================================
        # 步驟 C: 統計計算 (只使用步驟 A 的 raw_data)
        # ==========================================================
        
        set_status_callback("8/9 正在計算總結數據...")
        
        # 統計數據 (只使用第一個頁面抓取的 raw_data，忽略週別和節次)
        summary_data = defaultdict(lambda: defaultdict(float)) 
        
        for course_name, status, _, _ in raw_data: 
            if status in config_data.ABSENCE_TYPES:
                summary_data[course_name][status] += 1
                summary_data[course_name]['總缺課數量'] += 1
                
        # 整理最終輸出列表
        recorded_courses: Set[str] = set(summary_data.keys())
        factor_courses: Set[str] = set(course_factors.keys())
        
        final_course_list: List[str] = list(factor_courses)
        for course in sorted(list(recorded_courses - factor_courses)):
            final_course_list.append(course)
            
        output_rows = []
        
        for course_name in final_course_list:
            counts = summary_data[course_name] 
            total_absent = counts.get('總缺課數量', 0)
            factor = course_factors.get(course_name)
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
                     set_status_callback(f"⚠️ 警告: 課程【{course_name}】缺少應計節次，總天數無法計算 (N/A)。", is_error=True)

            row: List[str] = [course_name]
            for status in config_data.ABSENCE_TYPES:
                row.append(str(int(counts.get(status, 0)))) 
            row.append(str(int(total_absent))) 
            row.append(calculated_days_str) 
            
            output_rows.append(row)
        
        set_status_callback("9/9 資料抓取與計算完成！")
        return output_rows

    except (TimeoutException, NoSuchElementException) as e:
        set_status_callback(f"錯誤：抓取頁面元素或登入超時。請檢查帳密或網路。錯誤: {e.__class__.__name__}", is_error=True)
        return []
    except WebDriverException as e:
        set_status_callback(f"錯誤：瀏覽器驅動程式問題。請確保 Chrome 和 ChromeDriver 版本匹配。錯誤: {e.__class__.__name__}", is_error=True)
        return []
    except Exception as e:
        set_status_callback(f"發生未預期的錯誤: {e}", is_error=True)
        return []
    finally:
        if driver:
            driver.quit()