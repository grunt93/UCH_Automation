from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
import time

# ===============================================
#                【設定變數區】
# ===============================================

# 替換成您的目標登入頁面網址
LOGIN_URL = "https://std.uch.edu.tw/Std_Xerox/Login_Index.aspx" 
YOUR_ACCOUNT = "D11213201" # 替換成您的學號/帳號
MY_PASSWORD = "Gg0976682163" # 替換成您的密碼

# 登入成功後要跳轉的目標網址 (缺曠記錄頁面)
TARGET_URL = "https://std.uch.edu.tw/Std_Xerox/Miss_ct.aspx" 

# 缺曠記錄表格的 ID (來自您提供的 HTML 內容)
TABLE_ID = "ctl00_ContentPlaceHolder1_gw_absent"

try:
    # 1. 初始化瀏覽器並訪問登入頁面
    driver = webdriver.Chrome()
    driver.get(LOGIN_URL)
    print(f"已成功訪問登入頁面: {LOGIN_URL}")
    
    # 2. 執行登入操作
    account_input = driver.find_element(By.NAME, "account")
    password_input = driver.find_element(By.NAME, "account_pass")
    sign_in_button = driver.find_element(By.NAME, "SignIn")
    
    account_input.send_keys(YOUR_ACCOUNT)
    password_input.send_keys(MY_PASSWORD) 
    print("帳號密碼已填寫完畢。")
    
    sign_in_button.click()
    print("已點擊登入按鈕。")
        
    # 給予伺服器一點時間處理登入請求並跳轉 (例如 3 秒)
    print("等待 3 秒讓登入流程完成...")
    time.sleep(3)
    
    # 3. 直接跳轉到目標頁面 (Miss_ct.aspx)
    driver.get(TARGET_URL)
    print(f"✅ 已直接跳轉到目標頁面: {TARGET_URL}")
    
    # 4. 擷取表格資訊
    
    # 等待表格元素出現 (最多 10 秒)，確保頁面已載入
    print(f"等待表格 ({TABLE_ID}) 載入...")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, TABLE_ID))
    )
    
    # 找到表格元素
    table = driver.find_element(By.ID, TABLE_ID)
    
    # 抓取所有行 (Row)
    rows = table.find_elements(By.TAG_NAME, "tr")
    
    extracted_data = []
    
    # 抓取表頭 (Header)
    if rows:
        header_row = rows[0]
        # 表頭使用 <th> 標籤
        headers = [th.text for th in header_row.find_elements(By.TAG_NAME, "th")]
        extracted_data.append(headers)
        
        # 抓取資料行 (從第二行開始)
        for row in rows[1:]:
            # 資料使用 <td> 標籤
            cols = row.find_elements(By.TAG_NAME, "td")
            if cols:
                data_row = [col.text for col in cols]
                extracted_data.append(data_row)
    
    # 5. 顯示擷取到的資料
    
    print("\n" + "="*70)
    print("                      【擷取到的缺曠記錄】")
    print("="*70)
    
    if extracted_data:
        # 為了美觀，計算每個欄位的最大寬度
        col_widths = [max(len(str(item)) for item in col) for col in zip(*extracted_data)]
        # 由於課程名稱最長，需要確保它的寬度足夠
        col_widths[2] = max(col_widths[2], 15) # 確保課號欄位至少有 15 個寬度
        
        for i, row in enumerate(extracted_data):
            # 格式化並列印行
            formatted_row = " | ".join(f"{item:<{col_widths[j]}}" for j, item in enumerate(row))
            print(formatted_row)
            if i == 0:
                # 在標題下方列印分隔線
                print("-" * (sum(col_widths) + 3 * (len(col_widths) - 1)))
    else:
        print("未找到缺曠記錄或表格中沒有資料。")
        
    print("="*70)
    
    # 6. 結束
    time.sleep(5) # 讓您有時間看到結果
    driver.quit()
    print("\n程式執行完畢，瀏覽器已關閉。")

except Exception as e:
    print(f"初始化或整體執行過程中發生錯誤: {e}")
    try:
        driver.quit()
    except:
        pass