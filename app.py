from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait 
from selenium.webdriver.support import expected_conditions as EC 
import time

# 替換成您的目標登入頁面網址
LOGIN_URL = "https://std.uch.edu.tw/Std_Xerox/Login_Index.aspx" 
YOUR_ACCOUNT = "D11213201"
MY_PASSWORD = "Gg0976682163" # 請務必填寫您的密碼

# 登入頁面原本的 Title，用於判斷頁面是否發生跳轉
ORIGINAL_TITLE = "系統登入" 

try:
    driver = webdriver.Chrome()
    driver.get(LOGIN_URL)
    print(f"已成功訪問登入頁面: {LOGIN_URL}")
    
    # 執行登入操作
    account_input = driver.find_element(By.NAME, "account")
    password_input = driver.find_element(By.NAME, "account_pass")
    sign_in_button = driver.find_element(By.NAME, "SignIn")
    
    account_input.send_keys(YOUR_ACCOUNT)
    password_input.send_keys(MY_PASSWORD) 
    print("帳號密碼已填寫完畢。")
    
    sign_in_button.click()
    print("已點擊登入按鈕，等待頁面跳轉...")
    
    # --- 關鍵：等待頁面跳轉並抓取 Title ---
    
    # 7. 等待頁面標題改變，如果標題仍然是 ORIGINAL_TITLE，則判定為登入失敗
    try:
        # 等待標題不再是「系統登入」（最多 10 秒）
        WebDriverWait(driver, 10).until_not(EC.title_is(ORIGINAL_TITLE))
        
        # 抓取並顯示當前頁面標題
        current_title = driver.title
        
        print("\n--- 登入後伺服器回傳資訊 ---")
        print(f"✅ 登入成功！")
        print(f"新頁面的 Title 是：【{current_title}】")
        
        # 由於您要求抓取 Title，我們到此為止
        # 如果需要抓取其他元素，可以在這裡繼續添加程式碼
        
    except Exception:
        # 如果等待超時，表示頁面標題仍為「系統登入」，或發生其他錯誤
        if driver.title == ORIGINAL_TITLE:
             print("❌ 登入失敗！頁面標題未改變，請檢查帳號密碼。")
        else:
             print("⚠️ 登入過程中發生未知錯誤。")
        
    # 8. 結束
    time.sleep(5) 
    driver.quit()
    print("\n程式執行完畢，瀏覽器已關閉。")

except Exception as e:
    print(f"初始化或整體執行過程中發生錯誤: {e}")
    try:
        driver.quit()
    except:
        pass