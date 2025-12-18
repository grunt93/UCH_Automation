from selenium import webdriver
from selenium.webdriver.common.by import By
import time

# 替換成您的目標登入頁面網址
LOGIN_URL = "https://std.uch.edu.tw/Std_Xerox/Login_Index.aspx" 
YOUR_ACCOUNT = "D11213201"
MY_PASSWORD = "Gg0976682163" # 請務必填寫您的密碼

try:
    # 1. 初始化瀏覽器
    driver = webdriver.Chrome()
    driver.get(LOGIN_URL)
    print(f"已成功訪問登入頁面: {LOGIN_URL}")
    
    # 2. 執行登入操作
    account_input = driver.find_element(By.NAME, "account")
    password_input = driver.find_element(By.NAME, "account_pass")
    sign_in_button = driver.find_element(By.NAME, "SignIn")
    
    # 3. 填寫帳號密碼
    account_input.send_keys(YOUR_ACCOUNT)
    password_input.send_keys(MY_PASSWORD) 
    print("帳號密碼已填寫完畢。")
    
    # 4. 點擊登入按鈕
    sign_in_button.click()
    print("已點擊登入按鈕。")
    
    # 5. 等待一段時間，讓頁面有時間跳轉或顯示結果
    print("等待 5 秒...")
    time.sleep(5) 
    
    # 6. 結束
    driver.quit()
    print("\n程式執行完畢，瀏覽器已關閉。")

except Exception as e:
    print(f"初始化或整體執行過程中發生錯誤: {e}")
    try:
        driver.quit()
    except:
        pass