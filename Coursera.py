import requests
import time
import random
import string
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import concurrent.futures

# ================= 配置区 =================
# RoxyBrowser 本地 API 配置
API_HOST = "http://127.0.0.1:5000"
API_TOKEN = "4ba21591e96dad03142b86e7ef106598"  
HEADERS = {"token": API_TOKEN}
# ==========================================

def _collect_profile_ids(node, profile_ids):
    """Recursively collect profile IDs from JSON payload."""
    if isinstance(node, dict):
        profile_id = node.get("profileId") or node.get("profile_id")
        if profile_id:
            profile_ids.add(str(profile_id))
        for key in ("id", "_id"):
            value = node.get(key)
            if value and isinstance(value, (str, int)):
                profile_ids.add(str(value))
        for value in node.values():
            _collect_profile_ids(value, profile_ids)
    elif isinstance(node, list):
        for item in node:
            _collect_profile_ids(item, profile_ids)

def get_all_profile_ids():
    """Fetch all profile IDs from local RoxyBrowser API."""
    endpoints = [
        "/api/v1/profile/list",
        "/api/v1/browser/list",
    ]
    for endpoint in endpoints:
        try:
            response = requests.get(f"{API_HOST}{endpoint}", headers=HEADERS, timeout=10)
            data = response.json()
            if data.get("code") == 0 or data.get("success"):
                payload = data.get("data", data)
                profile_ids = set()
                _collect_profile_ids(payload, profile_ids)
                profile_ids = sorted(profile_ids)
                if profile_ids:
                    return profile_ids
        except Exception as e:
            print(f"[ProfileAPI] {endpoint} failed: {e}")
    return []

def save_profile_ids(profile_ids, output_file="profile_ids.txt"):
    """Write profile IDs to a local text file for quick verification."""
    with open(output_file, "w", encoding="utf-8") as f:
        for profile_id in profile_ids:
            f.write(profile_id + "\n")

def get_random_account_info():
    """生成随机的 Gmail、姓名、密码和 Zipcode"""
    prefix_len = random.randint(8, 12)
    prefix = ''.join(random.choices(string.ascii_lowercase + string.digits, k=prefix_len))
    email = f"{prefix}@gmail.com"
    password = f"{prefix}pw"
    
    first_name = ''.join(random.choices(string.ascii_lowercase, k=5)).capitalize()
    last_name = ''.join(random.choices(string.ascii_lowercase, k=6)).capitalize()
    full_name = f"{first_name} {last_name}"
    
    zipcode = str(random.randint(10000, 99999))
    return email, full_name, password, zipcode

def get_card_from_xml():
    """从 account.xml 中随机抽取以 '---' 分割的信用卡信息"""
    file_path = "account.xml"
    if not os.path.exists(file_path):
        print("未找到 account.xml，使用默认测试卡号")
        return "4242424242424242", "12/25", "123"
    
    with open(file_path, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f.readlines() if '---' in line]
        
    if not lines:
        return "4242424242424242", "12/25", "123"
        
    selected_line = random.choice(lines)
    parts = selected_line.split('---')
    return parts[0], parts[1], parts[2]

def save_link_to_xml(link):
    """将获取到的链接追加写入 link.xml"""
    with open("link.xml", "a", encoding="utf-8") as f:
        f.write(link + "\n")
    print(f"✅ 成功保存链接至 link.xml: {link}")

def js_click(driver, element):
    """防遮挡点击：使用 JavaScript 强制点击元素"""
    driver.execute_script("arguments[0].click();", element)

# ----------------- 浏览器环境控制 API -----------------
def start_roxy_browser(profile_id):
    """唤醒 RoxyBrowser 环境"""
    url = f"{API_HOST}/api/v1/browser/start?profileId={profile_id}"
    try:
        response = requests.get(url, headers=HEADERS).json()
        if response.get("code") == 0 or response.get("success"):
            data = response.get("data", {})
            return data.get("debug_port"), data.get("webdriver")
    except Exception as e:
        print(f"[{profile_id}] 启动异常: {e}")
    return None, None

def close_roxy_browser(profile_id):
    """关闭 RoxyBrowser 环境"""
    url = f"{API_HOST}/api/v1/browser/stop?profileId={profile_id}"
    try:
        requests.get(url, headers=HEADERS)
    except Exception:
        pass

# ----------------- 核心业务逻辑 -----------------
def run_coursera_workflow(driver, profile_id):
    wait = WebDriverWait(driver, 20)
    email, full_name, password, zipcode = get_random_account_info()
    card_num, exp_date, cvc = get_card_from_xml()

    try:
        print(f"[{profile_id}] 步骤 2: 点击 Enroll for free")
        enroll_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Enroll for free') or contains(., 'Enroll')]")))
        js_click(driver, enroll_btn)

        print(f"[{profile_id}] 步骤 3: 输入邮箱")
        email_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@type='email']")))
        email_input.send_keys(email)
        continue_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Continue')]")))
        js_click(driver, continue_btn)

        print(f"[{profile_id}] 步骤 4: 填写姓名与密码")
        name_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@placeholder='Enter your full name' or @name='name']")))
        name_input.send_keys(full_name)
        pass_input = driver.find_element(By.XPATH, "//input[@type='password']")
        pass_input.send_keys(password)
        join_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Join for Free')]")))
        js_click(driver, join_btn)

        print(f"[{profile_id}] 步骤 5: 接受条款")
        accept_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'I accept')]")))
        js_click(driver, accept_btn)

        print(f"[{profile_id}] 步骤 6: 确认开始试用")
        trial_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Start Free Trial')]")))
        js_click(driver, trial_btn)

        print(f"[{profile_id}] 步骤 7: 填写账单国家与邮编")
        try:
            country_select = wait.until(EC.presence_of_element_located((By.XPATH, "//select[contains(@name, 'country') or contains(@id, 'country')]")))
            Select(country_select).select_by_visible_text("United States")
        except:
            country_box = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Select your country')]/..")))
            js_click(driver, country_box)
            us_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='United States']")))
            js_click(driver, us_option)
            
        zip_input = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[contains(@name, 'zip') or contains(@id, 'postal')]")))
        zip_input.send_keys(zipcode)

        print(f"[{profile_id}] 步骤 8: 填入信用卡并提交")
        try:
            iframe = wait.until(EC.presence_of_element_located((By.XPATH, "//iframe[contains(@name, '__privateStripeFrame') or contains(@title, 'Secure payment')]")))
            driver.switch_to.frame(iframe)
        except Exception:
            pass 

        card_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='cardnumber']")))
        card_input.send_keys(card_num)
        exp_input = driver.find_element(By.XPATH, "//input[@name='exp-date']")
        exp_input.send_keys(exp_date)
        cvc_input = driver.find_element(By.XPATH, "//input[@name='cvc']")
        cvc_input.send_keys(cvc)
        driver.switch_to.default_content() 
        
        submit_checkout_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Start Free Trial') or contains(., 'Submit')]")))
        js_click(driver, submit_checkout_btn)

        print(f"[{profile_id}] 步骤 9: 承诺并开始课程")
        try:
            commit_checkbox = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='checkbox']")))
            js_click(driver, commit_checkbox)
        except:
            pass 
        start_course_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Start the course')]")))
        js_click(driver, start_course_btn)

        print(f"[{profile_id}] 步骤 10: 弹窗确认继续")
        continue_success_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Continue')]")))
        js_click(driver, continue_success_btn)

        print(f"[{profile_id}] 步骤 11: 展开 Module 2 并点击目标章节")
        module_2 = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Module 2')]")))
        js_click(driver, module_2)
        redeem_item = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Redeem your Google AI Pro trial')]")))
        js_click(driver, redeem_item)

        print(f"[{profile_id}] 步骤 12: 同意荣誉准则并启动应用")
        honor_checkbox = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='checkbox']")))
        js_click(driver, honor_checkbox) 
        launch_btn = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Launch App')]")))
        js_click(driver, launch_btn)

        print(f"[{profile_id}] 步骤 13: 提取链接")
        time.sleep(5) 
        if len(driver.window_handles) > 1:
            driver.switch_to.window(driver.window_handles[-1])
            save_link_to_xml(driver.current_url)
        else:
            link_element = wait.until(EC.presence_of_element_located((By.XPATH, "//a[starts-with(@href, 'https://')]")))
            save_link_to_xml(link_element.get_attribute("href"))

        print(f"[{profile_id}] 🎉 业务流执行成功完成！")

    except Exception as e:
        print(f"[{profile_id}] ❌ 运行受阻，报错信息: {e}")

def run_automation(profile_id):
    debug_port, webdriver_path = start_roxy_browser(profile_id)
    if not debug_port:
        return

    chrome_options = Options()
    chrome_options.add_experimental_option("debuggerAddress", debug_port)
    service = Service(executable_path=webdriver_path) if webdriver_path else Service()
    
    driver = None
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        target_url = "https://www.coursera.org/professional-certificates/google-ai?action=enroll"
        if "professional-certificates/google-ai" not in driver.current_url:
            driver.get(target_url)
            
        run_coursera_workflow(driver, profile_id)
        
    except Exception as e:
        print(f"[{profile_id}] 驱动连接报错: {e}")
    finally:
        if driver:
            try:
                driver.quit() 
            except:
                pass
        close_roxy_browser(profile_id)

def main():
    # 【配置】在此填入您要并发操作的 RoxyBrowser 环境 ID (Profile ID)
    profile_ids = get_all_profile_ids()
    if not profile_ids:
        print("No profile IDs found from RoxyBrowser API.")
        return

    save_profile_ids(profile_ids)
    default_workers = 2
    workers_from_env = os.getenv("MAX_WORKERS", str(default_workers)).strip()
    try:
        max_workers = int(workers_from_env)
    except ValueError:
        print(f"Invalid MAX_WORKERS='{workers_from_env}', fallback to {default_workers}")
        max_workers = default_workers
    max_workers = max(1, min(max_workers, len(profile_ids)))
    
    print(f"🚀 开始 Coursera 自动化任务，最大并发窗口: {max_workers}")
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        executor.map(run_automation, profile_ids)

if __name__ == "__main__":
    main()
