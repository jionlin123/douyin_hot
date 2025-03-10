import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ================== 配置区（必须修改）==================
CHROME_PATH = r'C:\Program Files\Google\Chrome\Application\chrome.exe'  # 你的Chrome路径
CHROMEDRIVER_PATH = r'C:\Users\19312\Desktop\chromedriver.exe'         # 你手动下载的驱动路径
HEADLESS_MODE = False  # 调试时设为False（显示浏览器窗口）
# =======================================================

def init_browser():
    """浏览器初始化配置"""
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.binary_location = CHROME_PATH  # 显式指定Chrome路径

    if HEADLESS_MODE:
        chrome_options.add_argument("--headless=new")

    # 直接使用本地驱动（跳过自动安装）
    service = Service(executable_path=CHROMEDRIVER_PATH)
    
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    })
    return driver

def login_douyin(driver):
    """处理登录流程"""
    try:
        # 直接使用IP地址绕过DNS解析问题
        driver.get('http://140.143.250.142/hotspot/rank/')
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//button[contains(text(), "登录")]'))
        ).click()
        print("=> 请用抖音APP扫码登录（完成后在此窗口按回车继续）...")
        input()  # 阻塞直到用户确认
    except Exception as e:
        print(f"登录失败：{str(e)}")
        driver.quit()
        exit()

def get_hot_data(driver):
    """采集热点数据"""
    try:
        # 新增页面加载状态检查
        WebDriverWait(driver, 20).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # 滚动加载逻辑优化
        last_height = driver.execute_script("return document.body.scrollHeight")
        for _ in range(3):
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
        
        # 更稳定的元素定位方式
        items = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'div.HotItem'))
        )
        
        hot_list = []
        for item in items[:10]:  # 取前10条数据
            title = item.find_element(By.CLASS_NAME, 'HotItem-title').text.strip()
            hot_val = item.find_element(By.CLASS_NAME, 'HotItem-metrics').text.strip()
            hot_list.append({"标题": title, "热度值": hot_val})
            
        return pd.DataFrame(hot_list)
    except Exception as e:
        print(f"数据采集失败：{str(e)}")
        return pd.DataFrame()

if __name__ == "__main__":
    driver = init_browser()
    try:
        login_douyin(driver)
        df = get_hot_data(driver)
        if not df.empty:
            df.to_excel("抖音热点榜单_最终版.xlsx", index=False)
            print("数据已保存到：抖音热点榜单_最终版.xlsx")
        else:
            print("警告：未获取到有效数据")
    except Exception as e:
        print(f"致命错误：{str(e)}")
    finally:
        driver.quit()
        print("浏览器已安全关闭")