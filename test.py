# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from datetime import datetime

# ========== 配置区 ==========
USERNAME = "13957990857@139.com"
PASSWORD = "a6KD8^DJE#"
LOGIN_URL = "http://10.76.148.208/dash/login1?redirectPath=%2Fdash%2Fhome"
START_DATE = "2026-05-11"
END_DATE = datetime.now().strftime("%Y-%m-%d")
BASE_SLEEP = 3
# ===========================

# 浏览器配置，屏蔽自动化提示、增强稳定性
chrome_opts = webdriver.ChromeOptions()
chrome_opts.add_argument("--start-maximized")
chrome_opts.add_argument("--disable-gpu")
chrome_opts.add_argument("--no-sandbox")
chrome_opts.add_argument("--disable-dev-shm-usage")
chrome_opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
chrome_opts.add_experimental_option("useAutomationExtension", False)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_opts)
wait = WebDriverWait(driver, 30)

try:
    # 1. 登录流程
    driver.get(LOGIN_URL)
    time.sleep(BASE_SLEEP)
    input_list = driver.find_elements(By.TAG_NAME, "input")
    input_list[0].send_keys(USERNAME)
    input_list[1].send_keys(PASSWORD)
    input_list[1].send_keys("\n")
    time.sleep(3)
    print("【1/7】登录成功")

    # 2. 打开报告 + 逐层菜单
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='报告']"))).click()
    time.sleep(BASE_SLEEP)
    print("【2/7】进入报告页面")

    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='金华']"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='金华分公司']"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='殷骏']"))).click()
    time.sleep(2)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//*[text()='FTTR+组网(1)']"))).click()
    time.sleep(BASE_SLEEP * 2)
    print("【3/7】进入目标报表页面")

    # 3. 核心：纯原生JS 完整操作时间弹窗（无jQuery、无特殊选择器）
    print("【4/7】开始执行时间筛选弹窗操作...")
    js_code = f"""
    (function() {{
        // 1. 定位时间筛选按钮并唤起弹窗
        const timeBtn = document.querySelector('.m-dashbox-dateTimeFilter');
        if (!timeBtn) {{
            console.log("未找到时间筛选按钮");
            return;
        }}
        timeBtn.click();

        // 延时执行弹窗内操作，等待DOM渲染
        setTimeout(function() {{
            // 2. 切换 固定时间 标签
            const tabList = document.querySelectorAll('.ant-tabs-nav div');
            for(let i = 0; i < tabList.length; i++) {{
                if(tabList[i].innerText.trim() === "固定时间") {{
                    tabList[i].click();
                    break;
                }}
            }}

            // 3. 定位起止日期输入框
            const dateInputs = document.querySelectorAll('.ant-modal-body input');
            if(dateInputs.length < 2) {{
                console.log("未找到日期输入框");
                return;
            }}
            const startInp = dateInputs[0];
            const endInp = dateInputs[1];

            // 解除只读、赋值、触发框架必需事件
            startInp.removeAttribute("readonly");
            endInp.removeAttribute("readonly");

            startInp.value = "{START_DATE}";
            endInp.value = "{END_DATE}";

            const evtInput = new Event('input', {{bubbles: true}});
            const evtChange = new Event('change', {{bubbles: true}});
            startInp.dispatchEvent(evtInput);
            startInp.dispatchEvent(evtChange);
            endInp.dispatchEvent(evtInput);
            endInp.dispatchEvent(evtChange);

            // 4. 点击确定关闭弹窗
            const btnList = document.querySelectorAll('.ant-modal-footer button');
            for(let j = 0; j < btnList.length; j++) {{
                if(btnList[j].innerText.trim() === "确定") {{
                    btnList[j].click();
                    break;
                }}
            }}
        }}, 1800);
    }})();
    """
    driver.execute_script(js_code)
    time.sleep(12)
    print(f"【5/7】时间筛选完成：{START_DATE} ~ {END_DATE}")

    # 4. JS 执行导出操作（纯原生遍历文本匹配）
    print("【6/7】开始导出Excel...")
    export_js = """
    (function(){
        // 找到导出主按钮
        const allBtns = document.querySelectorAll('button');
        let exportMain = null;
        for(let b of allBtns){
            if(b.innerText.trim() === "导出"){
                exportMain = b;
                break;
            }
        }
        if(exportMain){
            exportMain.click();
            setTimeout(function(){
                // 下拉菜单找Excel选项
                const menuItems = document.querySelectorAll('.ant-dropdown-menu-item');
                for(let item of menuItems){
                    if(item.innerText.includes("Excel")){
                        item.click();
                    }
                }
            }, 1200);
        }
    })();
    """
    driver.execute_script(export_js)
    time.sleep(6)
    print("【7/7】全部流程执行完毕，请查看下载文件夹")

except Exception as err:
    print(f"❌ 运行异常：{err}")
    driver.save_screenshot("err_snap.png")
    print("异常截图已保存至当前目录")
finally:
    print("\n流程结束，浏览器保持打开")
    print("111")
