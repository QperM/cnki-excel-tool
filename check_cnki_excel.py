import os
import threading
import time
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import re
import unicodedata


# ======== 知网检索页面基础 URL ========
BASE_SEARCH_URL = (
    "https://navi.cnki.net/knavi/detail?"
    "p=AcKg9NN3ni81TL7fIoPa1hV3EkBDcfLnU3iDohbmp01OJBRi7zJ6IDVALC8IqSkj7dZtqkKH59uBOpcCAdVL3kJySld8n7fsNldUb69KAwWEhfjCCewVHA=="
    "&uniplatform=NZKPT"
)

# 可调的网络重试次数和等待
OPEN_RETRY = 3
OPEN_RETRY_DELAY = 3


def normalize_title_strict(s: str) -> str:
    """
    标题严格匹配的“最小必要规范化”：
    - Unicode NFKC（全角/半角、兼容字符）
    - 去掉零宽字符
    - 合并所有空白（含换行/制表）为单个空格，并 strip
    不做标点替换/删减，以尽量保持“所有字符相同”的要求。
    """
    if s is None:
        return ""
    s = str(s)
    s = unicodedata.normalize("NFKC", s)
    # 移除零宽字符
    s = re.sub(r"[\u200B-\u200D\uFEFF]", "", s)
    # 合并空白
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ======== Selenium 启动配置 ========
def make_driver():
    options = webdriver.ChromeOptions()
    # 不使用无头模式，方便调试和查看；若需无头，可取消下一行注释
    # options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    # 规避部分浏览器安全限制，减少 ERR_CONNECTION_CLOSED 概率
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-insecure-localhost")
    options.add_argument("--disable-features=BlockInsecurePrivateNetworkRequests")
    options.add_argument("--remote-allow-origins=*")
    # 伪装为常见浏览器，绕过部分反自动化检测
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/118.0.5993.117 Safari/537.36"
    )

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options,
    )
    # 进一步降低被检测概率
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        """
    })
    driver.maximize_window()
    return driver


# ======== 打开页面，带重试，缓解 ERR_CONNECTION_CLOSED ========
def open_page_with_retry(driver, url: str, retries: int = OPEN_RETRY, delay: int = OPEN_RETRY_DELAY) -> bool:
    for i in range(retries):
        try:
            driver.get(url)
            return True
        except Exception as e:
            if "ERR_CONNECTION_CLOSED" in str(e):
                print(f"连接被关闭，重试 {i + 1}/{retries} ...")
                time.sleep(delay)
                continue
            # 其它异常直接抛出
            print(f"打开页面异常：{e}")
            time.sleep(delay)
    return False


# ======== 点击时间选择器并选择日期 ========
def select_date_by_click(driver, pub_date_str: str, debug_callback=None):
    """
    通过点击时间选择器来选择日期
    pub_date_str 格式：'2019-12-31' 或 '2019/12/31'
    debug_callback: 用于输出调试信息的回调函数
    """
    try:
        def debug_print(msg):
            print(f"[选择日期] {msg}")
            if debug_callback:
                debug_callback(msg)
        
        debug_print(f"开始选择日期：{pub_date_str}")
        
        # 解析日期
        date_obj = datetime.strptime(pub_date_str.replace('/', '-'), '%Y-%m-%d')
        year = date_obj.year
        month = date_obj.month
        day = date_obj.day
        
        debug_print(f"解析结果：年份={year}, 月份={month}, 日期={day}")
        
        # 显示当前页面信息
        current_url = driver.current_url
        page_title = driver.title
        debug_print(f"当前页面URL: {current_url[:100]}...")
        debug_print(f"当前页面标题: {page_title}")
        
        # 等待页面加载
        wait = WebDriverWait(driver, 10)
        time.sleep(2)  # 额外等待页面完全加载
        
        # 优先尝试下拉框 select#yearlist
        try:
            debug_print("尝试使用下拉框 #yearlist 选择年份...")
            year_select_el = wait.until(EC.presence_of_element_located((By.ID, "yearlist")))
            sel = Select(year_select_el)
            sel.select_by_visible_text(f"{year}年")
            debug_print("✓ 使用 yearlist 成功选择年份")
            time.sleep(1)
        except Exception as e:
            debug_print(f"使用 yearlist 失败: {e}")
            # 退回到旧的泛化点击逻辑
            # 点击左侧时间选择框（常见的类名或id，可能需要根据实际页面调整）
            time_selectors = [
                "//div[@class='time-select']",
                "//div[contains(@class, 'time')]",
                "//div[@id='timeSelect']",
                "//span[contains(text(), '时间')]",
                "//div[contains(@class, 'date')]",
                "//input[@placeholder*='时间']",
                "//input[@placeholder*='日期']",
                "//div[contains(@class, 'left')]//div[contains(@class, 'time')]",
                "//div[contains(@class, 'left')]//span[contains(text(), '时间')]",
                "//div[contains(@class, 'filter')]//div[contains(@class, 'time')]",
            ]
            
            debug_print(f"尝试查找时间选择器，共 {len(time_selectors)} 个选择器...")
            time_element = None
            found_selector = None
            for i, selector in enumerate(time_selectors):
                try:
                    debug_print(f"  尝试选择器 {i+1}/{len(time_selectors)}: {selector}")
                    time_element = wait.until(
                        EC.element_to_be_clickable((By.XPATH, selector))
                    )
                    if time_element:
                        found_selector = selector
                        debug_print(f"  ✓ 找到时间选择器！使用选择器: {selector}")
                        break
                except Exception as e2:
                    debug_print(f"  ✗ 选择器失败: {str(e2)[:50]}")
                    continue
            
            if not time_element:
                debug_print("✗ 所有时间选择器都失败，尝试查找页面所有可点击元素...")
                # 尝试查找页面左侧的所有div
                try:
                    left_divs = driver.find_elements(By.XPATH, "//div[contains(@class, 'left')]//div")
                    debug_print(f"  找到左侧 {len(left_divs)} 个div元素")
                    for i, div in enumerate(left_divs[:10]):  # 只检查前10个
                        try:
                            text = div.text[:50] if div.text else ""
                            debug_print(f"    div[{i}]: class={div.get_attribute('class')}, text={text}")
                        except:
                            pass
                except Exception as e3:
                    debug_print(f"  查找左侧元素失败: {e3}")
                
                debug_print("✗ 无法找到时间选择器，返回False")
                return False
            
            # 点击时间选择框
            debug_print("点击时间选择器...")
            try:
                driver.execute_script("arguments[0].click();", time_element)
                debug_print("✓ 时间选择器点击成功")
                time.sleep(2)  # 等待下拉菜单出现
            except Exception as e4:
                debug_print(f"✗ 点击时间选择器失败: {e4}")
                return False
            
            # 选择年份（点击年份下拉或年份选择器）
            debug_print(f"开始选择年份: {year}")
            year_selectors = [
                f"//span[text()='{year}']",
                f"//li[text()='{year}']",
                f"//div[text()='{year}']",
                f"//a[text()='{year}']",
                f"//span[contains(text(), '{year}')]",
                f"//li[contains(text(), '{year}')]",
            ]
            
            year_found = False
            for selector in year_selectors:
                try:
                    debug_print(f"  尝试年份选择器: {selector}")
                    year_elem = wait.until(EC.element_to_be_clickable((By.XPATH, selector)))
                    driver.execute_script("arguments[0].click();", year_elem)
                    debug_print(f"  ✓ 成功选择年份: {year}")
                    time.sleep(1)
                    year_found = True
                    break
                except Exception as e5:
                    debug_print(f"  ✗ 年份选择器失败: {str(e5)[:50]}")
                    continue
            
            if not year_found:
                debug_print(f"✗ 无法选择年份: {year}")
                return False
        
        # 选择月份 + 日期（按你给的结构：h1.jcfirstcol / dl.jcsecondcol / dd / a[text()=yyyy-mm-dd]）
        debug_print(f"开始选择月份: {month}")
        month_text = f"{month}月"

        # 先尝试直接点日期（有时月份默认已展开，或页面已包含该日期）
        debug_print(f"开始选择日期(直达): {pub_date_str}")
        date_xpath_any = f"//dl[contains(@class,'jcsecondcol')]//a[normalize-space(text())='{pub_date_str}']"
        try:
            date_a = wait.until(EC.presence_of_element_located((By.XPATH, date_xpath_any)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", date_a)
            time.sleep(0.2)
            driver.execute_script("arguments[0].click();", date_a)
            debug_print(f"  ✓ 直达点击日期成功: {pub_date_str}")
            time.sleep(0.8)
        except Exception as e:
            debug_print(f"  直达点击日期失败，将展开月份后再点: {str(e)[:120]}")

            # 展开月份：优先点击 li 下的 ins（通常是展开按钮），其次点 h1 / li 本身
            try:
                month_li = wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, f"//h1[@class='jcfirstcol' and normalize-space(text())='{month_text}']/parent::li")
                    )
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", month_li)
                time.sleep(0.2)

                expanded = False
                for exp_xpath, name in [
                    (".//ins", "ins(展开按钮)"),
                    (".//h1[@class='jcfirstcol']", "h1(月份标题)"),
                    (".", "li(月份容器)"),
                ]:
                    try:
                        el = month_li.find_element(By.XPATH, exp_xpath)
                        driver.execute_script("arguments[0].click();", el)
                        debug_print(f"  ✓ 点击 {name} 以展开月份: {month_text}")
                        time.sleep(0.6)
                        expanded = True
                        break
                    except Exception as e2:
                        debug_print(f"  点击 {name} 失败: {str(e2)[:80]}")

                if not expanded:
                    debug_print(f"  ✗ 无法展开月份 {month_text}")
                    return False
            except Exception as e3:
                debug_print(f"  ✗ 找不到月份 {month_text} 的容器 li: {str(e3)[:120]}")
                return False

            # 展开后：在该月份 li 下精确点日期
            debug_print(f"开始选择日期(展开后): {pub_date_str}")
            date_xpath_in_month = (
                f"//h1[@class='jcfirstcol' and normalize-space(text())='{month_text}']"
                f"/parent::li//dl[contains(@class,'jcsecondcol')]//a[normalize-space(text())='{pub_date_str}']"
            )
            try:
                date_a2 = wait.until(EC.presence_of_element_located((By.XPATH, date_xpath_in_month)))
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", date_a2)
                time.sleep(0.2)
                # 直接 JS click（不要求可点击）
                driver.execute_script("arguments[0].click();", date_a2)
                debug_print(f"  ✓ 展开后点击日期成功: {pub_date_str}")
                time.sleep(0.8)
            except Exception as e4:
                debug_print(f"  ✗ 展开后仍未找到/无法点击日期 {pub_date_str}: {str(e4)[:160]}")
                # 额外调试：列出该月份前几个日期，看看文本是否一致
                try:
                    sample = month_li.find_elements(By.XPATH, ".//dl[contains(@class,'jcsecondcol')]//a")[:5]
                    debug_print(f"  该月份示例日期(前5个): {[s.text for s in sample]}")
                except Exception:
                    pass
                return False
        
        # 确认选择（如果有确认按钮）
        try:
            confirm_btn = driver.find_element(By.XPATH, "//button[contains(text(), '确定')] | //button[contains(text(), '确认')] | //a[contains(text(), '确定')]")
            driver.execute_script("arguments[0].click();", confirm_btn)
            debug_print("✓ 点击确认按钮")
            time.sleep(1)
        except:
            debug_print("未找到确认按钮，跳过")
        
        debug_print("✓ 日期选择完成")
        return True
        
    except Exception as e:
        error_msg = f"选择日期时出错：{e}"
        print(f"[选择日期] {error_msg}")
        if debug_callback:
            debug_callback(error_msg)
        import traceback
        print(traceback.format_exc())
        return False


# ======== 在标题输入框中输入标题并检索 ========
def search_title(driver, title: str, debug_callback=None):
    """
    在检索框中输入标题并点击检索
    debug_callback: 用于输出调试信息的回调函数
    """
    try:
        def debug_print(msg):
            print(f"[检索标题] {msg}")
            if debug_callback:
                debug_callback(msg)
        
        debug_print(f"开始检索标题: {title[:50]}...")
        
        wait = WebDriverWait(driver, 10)
        time.sleep(1)  # 等待页面稳定
        
        # 查找标题输入框（尝试多种可能的选择器）
        title_input_selectors = [
            "//input[@placeholder*='题名']",
            "//input[@placeholder*='标题']",
            "//input[@placeholder*='关键词']",
            "//input[@id*='title']",
            "//input[@id*='keyword']",
            "//input[@name*='title']",
            "//input[@name*='keyword']",
            "//input[@class*='search']",
            "//textarea[@placeholder*='题名']",
            "//input[@type='text']",
            "//input",
        ]
        
        debug_print(f"尝试查找标题输入框，共 {len(title_input_selectors)} 个选择器...")
        title_input = None
        found_selector = None
        for i, selector in enumerate(title_input_selectors):
            try:
                debug_print(f"  尝试选择器 {i+1}/{len(title_input_selectors)}: {selector}")
                title_input = wait.until(EC.presence_of_element_located((By.XPATH, selector)))
                if title_input:
                    found_selector = selector
                    debug_print(f"  ✓ 找到标题输入框！使用选择器: {selector}")
                    # 显示输入框信息
                    try:
                        placeholder = title_input.get_attribute("placeholder") or ""
                        input_id = title_input.get_attribute("id") or ""
                        input_class = title_input.get_attribute("class") or ""
                        debug_print(f"    输入框信息: placeholder={placeholder}, id={input_id}, class={input_class[:50]}")
                    except:
                        pass
                    break
            except Exception as e:
                debug_print(f"  ✗ 选择器失败: {str(e)[:50]}")
                continue
        
        if not title_input:
            debug_print("✗ 所有输入框选择器都失败，尝试查找页面所有input元素...")
            try:
                all_inputs = driver.find_elements(By.XPATH, "//input")
                debug_print(f"  找到 {len(all_inputs)} 个input元素")
                for i, inp in enumerate(all_inputs[:10]):  # 只检查前10个
                    try:
                        placeholder = inp.get_attribute("placeholder") or ""
                        inp_id = inp.get_attribute("id") or ""
                        debug_print(f"    input[{i}]: id={inp_id}, placeholder={placeholder}")
                    except:
                        pass
            except Exception as e:
                debug_print(f"  查找input元素失败: {e}")
            
            debug_print("✗ 无法找到标题输入框，返回False")
            return False
        
        # 清空并输入标题
        debug_print("清空输入框并输入标题...")
        try:
            title_input.clear()
            time.sleep(0.3)
            title_input.send_keys(title)
            debug_print(f"✓ 已输入标题: {title[:50]}...")
            time.sleep(0.5)
        except Exception as e:
            debug_print(f"✗ 输入标题失败: {e}")
            return False
        
        # 查找并点击检索按钮
        debug_print("查找检索按钮...")
        search_btn_selectors = [
            "//button[contains(text(), '检索')]",
            "//a[contains(text(), '检索')]",
            "//input[@type='submit']",
            "//button[@type='submit']",
            "//div[contains(@class, 'search-btn')]",
            "//span[contains(text(), '检索')]",
            "//button[contains(@class, 'search')]",
            "//a[contains(@class, 'search')]",
        ]
        
        search_btn = None
        found_btn_selector = None
        for selector in search_btn_selectors:
            try:
                search_btn = driver.find_element(By.XPATH, selector)
                if search_btn:
                    found_btn_selector = selector
                    debug_print(f"  ✓ 找到检索按钮！使用选择器: {selector}")
                    break
            except:
                continue
        
        if search_btn:
            try:
                driver.execute_script("arguments[0].click();", search_btn)
                debug_print("✓ 点击检索按钮成功")
            except Exception as e:
                debug_print(f"✗ 点击检索按钮失败: {e}")
                return False
        else:
            debug_print("未找到检索按钮，尝试按回车键...")
            try:
                title_input.send_keys("\n")
                debug_print("✓ 已按回车键")
            except Exception as e:
                debug_print(f"✗ 按回车键失败: {e}")
                return False
        
        # 等待检索结果加载
        debug_print("等待检索结果加载...")
        time.sleep(3)
        debug_print(f"当前页面URL: {driver.current_url[:100]}...")
        debug_print(f"当前页面标题: {driver.title}")
        debug_print("✓ 检索完成")
        return True
        
    except Exception as e:
        error_msg = f"检索标题时出错：{e}"
        print(f"[检索标题] {error_msg}")
        if debug_callback:
            debug_callback(error_msg)
        import traceback
        print(traceback.format_exc())
        return False


# ======== 在结果列表中查找完全匹配的标题（处理分页） ========
def find_title_in_results(driver, title: str, max_pages: int = 50, debug_callback=None) -> bool:
    """
    在检索结果中查找完全匹配的标题，支持翻页
    返回 True 表示找到，False 表示未找到
    debug_callback: 用于输出调试信息的回调函数
    """
    try:
        def debug_print(msg):
            print(f"[查找标题] {msg}")
            if debug_callback:
                debug_callback(msg)
        
        wait = WebDriverWait(driver, 10)
        raw_title = title
        title = normalize_title_strict(title)
        debug_print(f"开始查找标题(规范化后): {title[:80]}...")
        debug_print(f"最多查找 {max_pages} 页")
        
        current_page = 1
        
        while current_page <= max_pages:
            debug_print(f"\n--- 第 {current_page} 页 ---")
            # 等待当前页结果加载
            time.sleep(2)
            
            # 显示当前页面信息
            current_url = driver.current_url
            page_title = driver.title
            debug_print(f"当前页面URL: {current_url[:100]}...")
            debug_print(f"当前页面标题: {page_title}")
            
            # 读取页面上的当前页和总页数信息（在循环开始时读取）
            actual_current_page = current_page
            actual_total_pages = max_pages
            try:
                current_page_elem = driver.find_element(By.ID, "partiallistcurrent")
                total_pages_elem = driver.find_element(By.ID, "partiallistcount2")
                actual_current_page = int(current_page_elem.text.strip())
                actual_total_pages = int(total_pages_elem.text.strip())
                debug_print(f"页面分页信息：当前页 {actual_current_page}/{actual_total_pages}")
                
                # 如果已经超过总页数，停止
                if actual_current_page > actual_total_pages:
                    debug_print(f"  ✗ 当前页({actual_current_page})已超过总页数({actual_total_pages})，停止查找")
                    break
            except Exception as e:
                debug_print(f"  ⚠ 无法读取页面分页信息: {e}，继续使用循环计数")
            
            # 方式1: 使用结果列表结构化抓取（最可靠）
            debug_print("方式1: 结构化抓取结果标题 td.name a ...")
            try:
                result_title_elems = driver.find_elements(By.XPATH, "//td[contains(@class,'name')]//a")
                debug_print(f"  找到结果标题元素数量: {len(result_title_elems)}")

                # 输出前5条用于调试
                for i, el in enumerate(result_title_elems[:5]):
                    t = normalize_title_strict(el.text)
                    debug_print(f"    结果[{i+1}]: {t[:80]}")

                for el in result_title_elems:
                    t = normalize_title_strict(el.text)
                    if t == title:
                        debug_print("  ✓✓✓ 结构化列表中找到完全匹配标题！")
                        return True
            except Exception as e:
                debug_print(f"  ✗ 结构化抓取失败: {e}")

            # 方式2: 页面全文搜索（Ctrl+F 思路），但用规范化后再包含匹配
            debug_print("方式2: 页面全文搜索（Ctrl+F方式，规范化后包含）...")
            title_selectors = [
                "//td[contains(@class,'name')]//a",
                "//a[contains(@href, 'kcms2/article/abstract')]",
                "//a",
            ]
            
            title_elements = []
            found_selector = None
            for selector in title_selectors:
                try:
                    elements = driver.find_elements(By.XPATH, selector)
                    if elements:
                        title_elements = elements
                        found_selector = selector
                        debug_print(f"  ✓ 找到 {len(elements)} 个标题元素，使用选择器: {selector}")
                        break
                except:
                    continue
            
            if not title_elements:
                debug_print("  ✗ 未找到任何标题元素")
            else:
                debug_print(f"  检查前 {min(10, len(title_elements))} 个标题元素...")
                # 在当前页查找完全匹配的标题
                for i, elem in enumerate(title_elements[:20]):  # 只检查前20个
                    try:
                        text = normalize_title_strict(elem.text)
                        if text:
                            if text == title:
                                debug_print(f"  ✓✓✓ 在第 {i+1} 个元素中找到完全匹配的标题！")
                                debug_print(f"    匹配的文本: {text[:50]}...")
                                return True
                            elif i < 5:  # 只显示前5个用于调试
                                debug_print(f"    元素[{i+1}]: {text[:50]}...")
                    except Exception as e:
                        if i < 5:
                            debug_print(f"    元素[{i+1}]获取文本失败: {e}")

            # 方式3: 真正的 Ctrl+F（document.body.innerText），但先规范化
            debug_print("方式3: document.body.innerText 规范化后包含匹配...")
            try:
                page_text = driver.execute_script("return document.body ? document.body.innerText : '';") or ""
                page_norm = normalize_title_strict(page_text)
                if title and title in page_norm:
                    debug_print("  ✓ 在页面全文(规范化)中包含该标题")
                    return True
                else:
                    debug_print("  ✗ 页面全文(规范化)不包含该标题")
            except Exception as e:
                debug_print(f"  ✗ 全文提取失败: {e}")
            
            # 如果当前页没找到，尝试翻到下一页
            # 使用实际读取的页面信息来判断
            if actual_current_page < actual_total_pages and current_page < max_pages:
                debug_print(f"当前页未找到，尝试翻到第 {current_page + 1} 页...")
                # 查找"下一页"按钮（使用正确的选择器）
                next_btn_selectors = [
                    "//a[@class='page-next']",  # 优先使用这个
                    "//a[contains(@class, 'page-next')]",
                    "//a[contains(text(), '下一页')]",
                    "//a[contains(text(), '下页')]",
                    "//a[contains(@class, 'next')]",
                ]
                
                next_btn = None
                found_next_selector = None
                for selector in next_btn_selectors:
                    try:
                        next_btn = driver.find_element(By.XPATH, selector)
                        if next_btn:
                            # 检查是否有disable类（注意：是disable不是disabled）
                            classes = next_btn.get_attribute("class") or ""
                            if "disable" not in classes.lower() and "disabled" not in classes.lower():
                                found_next_selector = selector
                                debug_print(f"  ✓ 找到下一页按钮，使用选择器: {selector}")
                                debug_print(f"    按钮class: {classes}")
                                break
                    except Exception as e:
                        debug_print(f"  ✗ 选择器 {selector} 失败: {str(e)[:50]}")
                        continue
                
                if next_btn and found_next_selector:
                    try:
                        driver.execute_script("arguments[0].click();", next_btn)
                        current_page += 1
                        debug_print(f"  ✓ 已点击下一页，等待加载...")
                        time.sleep(3)  # 等待下一页加载
                        # 重新读取页面信息以确认翻页成功
                        try:
                            current_page_elem = driver.find_element(By.ID, "partiallistcurrent")
                            actual_current_page = int(current_page_elem.text.strip())
                            debug_print(f"  翻页后确认：当前页 {actual_current_page}/{actual_total_pages}")
                        except:
                            pass
                        continue
                    except Exception as e:
                        debug_print(f"  ✗ 点击下一页按钮失败: {e}")
                        break
                else:
                    debug_print("  ✗ 未找到可用的下一页按钮（可能已禁用或不存在），已到最后一页")
                    break
            else:
                debug_print(f"  ✗ 已达到最大页数限制或已到最后一页（{actual_current_page}/{actual_total_pages}），停止翻页")
                break
            
            current_page += 1
        
        debug_print(f"\n✗✗✗ 在所有 {current_page-1} 页中都未找到完全匹配的标题")
        return False
        
    except Exception as e:
        error_msg = f"查找标题时出错：{e}"
        print(f"[查找标题] {error_msg}")
        if debug_callback:
            debug_callback(error_msg)
        import traceback
        print(traceback.format_exc())
        return False


# ======== 检查单行：按日期+标题在知网页面检索并校验 ========
def check_title_at_date(driver, pub_date_str: str, title: str, debug_callback=None) -> bool:
    """
    检查在给定日期下是否能找到完全相同的标题
    返回 True 表示能找到，False 表示未找到
    debug_callback: 用于输出调试信息的回调函数
    """
    try:
        def debug_print(msg):
            print(f"[检查标题] {msg}")
            if debug_callback:
                debug_callback(msg)
        
        debug_print("="*60)
        debug_print(f"开始检查：日期={pub_date_str}, 标题={title[:50]}...")
        debug_print("="*60)
        
        # 打开检索页面
        debug_print("步骤1: 打开检索页面...")
        ok = open_page_with_retry(driver, BASE_SEARCH_URL)
        if not ok:
            debug_print("✗ 页面多次重试仍失败，跳过该行")
            return False
        debug_print("✓ 页面打开成功")
        time.sleep(2)
        
        # 1. 点击时间选择器并选择日期
        debug_print("\n步骤2: 选择日期...")
        if not select_date_by_click(driver, pub_date_str, debug_callback):
            debug_print(f"✗ 无法选择日期：{pub_date_str}")
            return False
        debug_print("✓ 日期选择完成")
        
        # 2. 不再输入标题检索（容易误点到登录框/被遮罩），改为按日期筛选后直接在结果列表分页查找
        debug_print("\n步骤3: 跳过输入框检索（按日期筛选后直接分页查找标题）...")

        # 3. 在结果列表中查找完全匹配的标题（处理分页）
        debug_print("\n步骤4: 在结果中查找标题（分页 + Ctrl+F思路）...")
        found = find_title_in_results(driver, title, debug_callback=debug_callback)
        
        if found:
            debug_print("\n" + "="*60)
            debug_print("✓✓✓ 检查结果：找到匹配的标题！")
            debug_print("="*60)
        else:
            debug_print("\n" + "="*60)
            debug_print("✗✗✗ 检查结果：未找到匹配的标题")
            debug_print("="*60)
        
        return found
        
    except Exception as e:
        error_msg = f"检查标题时出错：{e}"
        print(f"[检查标题] {error_msg}")
        if debug_callback:
            debug_callback(error_msg)
        import traceback
        print(traceback.format_exc())
        return False


# ======== 处理整个 Excel：逐行校验 ========
def process_excel(filepath, report_widget):
    try:
        # 支持多种 Excel 格式：.xlsx (新版), .xls (旧版), WPS 格式
        file_ext = os.path.splitext(filepath)[1].lower()
        
        # 根据文件扩展名选择引擎
        if file_ext == '.xls':
            # 旧版 Excel，使用 xlrd 引擎
            try:
                df = pd.read_excel(filepath, engine='xlrd')
            except ImportError:
                messagebox.showerror("缺少依赖", "读取 .xls 文件需要 xlrd 库，请安装：pip install xlrd<2.0")
                report_widget.insert(tk.END, "错误：缺少 xlrd 库（用于读取旧版 .xls 文件）\n")
                return
        elif file_ext in ['.xlsx', '.xlsm']:
            # 新版 Excel，使用 openpyxl 引擎（也支持 WPS 保存的 .xlsx）
            try:
                df = pd.read_excel(filepath, engine='openpyxl')
            except ImportError:
                messagebox.showerror("缺少依赖", "读取 .xlsx 文件需要 openpyxl 库，请安装：pip install openpyxl")
                report_widget.insert(tk.END, "错误：缺少 openpyxl 库（用于读取新版 .xlsx 文件）\n")
                return
        else:
            # 尝试自动检测（pandas 会自动选择引擎）
            df = pd.read_excel(filepath)
            
    except Exception as e:
        error_msg = f"读取 Excel 文件失败：{str(e)}"
        messagebox.showerror("读取 Excel 出错", error_msg)
        report_widget.insert(tk.END, f"错误：{error_msg}\n")
        report_widget.insert(tk.END, f"支持格式：.xlsx（新版Excel/WPS）、.xls（旧版Excel）\n")
        return
    
    # 检查必备列名
    required_cols = ["发布时间", "标题"]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        error_msg = f"Excel 必须包含列：{required_cols}，缺少：{missing_cols}"
        messagebox.showerror("格式错误", error_msg)
        report_widget.insert(tk.END, f"错误：{error_msg}\n")
        return
    
    # 更新GUI显示
    report_widget.insert(tk.END, f"开始处理文件：{os.path.basename(filepath)}\n")
    report_widget.insert(tk.END, f"共 {len(df)} 行数据需要校验\n")
    report_widget.insert(tk.END, "正在启动浏览器...\n")
    report_widget.update()
    
    # 启动浏览器
    driver = make_driver()
    
    try:
        # 先打开一次页面，让用户看到浏览器窗口
        open_page_with_retry(driver, BASE_SEARCH_URL)
        time.sleep(2)
        
        errors = []
        total_rows = len(df)
        
        # 遍历每一行
        for idx, row in df.iterrows():
            pub_date = row["发布时间"]
            title = row["标题"]
            
            # 格式化日期为字符串
            if pd.isna(pub_date):
                report_widget.insert(tk.END, f"第 {idx + 2} 行：发布时间为空，跳过\n")
                report_widget.update()
                continue
            
            if pd.isna(title) or not isinstance(title, str) or not title.strip():
                report_widget.insert(tk.END, f"第 {idx + 2} 行：标题为空，跳过\n")
                report_widget.update()
                continue
            
            # 转换日期格式
            if isinstance(pub_date, (pd.Timestamp, datetime)):
                pub_date_str = pub_date.strftime("%Y-%m-%d")
            else:
                # 尝试解析字符串日期
                try:
                    pub_date_str = str(pub_date).strip().replace('/', '-')
                    # 验证日期格式
                    datetime.strptime(pub_date_str, '%Y-%m-%d')
                except:
                    report_widget.insert(tk.END, f"第 {idx + 2} 行：日期格式错误 - {pub_date}\n")
                    report_widget.update()
                    continue
            
            # 更新进度
            progress = f"正在检查第 {idx + 2}/{total_rows + 1} 行：{title[:30]}..."
            report_widget.insert(tk.END, f"\n{progress}\n")
            report_widget.see(tk.END)
            report_widget.update()
            
            # 定义调试回调函数，将调试信息输出到GUI
            def debug_to_gui(msg):
                report_widget.insert(tk.END, f"  {msg}\n")
                report_widget.see(tk.END)
                report_widget.update()
            
            # 检查标题是否能在该日期下找到
            found = check_title_at_date(driver, pub_date_str, title, debug_callback=debug_to_gui)
            
            if not found:
                excel_row_num = idx + 2  # Excel行号（第1行是表头）
                errors.append(excel_row_num)
                error_msg = f"第 {excel_row_num} 行可能有问题：标题与发布时间不匹配"
                print(f"问题行：Excel 第 {excel_row_num} 行标题与发布时间可能不匹配")
                report_widget.insert(tk.END, f"  ⚠ {error_msg}\n")
                report_widget.see(tk.END)
                report_widget.update()
            else:
                report_widget.insert(tk.END, f"  ✓ 第 {idx + 2} 行匹配\n")
                report_widget.see(tk.END)
                report_widget.update()
        
        # 关闭浏览器
        driver.quit()
        
        # 在控制台打印所有问题行
        print("\n" + "="*50)
        if errors:
            print(f"共发现 {len(errors)} 行可能有问题：")
            for r in errors:
                print(f"  问题行：Excel 第 {r} 行")
        else:
            print("所有行看起来都匹配 ✓")
        print("="*50)
        
        # 在GUI文本框里输出最终结果
        report_widget.insert(tk.END, "\n" + "="*50 + "\n")
        report_widget.insert(tk.END, "校验完成！\n")
        if errors:
            report_widget.insert(tk.END, f"共发现 {len(errors)} 行可能有问题：\n")
            for r in errors:
                report_widget.insert(tk.END, f"  第 {r} 行\n")
        else:
            report_widget.insert(tk.END, "所有行看起来都匹配 ✓\n")
        report_widget.see(tk.END)
        report_widget.update()
        
    except Exception as e:
        error_msg = f"处理过程中出错：{str(e)}"
        print(error_msg)
        report_widget.insert(tk.END, f"\n错误：{error_msg}\n")
        report_widget.update()
        driver.quit()


# ======== GUI 部分：文件选择 + 报错窗口 ========
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("知网 Excel 校验工具")
        self.geometry("800x600")
        
        # 标题
        title_label = tk.Label(
            self,
            text="知网文章标题与发布时间校验工具",
            font=("Arial", 14, "bold")
        )
        title_label.pack(pady=10)
        
        # 文件选择按钮
        self.select_button = tk.Button(
            self,
            text="选择 Excel 文件开始校验\n\n文件需包含列：'发布时间' 和 '标题'",
            command=self.select_file,
            width=60,
            height=5,
            bg="#f0f0f0",
            font=("Arial", 10)
        )
        self.select_button.pack(pady=20)
        
        # 报错显示窗口
        report_label = tk.Label(self, text="校验结果：", font=("Arial", 10, "bold"))
        report_label.pack(anchor="w", padx=10)
        
        self.report = scrolledtext.ScrolledText(
            self,
            width=90,
            height=20,
            font=("Consolas", 9)
        )
        self.report.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        # 提示信息
        tip_label = tk.Label(
            self,
            text="提示：校验过程中会自动打开浏览器窗口，请勿关闭。",
            font=("Arial", 8),
            fg="gray"
        )
        tip_label.pack(pady=5)
    
    def select_file(self):
        filepath = filedialog.askopenfilename(
            title="请选择 Excel 文件",
            filetypes=[
                ("Excel 文件", "*.xlsx *.xls"),
                ("所有文件", "*.*")
            ]
        )
        if not filepath:
            return  # 用户取消选择
        
        # 清空旧结果
        self.report.delete("1.0", tk.END)
        self.report.insert(tk.END, f"文件：{os.path.basename(filepath)}\n")
        self.report.insert(tk.END, f"路径：{filepath}\n")
        self.report.insert(tk.END, "-" * 50 + "\n\n")
        self.report.update()
        
        # 开子线程执行，避免卡死界面
        t = threading.Thread(
            target=process_excel,
            args=(filepath, self.report),
            daemon=True
        )
        t.start()


if __name__ == "__main__":
    app = App()
    app.mainloop()

