from playwright.sync_api import Playwright, sync_playwright, expect
import re
import pandas as pd
from datetime import date
from dotenv import load_dotenv
import os
# 加载 .env 文件
load_dotenv()

def load_env():
    return os.getenv('id'),os.getenv('key'),os.getenv('app_name'),os.getenv('type')
def run(playwright: Playwright):
    data = []
    id_,key,app_name,day = load_env()
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://coze.nankai.edu.cn/login?redirect_uri=%2Fproduct%2Fllm%2Fpersonal%2Fpersonal-411%2Fapplication")
    page.get_by_text("SSO登录").click()
    page.get_by_role("textbox", name="请输入学号/工号").click()
    page.get_by_role("textbox", name="请输入学号/工号").fill(id_)
    page.get_by_role("textbox", name="请输入密码").click()
    page.get_by_role("textbox", name="请输入密码").fill(key)
    page.locator(".arco-checkbox-mask").first.click()
    page.get_by_role("button", name="登 录").click()
    page.get_by_text(app_name, exact=True).click()
    page.get_by_text("统计").click()
    page.get_by_text("日志和标注").click()
    page.get_by_text(f"近 {day} 天").click()
    page.get_by_text("条/页").click()
    page.get_by_role("option", name="100 条/页").click()
    # 定位所有符合条件的 <tr> 元素
    tr_elements = page.locator("tr[data-testid='c-m-table-row']")
    # 遍历每个 <tr> 元素
    for i in range(tr_elements.count()):
        tr = tr_elements.nth(i)
        span_text = tr.locator("td:nth-child(3) > .arco-table-cell > .arco-table-cell-wrap-value > .c-m-ellipsis").text_content()

        # 检查 <span> 元素的文本内容是否为 "0"
        if span_text.strip() != "0":
            # 点击 <div class="right"> 元素
            tr.locator("div.right").click()

            # 等待 <div class="cursor-pointer"> 元素出现
            page.wait_for_selector("div.cursor-pointer")

            # 获取所有 <div class="cursor-pointer"> 元素的文本内容并打印
            cursor_pointer_elements = page.locator("div.cursor-pointer")
            for j in range(cursor_pointer_elements.count()):
                text_content = cursor_pointer_elements.nth(j).text_content()
                data.append(text_content)

            # 点击关闭按钮 <span class="arco-icon-hover arco-drawer-close-icon">
            page.locator("span.arco-icon-hover.arco-drawer-close-icon").click()

    # ---------------------
    context.close()
    browser.close()
    return data

def save_data(data):
    data2 = {
        '问题': [],
        '回答': []
    }
    for da in data:
        part = re.split('回答：', da)
        question = part[0].replace('问题：', '')
        answer = part[1]
        data2['问题'].append(question)
        data2['回答'].append(answer)

    df = pd.DataFrame(data2)
    today = date.today().strftime('%Y-%m-%d')
    app_name = os.getenv('app_name')
    df.to_excel(f'{app_name}-对话导出-{today}.xlsx', index=False)

with sync_playwright() as playwright:
    data = run(playwright)

