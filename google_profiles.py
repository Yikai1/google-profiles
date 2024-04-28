import os
import threading
from enum import Enum
from math import log

import pandas as pd
import pyautogui
import pyperclip
import re
import random
import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import tkinter as tk
from tkinter import ttk

class Errors(Enum):
    SUCCESS = '成功'
    SERVER_ERROR = '服务器错误'


class Scholar:
    def __init__(self, out_filepath) -> None:
        self.out_filepath = out_filepath
        if not os.path.exists(self.out_filepath):
            os.mkdir(self.out_filepath)
        self.driver = None
        self.results = []

    def start_browser(self, wait_time=10):
        # 创建ChromeOptions对象
        options = Options()
        # 启用无头模式
        # options.add_argument("--headless")
        # 启用无痕模式
        options.add_argument("--incognito")
        options.add_argument("--disable-domain-reliability")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-client-side-phishing-detection")
        options.add_argument("--no-first-run")
        options.add_argument("--use-fake-device-for-media-stream")
        options.add_argument("--autoplay-policy=user-gesture-required")
        options.add_argument("--disable-features=ScriptStreaming")
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-save-password-bubble")
        options.add_argument("--mute-audio")
        options.add_argument("--no-sandbox")
        #options.add_argument("--disable-gpu")
        options.add_argument("--disable-extensions")
        options.add_argument("--disable-software-rasterizer")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-webgl")
        options.add_argument("--allow-running-insecure-content")
        options.add_argument("--no-default-browser-check")
        options.add_argument("--disable-full-form-autofill-ios")
        options.add_argument("--disable-autofill-keyboard-accessory-view[8]")
        options.add_argument("--disable-single-click-autofill")
        options.add_argument("--ignore-certificate-errors")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-blink-features")
        # 禁用实验性QUIC协议
        options.add_experimental_option("excludeSwitches", ["enable-quic"])
        # 禁用日志输出
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        # 创建Chrome浏览器实例
        self.driver = webdriver.Chrome(options=options)
        # 等待页面加载完成
        self.driver.implicitly_wait(wait_time)
    
           
    def search_onepage(self, filter_condition):
        """爬取当前页面文章的的信息"""
        results = []
                            
        gsc_1usr = self.driver.find_element(by=By.ID, value='gs_bdy').find_elements(by=By.CLASS_NAME, value='gsc_1usr')
        
        for i, item in (enumerate(gsc_1usr)):
           
            # 提取作者、作者机构和引用数等信息
            gs_ai_name = item.find_element(by=By.CLASS_NAME, value='gs_ai_name')            
            gs_ai_name_a = gs_ai_name.find_element(by=By.TAG_NAME, value='a') if self.check_element_exist(check_type='TAG_NAME', value='a', source=gs_ai_name.get_attribute('innerHTML')) else None
            gs_ai_aff = item.find_element(by=By.CLASS_NAME, value='gs_ai_aff')
            gs_ai_eml = item.find_element(by=By.CLASS_NAME,value='gs_ai_eml')
            gs_ai_cby = item.find_element(by=By.CLASS_NAME, value='gs_ai_cby')

            #学者
            #print(gs_ai_name)
            authors = gs_ai_name.text.strip()                        
            # #链接
            authors_link = gs_ai_name_a.get_attribute('href') if gs_ai_name_a else ''         
            # #单位
            affiliations = gs_ai_aff.text.strip()            
            #邮箱
            eml = gs_ai_eml.text.strip().strip('在').strip('经过验证').replace(' 的电子邮件','')
            #引用量
            cited_by = gs_ai_cby.text.strip().strip('被引用次数：')
            
            #print(f'[{i}] {authors} => {authors_link} => {affiliations} => {eml} => {cited_by}')
            
            pattern = filter_condition #检测邮箱
            #如果为空则不进行正则化
            if pattern:
                if re.search(pattern,eml):
                    results.append({
                        'authors': authors,
                        'link': authors_link,
                        'affiliations': affiliations,
                        'email': eml,
                        'cited_by': cited_by
                        })
            else:
                results.append({
                        'authors': authors,
                        'link': authors_link,
                        'affiliations': affiliations,
                        'email': eml,
                        'cited_by': cited_by
                        })
        return results


    def check_element_exist(self, value, check_type='CLASS_NAME', source=None) -> bool:
        """检查页面是否存在指定元素"""
        page_source = source if source else self.driver.page_source
        soup = BeautifulSoup(page_source, 'lxml')
        if check_type == 'ID':
            return len(soup.find_all(id=value)) != 0
        elif check_type == 'CLASS_NAME':
            return len(soup.find_all(class_=value)) != 0
        elif check_type == 'TAG_NAME':
            return len(soup.find_all(value)) != 0
        elif check_type == 'FULL':
            return value in page_source
        else:
            pyautogui.confirm(text=f'检查条件[{check_type}]不对',title='错误')           
            #print(f'>> 检查条件[{check_type}]不对')
        return False


    def check_captcha(self) -> bool:
        """检查是否需要人机验证；一个是谷歌学术的、一个是谷歌搜索的"""
        return self.check_element_exist(check_type='ID', value='gs_captcha_f') or \
               self.check_element_exist(check_type='ID', value='captcha-form')


    def process_error(self, error: Errors) -> bool:
        """尽可能尝试解决错误"""
        success = False
        if error == Errors.SERVER_ERROR:
            #遇到问题刷新
            self.driver.refresh()            
            success = True
        return success
    

    def check_error(self, try_solve = True) -> Errors:
        """检查当前页面是否出错"""
        error = Errors.SUCCESS
        if self.check_element_exist(check_type='FULL', value='服务器错误'):
            error = Errors.SERVER_ERROR
        
        # 尝试解决错误
        if try_solve and error != Errors.SUCCESS:
            error = Errors.SUCCESS if self.process_error(error) else error
        return error


    def search(self, url, filter_condition, max_pages=10000, min_cited=0, delay=5, filename='scholars.xlsx'):        
        global page, current_cited

        self.driver.get(url)        

        cited_by_list = []  # 初始化引用数列表        
                
        for page in range(1, max_pages+1):            
            #检测问题
            #保存地址以继续搜索
            currentPageUrl = self.driver.current_url
            urlname = 'current_url.txt'  # 文件名
            with open(urlname, 'w') as file:
                file.write(currentPageUrl)
            #人机验证
            while self.check_captcha():
                if pyautogui.alert(title='状态异常', text='请手动完成人机验证后，点击“已完成”', button='已完成')=='已完成':
                    self.driver.refresh()
                    time.sleep(2)
            
            #报错
            if self.check_error() != Errors.SUCCESS:
                if pyautogui.confirm(text='请检查页面出现了什么问题;\n解决后,点击“重试”;\n否则,点击“取消”提前结束脚本;', title='状态异常', buttons=['重试', '取消']) == '重试':
                    self.driver.refresh()
                else:    
                    break
                time.sleep(2)            
            
            #遇到网络问题或者网页更新不及时的问题自动刷新5次，若未解决弹出提示框
            MAX_REFRESH_COUNT = 5
            refresh_count = 0
            while not self.check_element_exist(check_type='ID', value='gsc_sa_ccl'):              
                if refresh_count < MAX_REFRESH_COUNT:
                    self.driver.refresh() 
                    refresh_count += 1
                if refresh_count >= MAX_REFRESH_COUNT:
                    if pyautogui.confirm(text=f'当前页面为空,若页面已加载完成,请点击跳过；若页面没有加载,请检测网络连接后点击重试',title='错误',buttons=['重试','跳过'])=='重试':
                        time.sleep(2)
                        self.driver.refresh()   
                        continue
                    else:
                        continue           
            
            #开始搜索
            onepage = self.search_onepage(filter_condition=filter_condition)           
            self.results.extend(onepage)
            
            # 保存当前页的结果
            self.save_file(filename=filename)
            
            cited_by_list = [int(i['cited_by']) for i in self.results] # 获取当前已经爬取的文章的引用数列表
            current_cited = int(min(cited_by_list))  # 当前最小引用量
            max_cited = int(max(cited_by_list))
            
            # 设置进度条的最大值
            bar_max = max_cited-min_cited
            bar_current = max_cited-current_cited
            
            # progress_bar["maximum"] = bar_max 
            stage1_max = round((bar_max) * 0.1)  # 初始的10%引文量
            stage2_max = bar_max - stage1_max  # 剩余的70%引文量
            progress_bar["maximum"] = bar_max  # 设置进度条的最大值
            
            #窗口更新说明
            search_label.config(text=f'正在搜索第{page}页，当前的引用量为{current_cited}')
            
            #分段更新进度条
            if bar_current > stage2_max:
                # 如果在阶段1
                progress = bar_current / stage1_max * stage1_max
            else:
                # 如果在阶段2
                progress = stage1_max + (bar_current - stage2_max) / stage2_max * stage2_max
            
            progress_bar["value"] = progress  # 更新进度条的值
            window.update_idletasks()  # 更新窗口组件

            #点击停止按钮
            if not is_searching:
                break                        
            
            # 爬取完成
            if current_cited <= min_cited:                
                break
            
            # 当前页为最后一页
            if not self.driver.find_element(by=By.CLASS_NAME, value='gs_btnPR').is_enabled():
                break
            
            self.driver.find_element(by=By.CLASS_NAME, value="gs_btnPR").click()
            time.sleep(delay) #必须睡，后面网页加载太快会连接不上网页           
                        
        return self.results


    def save_file(self, filename='scholars.xlsx'):
        unique_data = pd.DataFrame(self.results).dropna().reset_index(drop=True)
        filepath = os.path.join(self.out_filepath, filename)
    
        try:                
            # 设置authors_link为超链接
            unique_data['link'] = unique_data['link'].apply(lambda x: f'=HYPERLINK("{x}", "{x}")')            
            # 将cited_by的内容转换为数字
            unique_data['cited_by'] = pd.to_numeric(unique_data['cited_by'], errors='coerce')
            # 保存文件
            unique_data.to_excel(filepath, index=False)
        
        except Exception as e:            
            if pyautogui.confirm(text=f'文件保存失败[{str(e)}]\n点击“确定”将内容复制到剪切板', title='文件保存失败', buttons=['确定','取消']) == '确定':
                pyperclip.copy(str(unique_data))


    def close_browser(self):
        # 关闭浏览器
        self.driver.quit()
        os.system('taskkill /im Coogle Chrome /F')

    
   

if __name__ == '__main__':
    #创建窗口对象
    window = tk.Tk()
    window.title('谷歌档案馆')
    #自定义窗口大小
    window.geometry("450x350")  # 设置宽度为450，高度为350
    
    #创建三个标签和对应的输入框
    tk.Label(window, text='关键词').grid(row=0, column=0, padx=(40,10), pady=5)
    e1 = tk.Entry(window)  # 关键词输入框
    e1.grid(row=0, column=1,padx=10, pady=5)  
    tk.Label(window, text='网页搜索方式示例"computer_vision"').grid(row=1, column=1, padx=10, pady=5)
 
    tk.Label(window, text='最小引用量').grid(row=2, column=0, padx=(40,10), pady=5)    
    e2 = tk.Entry(window)  # 最大页数输入框  
    e2.grid(row=2, column=1,padx=10, pady=5)
    tk.Label(window, text='引用量越小耗时越长').grid(row=3, column=1, padx=10, pady=5)
    
    tk.Label(window, text='筛选条件').grid(row=4, column=0, padx=(40,10), pady=5)    
    e3 = tk.Entry(window)  # 筛选条件输入框  
    e3.grid(row=4, column=1,padx=10, pady=5)
    tk.Label(window, text='用分号作为分隔,示例:edu;ac.hk').grid(row=5, column=1, padx=10, pady=5)
    
    #创建搜索、退出和继续按钮
    btn_search = tk.Button(window, text='搜索')
    btn_search.grid(row=6, column=0, padx=(50, 10), pady=5)
    
    btn_continue = tk.Button(window, text='继续搜索')
    btn_continue.grid(row=6, column=1, padx=(10, 110), pady=5)
    
    btn_stop = tk.Button(window, text='停止搜索')
    btn_stop.grid(row=6, column=1, padx=(200,10) , pady=5)

    #创建进度条小部件
    tk.Label(window, text='进度').grid(row=7, column=0, padx=(40,10), pady=5)
    progress_bar = ttk.Progressbar(window, mode='determinate', length=200)
    progress_bar.grid(row=7, column=1, columnspan=2, padx=10, pady=5)   
    
    #创建搜索进度标签
    search_label = tk.Label(window, text='')
    search_label.grid(row=8, column=0, columnspan=3, padx=10, pady=5)

    is_searching = False
    should_stop_search = False

    def exit_program():
        global is_searching ,should_stop_search
        is_searching = False
        should_stop_search = True
    
    def on_enter_key_press(event):
        #检测回车键
        if event.keysym == "Return":
            search()
    
    def search():
        global is_searching
        is_searching = True
        
        # 获取第三个输入框中的信息
        keywords = e1.get().strip()
        min_cited = e2.get().strip()        
        filter_conditions = e3.get().strip()

        # 删除空格和将";"替换为"|" 
        filter_conditions = re.sub(r'\s+', '', filter_conditions)
        filter_conditions = re.sub(r';', '|', filter_conditions)
        filter_conditions = re.sub(r'；', '|', filter_conditions)
       
        # 对应到谷歌学术个人档案网址
        keywords = keywords.replace(' ', '+')        
        url = f'https://scholar.google.com/citations?hl=zh-CN&view_op=search_authors&mauthors=label:{keywords}&btnG='
        
        # 初始化进度条
        progress_bar["value"] = 0
        progress_bar.update()

        #设置保存地址
        filename='scholars.xlsx'
        out_filepath = '_'.join(keywords.replace('"', '').replace(':', '').split())
        filepath = os.path.join(out_filepath, filename)        
        # 检查文件是否存在
        if os.path.exists(filepath):
            # 获取文件名和扩展名
            name, ext = os.path.splitext(filename)
            i = 1
            # 在文件名后面加上"(数字)"来创建新的文件名
            while os.path.exists(os.path.join(out_filepath, f"{name} ({i}){ext}")):
                i += 1
            filename = f"{name} ({i}){ext}"
                
        def run_search():            
            global is_searching
            # 更新搜索提示信息为"正在搜索，请稍候..."                        
            scholar = Scholar('_'.join(keywords.replace('"', '').replace(':', '').split()))
            scholar.start_browser(wait_time=60)
            scholar.search(url, min_cited=int(min_cited), delay=random.randint(1, 5), filter_condition=filter_conditions, filename=filename)                       
            scholar.close_browser()
            
            if is_searching:
                # 更新搜索提示信息为"搜索完成！"
                search_label.config(text='搜索完成! Execl 文件保存在代码目录下 Keyword 同名文件夹中')
                is_searching = False
            else:
                # 提示搜索已停止
                search_label.config(text='搜索已停止')
        
        # 创建并启动新线程执行搜索
        search_thread = threading.Thread(target=run_search)
        search_thread.start()
    
    def continue_search():
        global is_searching
        is_searching = True
        
        # 获取上一次中断的网址
        urlname = 'Current_url.txt'  
        with open(urlname, 'r') as file:
            url = file.read()
        
        # 获取第三个输入框中的信息
        keywords = e1.get().strip()
        min_cited = e2.get().strip()        
        filter_conditions = e3.get().strip()

        # 删除空格和将";"替换为"|" 
        filter_conditions = re.sub(r'\s+', '', filter_conditions)
        filter_conditions = re.sub(r';', '|', filter_conditions)
        filter_conditions = re.sub(r'；', '|', filter_conditions)    
        
        keywords = keywords.replace(' ', '+')
        
        # 初始化进度条
        progress_bar["value"] = 0
        progress_bar.update()

        #设置保存地址
        filename='scholars.xlsx'
        out_filepath = '_'.join(keywords.replace('"', '').replace(':', '').split())
        filepath = os.path.join(out_filepath, filename)        
        # 检查文件是否存在
        if os.path.exists(filepath):
            # 获取文件名和扩展名
            name, ext = os.path.splitext(filename)
            i = 1
            # 在文件名后面加上"(数字)"来创建新的文件名
            while os.path.exists(os.path.join(out_filepath, f"{name} ({i}){ext}")):
                i += 1
            filename = f"{name} ({i}){ext}"
        
        def run_search():            
            global is_searching
            # 更新搜索提示信息为"正在搜索，请稍候..."                        
            scholar = Scholar('_'.join(keywords.replace('"', '').replace(':', '').split()))
            scholar.start_browser(wait_time=60)
            scholar.search(url, min_cited=int(min_cited), delay=random.randint(1, 5), filter_condition=filter_conditions, filename=filename)                       
            scholar.close_browser()
            
            if is_searching:
                # 更新搜索提示信息为"搜索完成！"
                search_label.config(text='搜索完成! Execl 文件保存在代码目录下 Keyword 同名文件夹中')
                is_searching = False
            else:
                # 提示搜索已停止
                search_label.config(text='搜索已停止')
        
        # 创建并启动新线程执行搜索
        search_thread = threading.Thread(target=run_search)
        search_thread.start()  
    
    # 将函数绑定到按钮上
    btn_search['command'] = search
    btn_continue['command'] = continue_search
    btn_stop['command'] = exit_program
    
    # 绑定窗口的 <Key> 事件到 on_enter_key_press 函数
    window.bind("<Key>", on_enter_key_press)    
    window.mainloop()
