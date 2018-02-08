'''
    1.我们先创建一个简单的脚本吧，在test文件夹创建test_baidu.py：
    2.多次搜索增加unittest
    3.导入config配置文件，获取配置文件路径
    4.导入日志配置文件
    5.读取excel配置文件
    6.引入报告生成文件
    7.引入发送邮件模块
'''
__author__ = 'ingshen'
import sys
sys.path.append('..')
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import unittest
from utils.config import Config, DRIVER_PATH, BASE_PATH, REPORT_PATH
from utils.log import logger
from utils.file_reader import ExcelReader
from utils.HTMLTestRunner_PY3 import HTMLTestRunner
from utils.mail import Email

class TestBaiDu(unittest.TestCase):
    URL = Config().get('URL')
    excel = os.path.join(BASE_PATH,'data','baidu.xlsx')

    locator_kw = (By.ID, 'kw')
    locator_su = (By.ID, 'su')
    locator_result = (By.XPATH, '//div[contains(@class, "result")]/h3/a')

    def sub_setUp(self):
        self.driver = webdriver.Chrome()
        self.driver.get(self.URL)

    def sub_tearDown(self):
        self.driver.quit()
    
    def test_search(self):
        datas = ExcelReader(self.excel).data
        for d in datas:
            with self.subTest(data=d):
                self.sub_setUp()
                self.driver.find_element(*self.locator_kw).send_keys(d['search'])
                self.driver.find_element(*self.locator_su).click()
                time.sleep(2)
                links = self.driver.find_elements(*self.locator_result)
                for link in links:
                    logger.info(link.text)
                self.sub_tearDown()
    
if __name__ == '__main__':
    report = os.path.join(REPORT_PATH,'report.html')
    with open(report,'wb') as f:
        runner = HTMLTestRunner(f, verbosity=2, title='从0开始搭建测试框架', description='修改html报告')
        runner.run(TestBaiDu('test_search'))
    
    e = Email(title='百度搜索报告',
               message='这是今天的报告，请查收！',
               receiver='收件人邮箱',
               server='发件人smtp服务器',
               sender='发件人邮箱',
               password='发件人邮箱密码',
               path=report
               )
    e.send()

