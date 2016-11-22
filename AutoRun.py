# -*- coding: utf-8 -*-

import os
import yaml
import time
from xlrd import open_workbook
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary


CONFIG = 'config.yaml'
DATA = 'data.xlsx'

ACTIONS = ['click', 'clear', 'sendkeys', 'submit', 'select']


class Error(Exception):
    """Base package Exception."""
    pass


class FileException(Error):
    """Base file exception.Thrown when a file is not available.
    For example:
        file not exists.
    """
    pass


class DataFileNotAvailableException(FileException):
    """Thrown when data file not available."""
    pass


class SheetTypeError(Error):
    """Thrown when sheet type passed in not int or str."""
    pass


class SheetError(Error):
    """Thrown when specified sheet not exists."""
    pass


class DataError(Error):
    """Thrown when something wrong with the data."""
    pass


class ExcelReader(object):
    def __init__(self, sheet=0):
        """Read workbook

        :param sheet: index of sheet or sheet name.
        """
        self.book_name = os.path.abspath(DATA)
        self.sheet_locator = sheet

        self.book = self._book()
        self.sheet = self._sheet()

    def _book(self):
        try:
            work_book = open_workbook(self.book_name)
        except IOError as e:
            raise DataFileNotAvailableException(e)
        return work_book

    def _sheet(self):
        """Return sheet"""
        try:
            return self.book.sheet_by_name(self.sheet_locator)  # by name
        except:
            raise SheetError('Sheet \'{0}\' not exists.'.format(self.sheet_locator))

    @property
    def title(self):
        """First row is title."""
        try:
            return self.sheet.row_values(0)
        except IndexError:
            raise DataError('This is a empty sheet, please check your file.')

    @property
    def data(self):
        """Return data in specified type:

            [{row1:row2},{row1:row3},{row1:row4}...]
        """
        sheet = self.sheet
        title = self.title
        data = list()

        # zip title and rows
        for col in range(1, sheet.nrows):
            s1 = sheet.row_values(col)
            s2 = [unicode(s).encode('utf-8') for s in s1]  # utf-8 encoding
            data.append(dict(zip(title, s2)))
        return data

    @property
    def nums(self):
        """Return the number of cases."""
        return len(self.data)


class YamlReader:
    """Read yaml file"""
    def __init__(self):
        self.yaml = os.path.abspath(CONFIG)

    @property
    def data(self):
        with open(self.yaml, 'r') as f:
            al = yaml.safe_load_all(f)
            y = [x for x in al]
            return y


class Config:
    def __init__(self, conf):
        self.browser = conf['browser'].lower() if 'browser' in conf else 'firefox'
        self.location = conf['location'] if 'location' in conf else None
        self.delay_submit = conf['delay_submit'] if 'delay_submit' in conf else 5


class Browser:

    def __init__(self, config):
        self.driver = None
        self.browser = config.browser
        self.location = config.location
        self.delay_submit = config.delay_submit

    def open(self):
        if self.browser == 'firefox':
            try:
                binary = FirefoxBinary(self.location)
                profile = webdriver.FirefoxProfile()
                profile.add_extension(os.path.abspath('random-agent-spoofer.xpi'))
                self.driver = webdriver.Firefox(firefox_binary=binary, firefox_profile=profile)
                self.driver.implicitly_wait(30)
                print u'[Info] 打开浏览器  firefox'
                return self
            except:
                print u'[Error] 打开firefox 浏览器失败 请检查浏览器路径配置以及random-agent-spoofer.xpi插件'
                os._exit(0)
        elif self.browser == 'chrome':
            try:
                option = webdriver.ChromeOptions()
                option.binary_location = self.location

                self.driver = webdriver.Chrome(executable_path='chromedriver.exe', chrome_options=option)
                self.driver.implicitly_wait(30)
                print u'[Info] 打开浏览器  chrome'
                return self
            except:
                print u'[Error] 打开chrome浏览器失败 请检查浏览器路径配置以及chromedriver.exe驱动'
                os._exit(0)
        else:
            print u'[Error] 不支持的浏览器类型'
            os._exit(0)

    def get(self, url):
        try:
            self.driver.get(url)
            print u'[Info] 打开URL  {}'.format(url)
            return self.driver
        except:
            print u'[Error] 打开URL失败，请检查配置'
            os._exit(0)

    def quit(self):
        self.driver.quit()
        print u'[Info] 关闭浏览器'


class Element:
    def __init__(self, driver, elem_info, params):
        try:
            locator = (elem_info[0], elem_info[1])
            self.element = WebDriverWait(driver, 15, 0.5).until(presence_of_element_located(locator))
            self.action = elem_info[2].lower()
            self.element_name = elem_info[3]
            self.params = params
            print u'[Info] 元素已找到  ',
            print elem_info
        except TimeoutException:
            print u'[Error] 未找到元素  ',
            print elem_info

    def do_its_work(self, delay_submit):
        if self.element:
            if self.action == 'click':
                self.element.click()
            elif self.action == 'clear':
                self.element.clear()
            elif self.action == 'submit':
                time.sleep(delay_submit)
                self.element.submit()
            elif self.action == 'sendkeys':
                self.element.send_keys(self.pick_value())
            elif self.action == 'select':
                Select(self.element).select_by_value(self.pick_value())
            else:
                print u"[Error] 不支持的action {}".format(self.action)

    def pick_value(self):
        value = self.params[self.element_name]
        print u'[Info] 从Excel中取得值 {}'.format(value)
        return value


class Task:
    def __init__(self, task):
        self.url = task.pop(0)['url']
        self.sheet = task.pop(0)['sheet']
        xls = ExcelReader(sheet=self.sheet)
        self.loop_times = xls.nums
        self.data = xls.data

        self.task = task

    def run(self, b):
        for t in range(self.loop_times):
            params = self.data[t]

            print u'======  任务开始  ======='
            driver = b.open().get(self.url)
            for page in self.task:
                for element in page['elements']:
                    if isinstance(element, dict):
                        time.sleep(3)
                        if driver.current_url == element['if']:
                            print u'[Info] URL为期待值，任务成功'
                            break
                    else:
                        Element(driver, element, params).do_its_work(b.delay_submit)
                        time.sleep(1)
                time.sleep(5)
            b.quit()
            print u'======  任务结束  ======='
            print


def main():
    try:
        tasks = YamlReader().data
        conf = Config(tasks.pop(0))
    except:
        print u'[Error] 读取配置文件出错'
    else:
        browser = Browser(conf)
        for task in tasks:
            print u'[Info] 执行任务  ',
            print task
            try:
                t = Task(task)
            except:
                print u'[Error] 初始化任务出错，请检查配置或数据文件，确认填写无误并且变量名与列名对应'
                os._exit(0)
            else:
                try:
                    t.run(browser)
                except:
                    print u'[Error] 执行任务出错，请检查配置与页面是否对应'
                    browser.quit()
                    os._exit(0)

        print u'[Info] 所有任务执行结束，请处理数据后重新启动程序'


if __name__ == '__main__':
    main()





