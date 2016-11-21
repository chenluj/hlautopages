# -*- coding: utf-8 -*-

import os
import yaml
import time
from xlrd import open_workbook
from selenium import webdriver
from selenium.webdriver.support.select import Select


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
        if type(self.sheet_locator) not in [int, str]:
            raise SheetTypeError('Please pass in <type \'int\'> or <type \'str\'>, not {0}'.format(type(self.sheet)))
        elif type(self.sheet_locator) == int:
            try:
                sheet = self.book.sheet_by_index(self.sheet_locator)  # by index
            except:
                raise SheetError('Sheet \'{0}\' not exists.'.format(self.sheet_locator))
        else:
            try:
                sheet = self.book.sheet_by_name(self.sheet_locator)  # by name
            except:
                raise SheetError('Sheet \'{0}\' not exists.'.format(self.sheet_locator))
        return sheet

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


class Browser:

    def __init__(self):
        self.driver = None
        # self.profile = webdriver.FirefoxProfile()

    def get(self, url):
        self.driver = webdriver.Firefox()
        self.driver.get(url)

    def close(self):
        self.driver.close()


class Page:

    def __init__(self, driver=None):
        if driver:
            self.driver = driver
        else:
            self.driver = webdriver.Firefox()

        self.data = None

    def get(self, url):
        self.driver.get(url)

    def element_work(self, work):
        locate_way = work[0]
        locate_expr = work[1]
        action = work[2]
        if len(work) > 4:
            # TODO: 循环读取excel中数据并
            param = work[4]
            xls = ExcelReader().data[0][work[3]]
            # print xls
            param = xls

        element = self.driver.find_element(by=locate_way, value=locate_expr)
        time.sleep(1)

        if action in ACTIONS:
            if action == 'click':
                element.click()
                print u'元素点击'
            elif action == 'clear':
                element.clear()
                print u'元素清空内容'
            elif action == 'sendkeys':
                element.send_keys(param)
                print u'元素输入值  {}'.format(param)
            elif action == 'submit':
                element.submit()
                print u'提交'
            elif action == 'select':
                Select(element).select_by_value(param)
                print u'选择选项  {}'.format(param)
        else:
            print u"Unsupported action: {}.".format(action)


def main():
    tasks = YamlReader().data
    for task in tasks:
        print u'======  任务开始  ======='
        browser = Browser()
        print u'打开浏览器'
        try:
            url = task.pop(0)['url']
            browser.get(url)
            print u'打开网页  {}'.format(url)
        except:
            raise
            print u'task 第一项必须为url'
        else:
            p = Page(driver=browser.driver)
            for page in task:
                for element in page['elements']:
                    print u'定位元素  ',
                    print element
                    p.element_work(element)
                time.sleep(5)

            time.sleep(2)
            browser.close()
            print u'关闭浏览器'
            print u'======  任务结束  ======'
            print


if __name__ == '__main__':
    # ym = YamlReader().data
    # for task in ym:
    #     # print task
    #     for item in task:
    #         print item
    #
    # xls = ExcelReader()
    # print xls.data
    # print xls.title
    # print xls.nums
    main()





