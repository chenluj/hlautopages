# -*- coding: utf-8 -*-

import os
import time
import random
from xlrd import open_workbook
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located, visibility_of_element_located
from selenium.common.exceptions import StaleElementReferenceException, ElementNotVisibleException

TASK_URL = 'http://whoer.net'
DATA = u'数据.xls'
RANDOM_AGENT_SPOOFER = 'random-agent-spoofer.xpi'
PROXYTOOL = 'C:\\Users\\Z\\Desktop\\work\\"911S5 2.0 2016-07-12"\\ProxyTool\\AutoProxyTool.exe -changeproxy/GB'


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


class SheetError(Error):
    """Thrown when specified sheet not exists."""
    pass


class ExcelReader(object):
    def __init__(self, sheet=0):
        """Read workbook

        :param sheet: index of sheet.
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
            return self.book.sheet_by_index(self.sheet_locator)  # by name
        except:
            raise SheetError('Sheet \'{0}\' not exists.'.format(self.sheet_locator))

    @property
    def data(self):
        sheet = self.sheet
        data = list()

        for col in range(0, sheet.nrows):
            s1 = sheet.row_values(col)
            s2 = [unicode(s).encode('utf-8') for s in s1]  # utf-8 encoding
            data.append(s2)
        return data

    @property
    def nums(self):
        """Return the number of cases."""
        return len(self.data)


class Browser:
    def __init__(self):
        self.driver = None

    def open(self):
        profile = webdriver.FirefoxProfile()
        profile.add_extension(os.path.abspath(RANDOM_AGENT_SPOOFER))
        self.driver = webdriver.Firefox(firefox_profile=profile)
        # self.driver.implicitly_wait(30)
        return self

    def get(self, url):
        self.driver.get(url)
        return self.driver

    def quit(self):
        self.driver.quit()


class Task:
    def __init__(self, data):
        self.data = data
        self.b = Browser()

    def change_proxy(self):
        os.system(PROXYTOOL)
        time.sleep(3)


    def run(self):
        driver = self.b.open().get(TASK_URL)

        sex = driver.find_elements_by_xpath('//div[@class="bg-radio"]')
        random.choice(sex).click()

        driver.find_element_by_id('firstname').send_keys(self.data[0])
        driver.find_element_by_id('lastname').send_keys(self.data[1])
        driver.find_element_by_id('email').send_keys(self.data[2])
        Select(driver.find_element_by_id('dobday')).select_by_index(random.choice(range(1, 28)))
        Select(driver.find_element_by_id('dobmonth')).select_by_index(random.choice(range(1, 11)))
        Select(driver.find_element_by_id('dobyear')).select_by_index(random.choice(range(1, 20)))

        driver.find_element_by_id('phone').send_keys(self.data[3])
        driver.find_element_by_id('postcode').send_keys(self.data[4])

        driver.find_element_by_id('searchpostcode').click()
        time.sleep(2)

        Select(driver.find_element_by_id('selectedaddress')).select_by_index(random.choice(range(1, 10)))
        driver.find_element_by_xpath('//label[@class="label-terms"]').click()

        while True:
            try:
                driver.find_element_by_id('singlebutton').click()
            except:
                pass
            time.sleep(5)
            try:
                WebDriverWait(driver, 50, 0.5).until(presence_of_element_located(('id', 'dispoffer')))
                break
            except:
                pass

        for i in range(10):
            time.sleep(1)
            if driver.find_element_by_id('overlay').is_displayed():
                break
            print i

            for div in driver.find_elements_by_xpath('//div[@class="displayedpage"]//div[@class="block-flat"]'):
                try:
                    boxcovers = div.find_elements_by_xpath('.//span[@class="box_cover "]')
                    if boxcovers:
                        for box in boxcovers:
                            radio = box.find_elements_by_xpath('.//label[@class="label_check circle-ticked"]')
                            check = box.find_elements_by_xpath('.//label[@class="label_check default"]')
                            tick = box.find_elements_by_xpath('.//label[@class="label_check ticked"]')

                            if radio:
                                try:
                                    random.choice(radio).click()
                                except:
                                    pass
                            if check:
                                try:
                                    random.choice(check).click()
                                except:
                                    pass
                            if tick:
                                try:
                                    random.choice(tick).click()
                                except:
                                    pass
                            time.sleep(1)

                    selects = div.find_elements_by_xpath('.//select[@class="field_class"]')
                    if selects:
                        for select in selects:
                            nums = len(Select(select).options)
                            Select(select).select_by_index(random.choice(range(1, nums)))
                            time.sleep(1)
                    time.sleep(1)
                except (StaleElementReferenceException, ElementNotVisibleException):
                    pass

            time.sleep(2)

        print 'done!'
        driver.quit()




def main():
    datas = ExcelReader().data
    for data in datas:
        task = Task(data)
        # task.change_proxy()
        task.run()


if __name__ == '__main__':
    main()
