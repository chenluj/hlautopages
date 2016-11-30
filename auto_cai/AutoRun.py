# -*- coding: utf-8 -*-

import os
import yaml
import time
import copy
import json
import tempfile
import shutil
import random
from xlrd import open_workbook
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located, visibility_of_element_located
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary


CONFIG = 'config.yaml'
DATA = 'data.xlsx'

ACTIONS = ['click', 'clear', 'sendkeys', 'submit', 'select']


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
            return open_workbook(self.book_name)
        except IOError as e:
            print u'[Error] 打开excel出错'
            raise e

    def _sheet(self):
        """Return sheet"""
        try:
            return self.book.sheet_by_name(self.sheet_locator)  # by name
        except Exception as e:
            print u'[Error] sheet {} 不存在'.format(self.sheet_locator)
            raise e

    @property
    def title(self):
        """First row is title."""
        try:
            return self.sheet.row_values(0)
        except IndexError as e:
            print u'[Error] sheet中没有数据'
            raise e

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


WEBDRIVER_PREFERENCES = """
{
  "frozen": {
    "app.update.auto": false,
    "app.update.enabled": false,
    "browser.displayedE10SNotice": 4,
    "browser.download.manager.showWhenStarting": false,
    "browser.EULA.override": true,
    "browser.EULA.3.accepted": true,
    "browser.link.open_external": 2,
    "browser.link.open_newwindow": 2,
    "browser.offline": false,
    "browser.reader.detectedFirstArticle": true,
    "browser.safebrowsing.enabled": false,
    "browser.safebrowsing.malware.enabled": false,
    "browser.search.update": false,
    "browser.selfsupport.url" : "",
    "browser.sessionstore.resume_from_crash": false,
    "browser.shell.checkDefaultBrowser": false,
    "browser.tabs.warnOnClose": false,
    "browser.tabs.warnOnOpen": false,
    "datareporting.healthreport.service.enabled": false,
    "datareporting.healthreport.uploadEnabled": false,
    "datareporting.healthreport.service.firstRun": false,
    "datareporting.healthreport.logging.consoleEnabled": false,
    "datareporting.policy.dataSubmissionEnabled": false,
    "datareporting.policy.dataSubmissionPolicyAccepted": false,
    "devtools.errorconsole.enabled": true,
    "dom.disable_open_during_load": false,
    "extensions.autoDisableScopes": 10,
    "extensions.blocklist.enabled": false,
    "extensions.checkCompatibility.nightly": false,
    "extensions.logging.enabled": true,
    "extensions.update.enabled": false,
    "extensions.update.notifyUser": false,
    "javascript.enabled": true,
    "network.manage-offline-status": false,
    "network.http.phishy-userpass-length": 255,
    "offline-apps.allow_by_default": true,
    "prompts.tab_modal.enabled": false,
    "security.csp.enable": false,
    "security.fileuri.origin_policy": 3,
    "security.fileuri.strict_origin_policy": false,
    "security.warn_entering_secure": false,
    "security.warn_entering_secure.show_once": false,
    "security.warn_entering_weak": false,
    "security.warn_entering_weak.show_once": false,
    "security.warn_leaving_secure": false,
    "security.warn_leaving_secure.show_once": false,
    "security.warn_submit_insecure": false,
    "security.warn_viewing_mixed": false,
    "security.warn_viewing_mixed.show_once": false,
    "signon.rememberSignons": false,
    "toolkit.networkmanager.disable": true,
    "toolkit.telemetry.prompted": 2,
    "toolkit.telemetry.enabled": false,
    "toolkit.telemetry.rejected": true,
    "xpinstall.signatures.required": false,
    "xpinstall.whitelist.required": false
  },
  "mutable": {
    "browser.dom.window.dump.enabled": true,
    "browser.laterrun.enabled": false,
    "browser.newtab.url": "about:blank",
    "browser.newtabpage.enabled": false,
    "browser.startup.page": 0,
    "browser.startup.homepage": "about:blank",
    "browser.usedOnWindows10.introURL": "about:blank",
    "dom.max_chrome_script_run_time": 30,
    "dom.max_script_run_time": 30,
    "dom.report_all_js_exceptions": true,
    "javascript.options.showInConsole": true,
    "network.http.max-connections-per-server": 10,
    "startup.homepage_welcome_url": "about:blank",
    "startup.homepage_welcome_url.additional": "about:blank",
    "webdriver_accept_untrusted_certs": true,
    "webdriver_assume_untrusted_issuer": true
  }
}
"""


class FirefoxProfile(webdriver.FirefoxProfile):
    def __init__(self, profile_directory=None):
        """
        Initialises a new instance of a Firefox Profile

        :args:
         - profile_directory: Directory of profile that you want to use.
           This defaults to None and will create a new
           directory when object is created.
        """
        if not FirefoxProfile.DEFAULT_PREFERENCES:
            FirefoxProfile.DEFAULT_PREFERENCES = json.loads(WEBDRIVER_PREFERENCES)

        self.default_preferences = copy.deepcopy(
            FirefoxProfile.DEFAULT_PREFERENCES['mutable'])
        self.native_events_enabled = True
        self.profile_dir = profile_directory
        self.tempfolder = None
        if self.profile_dir is None:
            self.profile_dir = self._create_tempfolder()
        else:
            self.tempfolder = tempfile.mkdtemp()
            newprof = os.path.join(self.tempfolder, "webdriver-py-profilecopy")
            shutil.copytree(self.profile_dir, newprof,
                            ignore=shutil.ignore_patterns("parent.lock", "lock", ".parentlock"))
            self.profile_dir = newprof
            self._read_existing_userjs(os.path.join(self.profile_dir, "user.js"))
        self.extensionsDir = os.path.join(self.profile_dir, "extensions")
        self.userPrefs = os.path.join(self.profile_dir, "user.js")

    def update_preferences(self):
        for key, value in self.DEFAULT_PREFERENCES['frozen'].items():
            self.default_preferences[key] = value
        self._write_user_prefs(self.default_preferences)

    def add_extension(self, extension=os.path.abspath('webdriver.xpi')):
        self._install_extension(extension)


class Config:
    def __init__(self, conf):
        self.browser = conf['browser'].lower() if 'browser' in conf else 'firefox'
        self.location = conf['location'] if 'location' in conf else None
        self.delay_submit = conf['delay_submit'] if 'delay_submit' in conf else 5
        self.wait_before_if = conf['if_wait'] if 'if_wait' in conf else 3
        self.random_agent = conf['random_agent_spoofer'] if 'random_agent_spoofer' in conf else 'random-agent-spoofer.xpi'
        self.loop = conf['loop'] if 'loop' in conf else False


class Browser:

    def __init__(self, conf):
        self.driver = None
        self.conf = conf
        self.browser = conf.browser
        self.location = conf.location
        # self.delay_submit = conf.delay_submit
        # self.wait_before_if = conf.wait_before_if
        self.random_agent = conf.random_agent
        # self.loop = conf.loop

    def open(self):
        if self.browser == 'firefox':
            try:
                binary = FirefoxBinary(self.location)
                profile = FirefoxProfile()
                if self.random_agent:
                    profile.add_extension(os.path.abspath(self.random_agent))
                self.driver = webdriver.Firefox(firefox_binary=binary, firefox_profile=profile)
                print u'[Info] 打开浏览器  firefox'
                return self
            except:
                print u'[Error] 打开firefox 浏览器失败'
                raise
        elif self.browser == 'chrome':
            try:
                option = webdriver.ChromeOptions()
                option.binary_location = self.location

                self.driver = webdriver.Chrome(executable_path='chromedriver.exe', chrome_options=option)
                print u'[Info] 打开浏览器  chrome'
                return self
            except:
                print u'[Error] 打开chrome浏览器失败'
                raise
        else:
            print u'[Error] 不支持的浏览器类型'
            os._exit(0)

    def get(self, url):
        try:
            self.driver.get(url)
            print u'[Info] 打开URL  {}',
            print url
            return self.driver
        except:
            print u'[Error] 打开URL失败，请检查配置'
            raise

    def quit(self):
        self.driver.quit()
        print u'[Info] 关闭浏览器'


class Element:
    def __init__(self, driver, elem_info, params):
        self.driver = driver
        try:
            self.locator = (elem_info[0], elem_info[1])
            self.element = WebDriverWait(self.driver, 15, 0.5).until(presence_of_element_located(self.locator))
            self.action = elem_info[2].lower()
            self.element_name = elem_info[3]
            self.params = params
            print u'[Info] 元素已找到  {}',
            print elem_info
        except TimeoutException:
            print u'[Error] 未找到元素  {}',
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
                if self.element_name == 'random':
                    nums = len(Select(self.element).options)
                    print u'[Info] 随机从网页选择'
                    Select(self.element).select_by_index(random.choice(range(1, nums)))
                elif isinstance(self.element_name, list):
                    print u'[Info] 从指定选项中随机选择： {}'.format(str(self.element_name))
                    Select(self.element).select_by_value(random.choice(self.element_name))
                else:
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
        self.log = os.path.abspath(os.curdir) + '\\' + self.sheet + '.log'

        if os.path.exists(self.log):
            with open(self.log, 'rb') as f:
                self.num = len(f.read())
        else:
            self.num = 0
        print u'[Info] 检测到已执行 {} 次该任务'.format(self.num)

        xls = ExcelReader(sheet=self.sheet)
        self.loop_times = xls.nums
        self.data = xls.data
        self.task = task

    def run(self, b):
        for t in range(self.num, self.loop_times):
            params = self.data[t]
            print u'[Info] Sheet: "{0}"  Line: "{1}" 开始执行'.format(self.sheet, self.num + 2)
            print u'[Info] ======  任务开始  ======='
            driver = b.open().get(self.url)
            used = 0
            for page in self.task:
                # 如果是Error Page，刷新一次，若仍失败，退出
                error_page = 0
                for i in range(1, 3):
                    try:
                        WebDriverWait(driver, 3, 0.5).until(visibility_of_element_located(('id', 'errorPageContainer')))
                        error_page = i
                        print u'[Error] 得到Error Page'
                        if error_page < 2:
                            print u'[Info] 刷新页面'
                            driver.refresh()
                            time.sleep(10)
                    except TimeoutException:
                        break
                if error_page == 2:
                    print u'[Error] 两次得到Error Page，任务失败'
                    break

                for element in page['elements']:
                    if isinstance(element, dict):
                        if 'if' in element:
                            time.sleep(b.conf.wait_before_if)
                            if element['if'] in driver.current_url:
                                print u'[Info] URL为期待值，任务成功'
                                break
                        elif 'wait' in element:
                            print u'[Info] wait {}s'.format(element['wait'])
                            time.sleep(element['wait'])
                    else:
                        try:
                            Element(driver, element, params).do_its_work(b.conf.delay_submit)
                        except:
                            print u'[Warning] 元素执行失败，跳过该元素'
                        time.sleep(1)
                # 程序执行完第一个elements，则标记为已执行，写入日志
                if used == 0:
                    with open(self.log, 'a') as f:
                        f.write('1')
                        used = 1
                time.sleep(5)
            b.quit()
            print u'[Info] ======  任务结束  ======='
            print u'[Info] Sheet: "{0}"  Line: "{1}" 执行结束\n'.format(self.sheet, self.num + 2)
            print
            if not b.conf.loop:
                return


def main():
    try:
        tasks = YamlReader().data
        conf = Config(tasks.pop(0))
    except:
        print u'[Error] 读取配置文件出错'
    else:
        browser = Browser(conf)
        for task in tasks:
            try:
                print u'[Info] 执行任务  {}'.format(str(task))
                try:
                    t = Task(task)
                except:
                    print u'[Error] 初始化任务出错，请检查配置或数据文件，确认填写无误并且变量名与列名对应'
                    raise
                else:
                    try:
                        t.run(browser)
                    except:
                        print u'[Error] 执行任务出错，请检查配置与页面是否对应'
                        raise
            except:
                try:
                    browser.quit()
                except:
                    pass
                return
        print u'[Info] 所有任务执行结束，请处理数据后重新启动程序\n'


if __name__ == '__main__':
    main()
