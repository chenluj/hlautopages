# -*- coding: utf-8 -*-

from utils import *
from selenium import webdriver
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located, visibility_of_element_located
from selenium.common.exceptions import TimeoutException





def main():
    try:
        tasks = YamlReader().data
        conf = Config(tasks.pop(0))
    except:
        logger.error(u'[Error] 读取配置文件出错')
    else:
        browser = Browser(conf)
        for task in tasks:
            logger.info(u'[Info] 执行任务  {}'.format(str(task)))
            try:
                t = Task(task)
            except:
                logger.error(u'[Error] 初始化任务出错，请检查配置或数据文件，确认填写无误并且变量名与列名对应')
                os._exit(0)
            else:
                try:
                    t.run(browser)
                except:
                    logger.error(u'[Error] 执行任务出错，请检查配置与页面是否对应')
                    browser.quit()
                    os._exit(0)

        logger.info(u'[Info] 所有任务执行结束，请处理数据后重新启动程序\n')


if __name__ == '__main__':
    main()
