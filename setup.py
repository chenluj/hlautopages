# -*- coding: utf-8 -*-

from distutils.core import setup
import py2exe, sys, os
sys.argv.append('py2exe')

wd_path = 'C:\\APP\\Python2.7.10\\Lib\\site-packages\\selenium\\webdriver'
required_data_files = [('selenium/webdriver/firefox',
                        ['{}\\firefox\\webdriver.xpi'.format(wd_path), '{}\\firefox\\webdriver_prefs.json'.format(wd_path)])]

options = {"py2exe": {
    "compressed": 1,  # 压缩
    "optimize": 2,
    # "skip_archive": True,
    "bundle_files": 1,  # 所有文件打包成一个exe文件
    # 'dll_excludes': ["mswsock.dll", "powrprof.dll"]
}}

setup(
    console=[{'script': "AutoRun.py", "icon_resources": [(1, "hacker.ico")]}],
    # data_files=required_data_files,
    options=options,
    zipfile=None
)

