# web crawler -> Excel
爬取国家社科基金项目数据库内容写入已有名称excel
数据库名称Excel_test.xls
可以在包下通过以下代码进行创建到本地文件中
  workbook = xlwt.Workbook(encoding = 'utf-8')
  worksheet = workbook.add_sheet('My Worksheet')
  workbook.save('Excel_test.xls')

但是head部分要自己填写
"","项目批准号","项目类别","学科分类","项目名称","立项时间","项目负责人","专业职务","工作单位","单位类别","所在省区市","所属系统","成果名称","成果形式","成果等级"
引用的包有这些
import sys
import time
import hashlib
import requests
import urllib3
import re
import ssl
import xlwt
import xlrd
from xlutils.copy import copy
可以使用pip install
