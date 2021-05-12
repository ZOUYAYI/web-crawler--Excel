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

ssl._create_default_https_context = ssl._create_unverified_context
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
#---------------------------------------------这是ip代理设置部分---------------------------------------------------------
_version = sys.version_info
is_python3 = (_version[0] == 3)
orderno = "ZF2021180544TKL1y6"
secret = "75bf2d0fb2f24993ab2f13f7cab949e4"
ip = "forward.xdaili.cn"
port = "80"
ip_port = ip + ":" + port
timestamp = str(int(time.time()))
string = "orderno=" + orderno + "," + "secret=" + secret + "," + "timestamp=" + timestamp
if is_python3:
    string = string.encode()
md5_string = hashlib.md5(string).hexdigest()
sign = md5_string.upper()
print(sign)
auth = "sign=" + sign + "&" + "orderno=" + orderno + "&" + "timestamp=" + timestamp
print(auth)
proxy = {"http": "forward.xdaili.cn:80","https": "forward.xdaili.cn:80"}
headers = {"Proxy-Authorization": auth,
               "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/89.0.4389.82 Safari/537.36"}

               # "User-Agent": "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.75 Safari/537.36"}
#---------------------------------------------ip代理设置部分end---------------------------------------------------------


def getres(url):
    try:
        res = requests.get(url, headers=headers, proxies=proxy, verify=False, allow_redirects=False, timeout=30)
        print("获取成功")
        return res
    except Exception as e:
        print('错误：', e)


def test(pagenum):
    # url = "http://fz.people.com.cn/skygb/sk/index.php/index/seach/2?pznum=&xmtype=0&xktype=%E9%A9%AC%E5%88%97%C2%B7%E7%A7%91%E7%A4%BE&xmname=&lxtime=0&xmleader=&zyzw=0&gzdw=&dwtype=0&szdq=0&ssxt=0&cgname=&cgxs=0&cglevel=0&jxdata=0&jxnum=&cbs=&cbdate=0&zz=&hj="
    url = "http://fz.people.com.cn/skygb/sk/index.php/index/seach/"+str(pagenum)+"?pznum=&xmtype=0&xktype=%E9%A9%AC%E5%88%97%C2%B7%E7%A7%91%E7%A4%BE&xmname=&lxtime=0&xmleader=&zyzw=0&gzdw=&dwtype=0&szdq=0&ssxt=0&cgname=&cgxs=0&cglevel=0&jxdata=0&jxnum=&cbs=&cbdate=0&zz=&hj="
    res = getres(url)
    res.encoding='utf-8'
    content = res.text
    # print(content)
    pattern = re.compile('<td.*?>.*?<span.*?>(.*?)</span>.*?</td>', re.S)
    contentlist = re.findall(pattern, content)
    print(contentlist)
    # pagesize = len(contentlist)/20
    words = formatlist(contentlist)
    print(words)
    append_to_excel(words,"Excel_test.xls")
    print()

def formatlist(contentlist):
    pagesize = len(contentlist) // 20
    diclist = []
    for i in range(pagesize):
        valuelist = []
        for j in range(14):
            valuelist.append(contentlist[i*20+j])
        # keylist = {"项目批准号","项目类别","学科分类","项目名称","立项时间","项目负责人","专业职务","工作单位","单位类别","所在省区市","所属系统","成果名称","成果形式","成果等级"}
        # keylist = {'项目批准号','项目类别','学科分类','项目名称','立项时间','项目负责人','专业职务','工作单位','单位类别','所在省区市','所属系统','成果名称','成果形式','成果等级'}
        a = {"":"","项目批准号":valuelist[0],"项目类别":valuelist[1],"学科分类":valuelist[2],"项目名称":valuelist[3],"立项时间":valuelist[4],"项目负责人":valuelist[5],"专业职务":valuelist[6],"工作单位":valuelist[7],"单位类别":valuelist[8],"所在省区市":valuelist[9],"所属系统":valuelist[10],"成果名称":valuelist[11],"成果形式":valuelist[12],"成果等级":valuelist[13]}
        # a = dict(zip(keylist,valuelist))
        diclist.append(a)
    return diclist




def append_to_excel(words, filename):
    '''
    追加数据到excel
    :param words: 【item】 [{},{}]格式
    :param filename: 文件名
    :return:
    '''
    try:
        # 打开excel
        word_book = xlrd.open_workbook(filename)
        # 获取所有的sheet表单。
        sheets = word_book.sheet_names()
        print('sheets:',sheets)
        # 获取第一个表单
        work_sheet = word_book.sheet_by_name(sheets[0])
        print('work_sheet:', work_sheet)
        # 获取已经写入的行数
        old_rows = work_sheet.nrows
        print('old_rows',old_rows)
        # 获取表头信息
        heads = work_sheet.row_values(0)
        print('heads:', heads)
        # 将xlrd对象变成xlwt
        new_work_book = copy(word_book)
        print("copy ok")
        # 添加内容
        new_sheet = new_work_book.get_sheet(0)
        print('new_sheet',new_sheet)
        i = old_rows
        for item in words:
            for j in range(len(heads)):
                new_sheet.write(i, j, item[heads[j]])
            i += 1
        print("ok")
        new_work_book.save(filename)
        print('追加成功！')
    except Exception as e:
        print('追加失败！', e)


if __name__ == '__main__':
    for i in range(2,172):
        test(i)
