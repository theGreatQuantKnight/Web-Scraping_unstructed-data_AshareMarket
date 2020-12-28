import requests
from bs4 import BeautifulSoup




class WebScrapInquires(object):

    def __init__(self):
        self.SHheaders = {
            'Referer': 'http://www.sse.com.cn/disclosure/credibility/supervision/inquiries/',
            'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36'
        }
        self.SZheaders={
            'Referer': 'http://www.szse.cn/disclosure/supervision/inquire/index.html',
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.140 Safari/537.36 Edge/17.17134'
        }

    def GetSHJsonurl(self,currentpage):
        SHurl = "http://query.sse.com.cn/commonSoaQuery.do?siteId=28&sqlId=BS_KCB_GGLL&extGGLX=&stockcode=&channelId=10743%2C10744%2C10012&extGGDL=&order=createTime%7Cdesc%2Cstockcode%7Casc&isPagination=true&pageHelp.pageSize=15&pageHelp.pageNo=" + repr(currentpage) + "&pageHelp.beginPage=" + repr(currentpage) + "&pageHelp.cacheSize=1&pageHelp.endPage=21&type=&_=1609121210269"
        return SHurl
    def GetSZJsonurl(self,currentpage):
        SZurl = "http://www.szse.cn/api/report/ShowReport/data?SHOWTYPE=JSON&CATALOGID=main_wxhj&TABKEY=tab1&PAGENO=" + repr(currentpage)
        return SZurl

    def GetSHList(self,pageNum):
        with open('上交所问询函.txt', "a") as f:
            for page in range(1,pageNum+1):
                r = requests.get(self.GetSHJsonurl(page), headers=self.SHheaders)
                for i in r.json()['result']:
                    f.write('\t'.join([i['cmsOpDate'], i['docTitle'], i['stockcode'], i['extWTFL'], i['extGSJC'], i['docType'],i['createTime'], i['docURL']]) + '\n')
                print('上交所问询函已爬取第%d页' % page)

    def GetSZList(self,pageNum):
        with open('深交所问询函.txt', "a") as f:
            for page in range(1,pageNum+1):
                r = requests.get(self.GetSZJsonurl(page), headers=self.SZheaders)
                for i in r.json()[0]['data']:
                    f.write('\t'.join([i['gsdm'], i['gsjc'], i['fhrq'], i['hjlb'],'reportdocs.static.szse.cn'+BeautifulSoup(r.json()[0]['data'][2]['ck'], "html.parser").a['encode-open'],'reportdocs.static.szse.cn'+BeautifulSoup(r.json()[0]['data'][2]['hfck'], "html.parser").a['encode-open']]) + '\n')
                print('深交所问询函已爬取第%d页' % page)





import pandas as pd
import time
from urllib.request import urlopen
from urllib.request import Request
from urllib.request import quote
import requests
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfinterp import PDFTextExtractionNotAllowed
from pdfminer.pdfparser import PDFParser, PDFDocument


data = pd.read_table('/Users/mengjiexu/深交所回复列表.txt',header=None,encoding='utf8',delim_whitespace=True)
data.columns=['函件编码','函件类型']


函件编码 = data.loc[:,'函件编码']
函件类型 = data.loc[:,'函件类型']

headers = {'content-type': 'application/json',
           'Accept-Encoding': 'gzip, deflate',
           'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64; rv:53.0) Gecko/20100101 Firefox/53.0'}

baseurl = "http://reportdocs.static.szse.cn/UpFiles/fxklwxhj/"

def parse(docucode):
    # 打开在线PDF文档
    _path = baseurl + quote(docucode) +"?random=0.3006649122149502"
    request = Request(url=_path, headers=headers)  # 随机从user_agent列表中抽取一个元素
    fp = urlopen(request)
    # 读取本地文件
    # path = './2015.pdf'
    # fp = open(path, 'rb')
    # 用文件对象来创建一个pdf文档分析器
    praser_pdf = PDFParser(fp)
    # 创建一个PDF文档
    doc = PDFDocument()
    # 连接分析器 与文档对象
    praser_pdf.set_document(doc)
    doc.set_parser(praser_pdf)
    # 提供初始化密码doc.initialize("123456")
    # 如果没有密码 就创建一个空的字符串
    doc.initialize()
    # 检测文档是否提供txt转换，不提供就忽略
    if not doc.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        # 创建PDf资源管理器 来管理共享资源
        rsrcmgr = PDFResourceManager()
        # 创建一个PDF参数分析器
        laparams = LAParams()
        # 创建聚合器
        device = PDFPageAggregator(rsrcmgr, laparams=laparams)
        # 创建一个PDF页面解释器对象
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        # 循环遍历列表，每次处理一页的内容
        # doc.get_pages() 获取page列表
        for page in doc.get_pages():
            # 使用页面解释器来读取
            interpreter.process_page(page)
            # 使用聚合器获取内容
            layout = device.get_result()
            # 这里layout是一个LTPage对象 里面存放着 这个page解析出的各种对象 一般包括LTTextBox, LTFigure, LTImage, LTTextBoxHorizontal 等等 想要获取文本就获得对象的text属性，
            for out in layout:
                # 判断是否含有get_text()方法，图片之类的就没有
                # if ``hasattr(out,"get_text"):
                docname = "/Users/mengjiexu/罗党论/年报问询函/深交所回复/"+str(docucode).split('.')[0]+'.txt'
                with open(docname,'a') as f:
                    if isinstance(out, LTTextBoxHorizontal):
                        results = out.get_text()
                        print(results)
                        f.write(results)


for i in range(len(函件编码)):
    函件名称 = (函件编码[i] + '.' + 函件类型[i])
    print(函件名称)
    开始爬取时间 = "这是第%d个公告"%i
    print(开始爬取时间)
    print(time.strftime('%Y.%m.%d.%H:%M:%S',time.localtime(time.time())))
    if 函件类型[i]=="pdf":
        parse(函件名称)
        print(函件名称 + "爬取成功")
    else:
        with open("/Users/mengjiexu/深交所回复/%s"%函件名称,'wb') as f:
            _path = baseurl + quote(函件名称) +"?random=0.3006649122149502"
            request = requests.get(url=_path, headers=headers)  # 随机从user_agent列表中抽取一个元素
            f.write(request.content)
    结束爬取时间 = time.strftime('%Y.%m.%d.%H:%M:%S', time.localtime(time.time()))
    print(结束爬取时间)
    print("第%d个公告爬取完成" % i)


import os
import docx2txt
from openpyxl import Workbook

content_list = []

wb = Workbook()
sheet = wb.active
sheet['A1'].value = '公告编码'
sheet['A2'].value = '公告内容'

def readdocx(filepath):
    content = docx2txt.process(filepath)  #打开传进来的路径
    docucode = filepath.split('/')[-1]
    content_list.append([docucode.split('.')[0],content])
    content_list.append([docucode.split('.')[0],content])

def readtxt(filepath):
    content = open(filepath, "r").read()     #打开传进来的路径
    docucode = filepath.split('/')[-1]
    content_list.append([docucode.split('.')[0],content])

def eachFile(filepath):
    pathDir = os.listdir(filepath) #获取当前路径下的文件名，返回List
    for s in pathDir:
        newDir=os.path.join(filepath,s)#将文件命加入到当前文件路径后面
        if os.path.isfile(newDir) :         #如果是文件
            doctype = os.path.splitext(newDir)[1]
            if doctype == ".txt":  #判断是否是txt
                readtxt(newDir)
            elif doctype == ".docx":
                readdocx(newDir)
            else:
                pass
        else:
            eachFile(newDir) #如果不是文件，递归这个文件夹的路径


eachFile("/Users/上交所txt/")
a = 1
for doc in content_list:
    sheet['A%d'%a].value = doc[0]
    print(doc[0])
    sheet['B%d'%a].value = doc[1]
    a += 1
wb.save('上交所问询函.xlsx')

