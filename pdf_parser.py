#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pymupdf
import os
import sys
import re
import datetime
import json
import hashlib
import pandas as pd
import io
import copy
import numpy
from PIL import Image

qreader = None

def extract_text_from_pdf(pdf_file):
    text = ""
    pdf_document = pymupdf.open(pdf_file)
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text += page.get_text()

    # global qreader
    # if qreader is None:
    #     from qreader import QReader
    #     qreader = QReader(model_size="n")
    #     print("inited qreader")

    # image_list = pdf_document.get_page_images(0)
    # for img_index, img in enumerate(image_list):
    #     # (10, 0, 200, 200, 8, 'DeviceRGB', '', 'Xi46', 'FlateDecode')
    #     # xref (int),smask (int), width and height (ints), bpc (int),colorspace (str), alt. colorspace (str), name (str),filter (str),referencer (int) the xref of the referencer. Zero if directly referenced by the page. Only present if full=True.
    #     xref, smask, width, height, bpc, colorspace, altcolorspace, name, adobefilter = img
    #     if width!=height or width<50:
    #         continue

    #     base_image  = pdf_document.extract_image(xref)
    #     image_bytes = base_image["image"]


    #     print(img)
    #     try:
    #         image = Image.open(io.BytesIO(image_bytes))
    #         image.show()
    #     except:
    #         image_name = "%s-img.png"%(pdf_file)
    #         with open(image_name, "wb") as image_file:
    #             image_file.write(image_bytes)
    #         input("saved")

    #     decoded_text = qreader.detect_and_decode(image=numpy.array(image))
    #     input(decoded_text)

    return text

def rm_dup(findall):
    "去除 list findall 中的重复项"
    nondup = []
    for i in findall:
        if i not in nondup:
            nondup.append(i)
    return nondup

testchinese = {
    '捌拾捌圆叁角壹':88.31,
    '叁仟玖佰叁拾肆圆伍角整':3934.5,
    '壹佰肆拾捌圆壹角':148.1,
    '壹万伍仟伍佰玖拾捌圆整':15598,
    '壹佰柒拾伍圆整':175,
    '肆拾肆圆陆角陆':44.66,
    '肆拾肆圆玖角捌':44.98,
    '柒圆肆角':7.4,
    '壹仟伍佰肆拾壹圆肆角壹':1541.41,
    '贰拾陆圆贰角陆':26.26,
    '叁拾玖圆壹角伍':39.15,
    '壹仟贰佰叁拾玖圆零捌分':1239.08,
    '叁拾圆整':30.00,
    '伍拾陆元整':56.00
    }

muldict = {'拾':10,'佰':100,'仟':1000,'万':10000,'圆':1,'元':1,'角':0.1,'分':0.01}
numdict = {'壹':1,'贰':2,'叁':3,'肆':4,'伍':5,'陆':6,'柒':7,'捌':8,'玖':9}

def chinese2num(chinese):
    "中文数字转数字"
    num = 0
    mul = 0.01
    for i in chinese[::-1]:
        if i in '整零':
            continue
        if i in muldict:
            mul = muldict[i]
        else:
            num+= numdict[i]*mul
    # lastnum = 0
    # for i in chinese:
    #     if i in '整零':
    #         continue
    #     if i in muldict:
    #         num += muldict[i]*lastnum
    #         lastnum = 0
    #     else:
    #         lastnum = numdict[i]
    # else:
    #     num += lastnum*0.01

    return round(num*10000)/10000

# 主要匹配以下情况:
# 最正常的: xx地xx有限公司
# 离谱: xx地xx有限公司xx店
# 还有: xx地xx(北京)有限公司
pcompany  = re.compile("([\u4e00-\u9fa5（）]+有限公司|[\u4e00-\u9fa5（）]+有限公司[\u4e00-\u9fa5]+店)\n")

pdate     = re.compile("(202[0-9])\\s{0,2}年\\s{0,2}([0-9]{2})\\s{0,2}月\\s?([0-9]{2})\\s{0,2}日")
# 有的发票是 "年 月 日" 是一个文字, 然后 "2024  01  01" 再塞到刚刚 "年 月 日" 的地方
pdate2    = re.compile("(202[0-9])\\s{1,2}([0-9]{2})\\s{1,2}([0-9]{2})")

pprice    = re.compile("[¥￥]\\s{0,2}([0-9]+\\.[0-9]{2})")
ppricecap = re.compile("[壹贰叁肆伍陆柒捌玖拾佰仟万圆元角分整零]+")

# 发票类别
pcategory = re.compile("\\*([\u4e00-\u9fa5]+)\\*") # 这串re是中文的意思

# 专票还是普票：电子发票(增值税专用发票) or 北京增值税电子普通发票
pzhuanpiao = re.compile("(专用发票|普通发票)")

# 获取上个月的月份
current_date = datetime.datetime.now()
this_month = current_date.month
last_month_date = current_date - datetime.timedelta(days=current_date.day)
last_month = last_month_date.month

def extract_invoice_info(pdf_file,check_month=True):
    pdf_text = extract_text_from_pdf(pdf_file)
    info = {}

    companies = rm_dup(i.strip() for i in pcompany.findall(pdf_text))
    info['companies'] = companies

    date = rm_dup(pdate.findall(pdf_text))
    if len(date)==0:
        date = rm_dup(pdate2.findall(pdf_text))
    if len(date)>0:
        date = tuple(int(i) for i in date[0])
        info['date'] = date

    pricecap = rm_dup(ppricecap.findall(pdf_text))
    pricecap = [i for i in pricecap if len(i)>2]
    info['pricecap'] = pricecap[0]
    info['pricecapnum'] = chinese2num(pricecap[0])

    prices = rm_dup(pprice.findall(pdf_text))
    prices = list(float(i) for i in prices if float(i)>0)
    prices.sort(reverse=True)
    info['prices'] = prices

    category = rm_dup(pcategory.findall(pdf_text))
    category.sort()
    zengzhi = pzhuanpiao.findall(pdf_text)[0]
    info['category'] = category
    info['zengzhi']  = zengzhi

    # print(json.dumps(info, indent=4, ensure_ascii=False))
    # if len(msgs)>0:
    #     print("errs",errs)
    #     msgs = "\n".join(msgs)
    #     if "fatal" in msgs:
    #         input(msgs)
    #     else:
    #         print(msgs)

    return info

def check_invoice_info(info,check_month=False,check_this_month=False):
    errs = []
    msgs = []
    if len(info["companies"])!=2:
        errs.append('companies')
        msgs.append("发票中有 %d 个公司: %s"%(len(info["companies"]),info["companies"]))
    elif '北京自伴科技有限公司' not in info["companies"]:
        errs.append('companies')
        msgs.append("fatal: 公司中没有自伴：%s"%(info["companies"]))

    if 'date' not in info:
        errs.append("date")
        msgs.append("发票中没有日期")
    else:
        if info["date"][0]!=2024 or not (1<=info["date"][1]<=12) or not (1<=info["date"][2]<=31):
            errs.append("date")
            msgs.append("发票日期错误: %s"%(info["date"],))

        if check_month:
            if not check_this_month and info["date"][1]!=last_month:
                errs.append("date")
                msgs.append("fatal: 发票不是上个月的：%s"%(info["date"],))
            elif check_this_month and info["date"][1]!=this_month:
                errs.append("date")
                msgs.append("fatal: 发票不是这个月的：%s"%(info["date"],))

    if len(info['prices'])!=3:
        errs.append("prices")
        msgs.append("发票中有 %d 个价格: %s"%(len(info['prices']),info['prices']))

    if info["prices"][0]!=info['pricecapnum']:
        errs += ['prices','pricecapnum']
        msgs.append("fatal: 大写价格与发票价格不符: %s(%s) != %s"%(info['pricecap'],info['pricecapnum'],info['prices']))

    return errs,msgs

def correct_invoice_info(info,errs,msgs):
    print("发票信息需要更正, 更正前信息",json.dumps(info, indent=4, ensure_ascii=False))
    if len(msgs)>0:
        print("错误条目:",errs)
        msgs = "\n".join(msgs)
        print("错误信息:",msgs)

    for e in errs:
        if e=="date":
            d = input("请帮我输入日期, 格式: 年,月,日\n")
            info["date"] = tuple(int(i) for i in d.split(','))
        elif e=="prices":
            d = input("请帮我输入日期, 格式: 价税合计,金额,税额\n")
            info["prices"] = list(float(i) for i in d.split(','))
        elif e=="pricecapnum":
            d = input("请帮我输入大写金额对应的数字\n")
            info["pricecapnum"] = float(d)

    print("更正后信息",json.dumps(info, indent=4, ensure_ascii=False))
    return info

def get_std_name(info):
    "根据信息获取标准名称"
    names = []
    # if info['zengzhi']=='专用发票':
    #     names.append('专票')

    if info['companies'][0]=="北京自伴科技有限公司":
        company = info['companies'][1]
    else:
        company = info['companies'][0]

    if '京东' in company:
        abstract = '京东'
        details = "".join(info['category'])
    elif '立创' in company:
        abstract = '嘉立创'
        details = None
    elif '滴滴' in company:
        abstract = '滴滴'
        details = None
    elif '象鲜' in company:
        abstract = '美团买菜'
        details = None
    elif '美团' in company:
        abstract = '美团'
        details = "".join(info['category'])
    elif '餐饮' in company:
        abstract = '餐饮'
        details = None
    else:
        abstract = "".join(info['category'])
        details = company

    names.append(abstract)
    names.append(str(info['prices'][0]))
    if details is not None:
        names.append(details)

    std_name = "-".join(names)
    h = hashlib.md5(std_name.encode('utf8')).hexdigest()
    std_name+= "-%s.pdf"%(h[0:2])
    return std_name

def check_std_name(name):
    names = name.split("/")[-1].split("-")
    realname = "-".join(names[0:-1])
    md5 = names[-1].split(".")[0]
    h = hashlib.md5(realname.encode('utf8')).hexdigest()
    if h[0:2]==md5:
        return True
    else:
        return False

def get_pdf_files(dir="."):
    "获取文件名"
    pdf_files = os.popen("ls %s/*.pdf"%(dir)).read().strip()
    return pdf_files.split("\n")

def deal_folder(folder_name=None):
    if folder_name is None:
        folder_name = "."
        check_month = True
    else:
        check_month = False

    check_this_month = "this-month" in " ".join(sys.argv)

    pdf_files = get_pdf_files(folder_name)

    invoices = []
    for pdf_file in pdf_files:
        print("dealing %s"%(pdf_file))
        info = extract_invoice_info(pdf_file)
        errs,msgs = check_invoice_info(info,check_month=check_month,check_this_month=check_this_month)
        while len(errs)>0:
            info = correct_invoice_info(info,errs,msgs)
            errs,msgs = check_invoice_info(info,check_month=check_month,check_this_month=check_this_month)
            input("继续?")

        std_name = get_std_name(info)
        info['std_name'] = std_name

        std_pdf_file = pdf_file.split("/")
        std_pdf_file[-1] = std_name
        std_pdf_file = "/".join(std_pdf_file)
        if pdf_file!=std_pdf_file:
            os.popen('mv "%s" "%s"'%(pdf_file,std_pdf_file))
            print("renamed as %s"%(std_pdf_file))

        for inv in invoices:
            if info['std_name'] == inv["std_name"]:
                input("与发票 %s 重复, 请手动删除一个"%(inv['std_name']))
                break
        else:
            invoices.append(info)

    # df = pd.DataFrame(invoices)
    # excel_name = "发票-%s月.xlsx"%(last_month)
    # df.to_excel("output.xlsx", index=False)
    # print("数据已成功导出到 %s"%(excel_name))
    tot = sum(i["pricecapnum"] for i in invoices)
    print("总计 %.4f 元"%(tot))

    excel = []
    for i in invoices:
        e = copy.deepcopy(i)
        e['companies'] = ', '.join(e['companies'])
        e['date'] = ''.join(["%02d"%(j) if j<99 else str(j) for j in e['date']])
        e['价税合计'] = e['prices'][0]
        e['金额'] = e['prices'][1]
        e['税额'] = e['prices'][2]
        e['项目名称'] = ', '.join(e['category'])
        for j in ['pricecap','pricecapnum','zengzhi','prices','category']:
            e.pop(j)
        excel.append(e)

    filename = "发票-%.2f.xlsx"%(tot,)
    df = pd.DataFrame(excel)
    writer = pd.ExcelWriter(filename,engine='xlsxwriter')
    df.to_excel(writer, sheet_name='sheet1', index = False)
    for column in df:
        column_len = max(2*df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['sheet1'].set_column(col_idx, col_idx, column_len)
    writer.close()

    print("发票信息已保存到 %s"%(filename))

if __name__=="__main__":
    # # test chinese
    # for i,j in testchinese.items():
    #     assert j==chinese2num(i), "%s -> %s"%(i,chinese2num(i))
    # input()

    # print(extract_text_from_pdf("202404-4659.87/餐饮服务-26.8-北京肯德基有限公司-1d.pdf"))
    # input()

    if len(sys.argv)>=2:
        if "help" in sys.argv[1]:
            print("usage:\n./pdf_parser.py\n./pdf_parser.py folder\n./pdf_parser.py --this-month")
            sys.exit()
        elif not sys.argv[1].startswith("-"):
            deal_folder(sys.argv[1])
        else:
            deal_folder()
    else:
        deal_folder()


