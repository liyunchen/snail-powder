# -*- coding: utf-8 -*-
### 2021-2-6  李运辰

import requests
import json
import re
import time
import operator
import jieba
from wordcloud import WordCloud
import xlrd
import xlwt
from xlutils.copy import copy
import matplotlib as mpl
from matplotlib import pyplot as plt
import pandas as pd
from stylecloud import gen_stylecloud

# 写入execl
def write_excel_xls_append(path, value):
    index = len(value)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        for j in range(0, len(value[i])):
            new_worksheet.write(i+rows_old, j, value[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入
    new_workbook.save(path)  # 保存工作簿

# 初始化execl表
def initexcel():

    # 创建一个workbook 设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('sheet1')
    workbook.save('螺蛳粉.xls')
    ##写入表头
    value1 = [["标题", "销售地", "销售量", "评论数", "销售价格", '商品惟一ID', '图片URL']]
    book_name_xls = '螺蛳粉.xls'
    write_excel_xls_append(book_name_xls, value1)

# 采集数据
def get_data():
    headers = {
            'Host':'s.taobao.com',
            'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/85.0.4183.121 Safari/537.36',
            'cookie':'cna=QsEFGOdo0BICARsnWHe+63/1; hng=CN%7Czh-CN%7CCNY%7C156; thw=cn; t=effdb32648fc8553a0d1a87926b80343; _m_h5_tk=a94dfbbc27ac02cdbf2cee2a89350b6a_1612614558558; _m_h5_tk_enc=c43b209ec0ed1292bcc622bef5ee6af5; cookie2=1a3fea5ffa0fad17b8c0bbaef21ebb68; _tb_token_=5de15eeea0fbe; xlly_s=1; _samesite_flag_=true; sgcookie=E1007k5qmQ9jBth1shqyTbJtsfmA3xbZNA9skFhamSfqcP7GZBjDZXwyW%2Fnbs39HPqifkG%2FiNiy0TB3VOa4TvxBSyg%3D%3D; unb=913134998; uc3=lg2=U%2BGCWk%2F75gdr5Q%3D%3D&vt3=F8dCuAc6zt7X28yBUrc%3D&id2=WvEIwUQBSki%2F&nk2=rW6iZSg5; csg=4de33d18; lgc=%5Cu897F%5Cu95E8%5Cu5EC9; cookie17=WvEIwUQBSki%2F; dnk=%5Cu897F%5Cu95E8%5Cu5EC9; skt=3fa41897557f2c39; existShop=MTYxMjYwNDU4NA%3D%3D; uc4=nk4=0%40r5%2FGFBQ7A5tJI1TpQam3MZQ%3D&id4=0%40WDb9t1Fxtm4iZCHd0tESONEjEoU%3D; publishItemObj=Ng%3D%3D; tracknick=%5Cu897F%5Cu95E8%5Cu5EC9; _cc_=WqG3DMC9EA%3D%3D; _l_g_=Ug%3D%3D; sg=%E5%BB%898a; _nk_=%5Cu897F%5Cu95E8%5Cu5EC9; cookie1=UUo1TGxcH8cPfpMWT7%2FuMD1anzLFJTzG47%2FnHaFSftY%3D; enc=1xoAdBLlK2BdC0gn79RjfmESRECbfDEgAmzpogjAgEE8dU2FQDF0xFpDq1gxeXD00WiK6XHZ9Wd3C3ltW9vaZw%3D%3D; mt=ci=10_1; uc1=pas=0&cookie15=Vq8l%2BKCLz3%2F65A%3D%3D&cookie21=UtASsssme%2BBq&cookie16=WqG3DMC9UpAPBHGz5QBErFxlCA%3D%3D&existShop=false&cookie14=Uoe1gB38uZ7EFQ%3D%3D; JSESSIONID=7137BBC97E23304D98ADE4E546DB686C; isg=BJ6eJexZctdNAZkZHuCIDdMx7zTgX2LZ0qNVJUgnCuHcaz5FsO-y6cQJZ3fnyFrx; l=eBIj49hqOGMgJqhbBOfanurza77OSIRYYuPzaNbMiOCP9Z5B5f2GW6MUrvY6C3GVh6XXR3yMI8QMBeYBqQAonxv92j-la_kmn; tfstk=c0ifByNUGsffR08N0x9P0RJhfBqOwvI7EgVrhqJE3SL7nW1mfMPBSlefNgULF',
            'accept': 'image/avif,image/webp,image/apng,image/*,*/*;q=0.8',
            'accept-encoding': 'gzip, deflate, br',
            'accept-language': 'zh-CN,zh;q=0.9',
            'upgrade-insecure-requests': '1',
            'referer':'https://s.taobao.com/',
        }
    ###请求url
    #每页44条 规律：s的跨度为44
    # s = 0 44 88 132
    for i in range(0,101):
        print(i)
        url="https://s.taobao.com/search?q=螺蛳粉&ie=utf8&bcoffset=0&ntoffset=0&s="+str(i*44)
        ###requests+请求头headers
        r = requests.get(url, headers=headers)
        r.encoding = 'utf8'
        s = (r.content)
        ###乱码问题
        html = s.decode('utf8')
        #print(html)

        # 正则模式
        p_title = '"raw_title":"(.*?)"'       #标题
        p_location = '"item_loc":"(.*?)"'    #销售地
        p_sale = '"view_sales":"(.*?)人付款"' #销售量
        p_comment = '"comment_count":"(.*?)"'#评论数
        p_price = '"view_price":"(.*?)"'     #销售价格
        p_nid = '"nid":"(.*?)"'              #商品惟一ID
        p_img = '"pic_url":"(.*?)"'          #图片URL

        # 数据集合
        data = []
        # 正则解析
        title = re.findall(p_title,html)
        location = re.findall(p_location,html)
        sale = re.findall(p_sale,html)
        comment = re.findall(p_comment,html)
        price = re.findall(p_price,html)
        nid = re.findall(p_nid,html)
        img = re.findall(p_img,html)
        for j in range(len(title)):
            data.append([title[j],location[j],sale[j],comment[j],price[j],nid[j],img[j]])

        book_name_xls = '螺蛳粉.xls'
        write_excel_xls_append(book_name_xls, data)
        time.sleep(3)
        #print(data)
        #print(len(data))


# matplotlib中文显示
plt.rcParams['font.family'] = ['sans-serif']
plt.rcParams['font.sans-serif'] = ['SimHei']
# 读取数据
# encoding='utf-8',engine='python'
IO = '螺蛳粉.xls'
data = pd.read_excel(io=IO)

###分析1：分析价格分布
def analysis1():
    # 价格分布
    plt.figure(figsize=(16, 9))
    plt.hist(data['销售价格'], bins=20, alpha=0.6)
    plt.title('价格频率分布直方图')
    plt.xlabel('价格')
    plt.ylabel('频数')
    plt.savefig('价格分布.png')

###分析2：销售地分布
def analysis2():
    # 销售地分布
    group_data = list(data.groupby('销售地'))
    loc_num = {}
    for i in range(len(group_data)):
        loc_num[group_data[i][0]] = len(group_data[i][1])
    plt.figure(figsize=(55, 9))
    plt.title('销售地')
    plt.scatter(list(loc_num.keys())[:20], list(loc_num.values())[:20], color='r')
    plt.plot(list(loc_num.keys())[:20], list(loc_num.values())[:20])
    plt.savefig('销售地.png')

    sorted_loc_num = sorted(loc_num.items(), key=operator.itemgetter(1), reverse=True)  # 排序
    loc_num_10 = sorted_loc_num[:10]  # 取前10
    loc_10 = []
    num_10 = []
    for i in range(10):
        loc_10.append(loc_num_10[i][0])
        num_10.append(loc_num_10[i][1])
    plt.figure(figsize=(16, 9))
    plt.title('销售地TOP10')
    plt.bar(loc_10, num_10, facecolor='lightskyblue', edgecolor='white')
    plt.savefig('销售地TOP10.png')

###分析3：词云分析
def analysis3():
    # 制作词云
    content = ''
    for i in range(len(data)):
        content += data['标题'][i]
    wl = jieba.cut(content, cut_all=True)
    wl_space_split = ' '.join(wl)
    pic = '词云图.png'
    gen_stylecloud(text=wl_space_split,
                   font_path='simsun.ttc',
                   # icon_name='fas fa-envira',
                   icon_name='fab fa-qq',
                   max_words=100,
                   max_font_size=70,
                   output_name=pic,
                   )  # 必须加中文字体，否则格式错误

###分析4：线性回归分析
def analysis4():
    datas = data
    datas = datas.dropna(axis=0, how='any')
    x = datas['销售量']
    y = datas['销售价格']
    x = x.tolist()
    y = y.tolist()
    for i in range(0, len(x)):
        j = x[i]
        if "+" in j:
            j = j.replace("+", "")
        if "万" in j:
            j = j.replace("万", "")
            j = float(j) * 10000
        x[i] = str(j)
    flg, ax = plt.subplots()
    ax.scatter(x,y, alpha=0.5,edgecolors= 'white')
    ax.set_xlabel('销量')
    ax.set_ylabel('价格')
    ax.set_title('商品价格对销量的影响')
    #隐藏刻度线和标签
    ax.set_xticks([])
    #plt.show()
    plt.savefig('商品价格对销量的影响.png')


    """
    # 调用线性回归函数
    clf = linear_model.LinearRegression()
    # 开始线性回归计算
    clf.fit(x, y)
    # 得到斜率
    print(clf.coef_[0])
    # 得到截距
    print(clf.intercept_)
    """

# 初始化execl表
#initexcel()
# 采集数据
#get_data()

###分析1：分析价格分布
#analysis1()
###分析2：销售地分布
#analysis2()
###分析3：词云分析
#analysis3()
###分析4：线性回归分析
analysis4()

