#!/usr/bin/env python
# -*- coding: UTF-8 -*-

import requests
import time
import sys
import json
import os
import xlsxwriter
from sys import argv
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from lxml import etree
import re
import threading


# 商品url汇总表
url_list = []
#anti-antiSpider
headers={'accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
'accept-encoding':'gzip, deflate, sdch, br',
'accept-language':'zh-CN,zh;q=0.8',
'cache-control':'max-age=0',
'cookie':'miid=8544429896003171898; l=Ag8PVJ/qEnokMJdW37RCD9jpH6kIR2MG; hng=CN%7Czh-CN%7CCNY%7C156; thw=cn; UM_distinctid=15f43fd46d81e9-03c4d39fdefd28-5e4f2b18-144000-15f43fd46d997; cna=du/JDqJS/D0CAX1H5Y0vwRKy; tracknick=%5Cu9AD8%5Cu7AEF%5Cu9ED1%5Cu5F8B%5Cu5E08%5Cu4E8B%5Cu52A1%5Cu6240; t=b40fa8e81f46f1abebb920cdd331ecce; _cc_=U%2BGCWk%2F7og%3D%3D; tg=0; x=e%3D1%26p%3D*%26s%3D0%26c%3D0%26f%3D0%26g%3D0%26t%3D0%26__ll%3D-1%26_ato%3D0; cookie2=1c3b943a96f1a1a6db9e886c576c83d7; v=0; mt=ci%3D-1_1; isg=AoCAf9vQzgGIHb_OBY3zUIpfUQ4jP7UuRkF2L_oR9xssdSKfohjDYZYH-epP; _tb_token_=e71e1b53b7905; JSESSIONID=BA225793264BED6A7AF9DE0A84272917',
'upgrade-insecure-requests':'1',
'user-agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36'}
# def main(work,key,num,mkname):

def get_date():
	a = time.strftime("%Y/%m/%d", time.localtime()).split('/')
	[year,month,date] = a
	return year,month,date

def main(work, key, num, mkname, start_price, end_price, page_num,thread=0):
    # print type(num),num
    if num == '1':
        url = '''https://s.m.taobao.com/search?event_submit_do_new_search_auction=1\
        &_input_charset=utf-8&topSearch=1&atype=b&searchfrom=1&action=home%3Aredirect_app_action&\
        from=1&q=''' + key + '''&sst=1&n=20&buying=buyitnow&m=api4h5&abtest=10&wlsort=10&page=''' + str(page_num)
    elif num == '3':
        url = '''https://s.m.taobao.com/search?event_submit_do_new_search_auction=1\
        &_input_charset=utf-8&topSearch=1&atype=b&searchfrom=1&action=home%3Aredirect_app_action&\
        from=1&q=''' + key + '''&sst=1&n=20&buying=buyitnow&m=api4h5&abtest=3\
        &wlsort=3&style=list&closeModues=nav%2Cselecthot%2Conesearch&\
        start_price=''' + str(start_price) + '''&end_price=''' + str(end_price) + '''&page=''' + str(page_num)
    else:
        url = '''https://s.m.taobao.com/search?event_submit_do_new_search_auction=1\
        &_input_charset=utf-8&topSearch=1&atype=b&searchfrom=1&action=home%3Aredirect_app_action&\
        from=1&q=''' + key + '''&sst34=1&n=20&buying=buyitnow&m=api4h5&abtest=14&\
        wlsort=14&style=list&closeModues=nav%2Cselecthot%2Conesearch&sort=_sale&page=''' + str(page_num)
    try:

        body = requests.get(url,headers=headers)
        body = body.text.encode('utf8')
        dic_body = eval(body)
        #print(dic_body)
    except Exception as e:
        print ("请求出错，请将下列url放于浏览器中看是否可以打开")
        print (url)
        print (e)
    for i in range(20):
        print ("当前正在采集第 ", i + 1, " 个,第", str(page_num), ' 页','第',str(thread),'号线程;','关键词为',str(key))
        try:
            num_id = dic_body["listItem"][i]['item_id']
        except:
            num_id = ''
        try:
            act = dic_body["listItem"][i]['act']  # 付款数
        except:
            act = ''
        try:
            area = dic_body["listItem"][i]['area']  # 地区
        except:
            area = ''
        try:
                shopid = dic_body["listItem"][i]['userId']  # 地区
        except:
                shopid = ''
        try:
            if dic_body["listItem"][i]['url'].find('https:') != -1:
                auctionURL = dic_body["listItem"][i]['url']  # 商品url
            else:
                auctionURL = "https:" + dic_body["listItem"][i]['url']  # 商品url

                # https://detail.tmall.com/item.htm?id="+str(num_id)
            # print len(auctionURL)
            if(len(auctionURL) > 250):
                auctionURL_1 = auctionURL[:250]
                auctionURL_2 = auctionURL[250::]
            else:
                auctionURL_1 = auctionURL
                auctionURL_2 = ''
        except:
            auctionURL = ''
            auctionURL_1 = ''
            auctionURL_2 = ''
        try:
            name = dic_body["listItem"][i]['name']  # 商品名
        except:
            name = ''
        try:
            nick = dic_body["listItem"][i]['nick']  # 店铺名
        except:
            nick = ''
        try:
            originalPrice = dic_body["listItem"][i]['originalPrice']  # 原始价格
        except:
            originalPrice = ''
        try:
            price = dic_body["listItem"][i]['price']  # 当前价格
        except:
            price = ''
        try:
            pic_path = dic_body["listItem"][i]['pic_path']  # 当前价格
            # print pic_path
            pic_path = pic_path.replace('60x60', '210x210')
            pic_name = str(i + 1 + (page_num - 1) * 20) + '-' + name
            #img_download(pic_name, pic_path, mkname + '/pic')
        except Exception as  e:
            print( e)
            pic_path = ''
        try:
            zkType = dic_body["listItem"][i]['zkType']  # 当前价格
        except:
            zkType = ''
        #调用download_date函数
        '''try:
            if len(auctionURL_2) > 10:
                first = 0
                html_date = download_date(
                    auctionURL_1 + auctionURL_2, work, i + 2, first)
            else:
                first = 0
                html_date = download_date(auctionURL_1, work, i + 2, first)
        except:
            html_date = ''
        print (html_date)'''
        year,month,date = get_date()
        html_date = year+'/'+month+'/'+date
        print((page_num-1)*20+i)

        # 获取店铺评分

        shopcard = requests.get('https://s.taobao.com/api?sid=' + str(shopid) + '&callback=shopcard&app=api&m=get_shop_card')
        shopcard = shopcard.text.lstrip('\n\nshopcard(').rstrip(');')
        attempts=0
        success=False
        while attempts < 3 and not success:
            try:
                #print('shopid=',shopid)
                shopcard=eval(shopcard)
                f_Rate=shopcard['favorableRate']
                D_Rate=shopcard['matchDescription']
                A_Rate=shopcard['serviceAttitude']
                S_Rate=shopcard['deliverySpeed']
                D_Compared=shopcard['descriptionCompared']['text']+shopcard['descriptionCompared']['rate']
                A_Compared=shopcard['attitudeCompared']['text']+shopcard['attitudeCompared']['rate']
                S_Compared=shopcard['deliveryCompared']['text']+shopcard['deliveryCompared']['rate']
                success = True
            except:
                shopcard = requests.get(
                    'https://s.taobao.com/api?sid=' + str(shopid) + '&callback=shopcard&app=api&m=get_shop_card')
                f_Rate = []
                D_Rate = []
                A_Rate = []
                S_Rate = []
                D_Compared = []
                A_Compared = []
                S_Compared = []
                print('店铺信息爬取失败！')
                attempts += 1
                if attempts == 3:
                    break
        date = [name, nick, act, price, originalPrice, zkType, area,
                auctionURL_1, auctionURL_2, pic_path, html_date, num_id,
                shopid,f_Rate,D_Rate,A_Rate,S_Rate,D_Compared,A_Compared,S_Compared]
        # print len(date)
        num = i + 2 + (int(page_num) - 1) * 20
        install_table(date, work, num)




    # 商品名 店铺  付款人数 当前价格 原始价格 优惠类型 地区 商品url  图片url  详情数据# 店铺id
    # name nick act price originalPrice zkType area auctionURL pic_path
    # html_date



def download_date(url, work, i, first):
    # if first == 1:
    '导入商品url，进行详情页面解析'
    if(url.find("taobao") != -1 and first != 1):
        print ("检测为淘宝的页面")
        try:
            driver = webdriver.Chrome()
            print( "正在获取详情页面,url为")
            #url ="https://item.taobao.com/item.htm?id=538287375253&abtest=10&rn=07abc745561bdfad6f726eb186dd990e&sid=46f938ba6d759f6e420440bf98b6caea"
            url = url
            print (url)
            driver.get(url)
            driver.implicitly_wait(40)  # 设置智能超时时间
            html = driver.page_source
            driver.quit()
        except Exception as e:
            print ("页面加载失败", e)
            return 0
        try:
            print ('正在解析页面')
            try:
                selector = etree.HTML(
                    html, parser=etree.HTMLParser(encoding='utf-8'))
            except Exception as  e:
                print( "页面加载失败", e)
                return 0

            try:  # 此部分用于采集每月销量的数据
                print ('正在解析页面1')
                # context=selector.xpath('//div[@class="tm-indcon"]')
                context = selector.xpath('//strong[@id="J_SellCounter"]')
                xiaoliang_date = u''
                for i in range(len(context)):
                    print( '正在解析页面2')
                    temp_date = etree.tostring(
                        context[i])  # .encode('utf-8')
                    print ('***', temp_date)
                    # 去除一切html标签 attributes-list
                    re_h = re.compile('</?\w+[^>]*>')
                    s = re_h.sub('', temp_date) + ','
                    xiaoliang_date += s
                print( '正在解析页面3')
                list_date = xiaoliang_date + ';'
            except Exception as e:
                print (e)
                list_date = u''
            context = selector.xpath('//ul[@class="attributes-list"]/li')
            for i in range(len(context)):  # attributes-list
                # .encode('utf-8')
                a = etree.tostring(context[i])
                b = a.split('>')
                end = b[1].split('<')[0] + ';'
                list_date += end
            print ('&&&&&&&&&&&', list_date.encode('utf8'))
            if len(list_date) < 50:
                print ("数据过少，尝试检测为天猫页面解析")
                try:
                    driver = webdriver.PhantomJS()
                    print( "正在获取详情页面,url为")
                    #url ="https://item.taobao.com/item.htm?id=538287375253&abtest=10&rn=07abc745561bdfad6f726eb186dd990e&sid=46f938ba6d759f6e420440bf98b6caea"
                    #num_id = re.findall('id=[0-9]+&',url)[0].replace('id=','').replace('&','')
                    #url = "https://detail.tmall.com/item.htm?id="+str(num_id)
                    print( url)
                    driver.get(url)
                    driver.implicitly_wait(40)  # 设置智能超时时间
                    html = driver.page_source
                    driver.quit()
                except Exception as e:
                    print ("页面加载失败", e)
                    return 0
                try:
                    print ('正在解析页面')
                    try:
                        selector = etree.HTML(
                            html, parser=etree.HTMLParser(encoding='utf-8'))
                    except Exception as e:
                        print ("页面加载失败", e)
                        return 0
                    try:
                        # 此部分用于采集每月销量的数据
                        context = selector.xpath('//div[@class="tm-indcon"]')
                        xiaoliang_date = u''
                        for i in range(len(context)):
                            temp_date = etree.tostring(
                                context[i], encoding="utf-8")  # .encode('utf-8')
                            re_h = re.compile('</?\w+[^>]*>')  # 去除一切html标签
                            s = re_h.sub('', temp_date) + ','
                            xiaoliang_date += s
                        list_date += xiaoliang_date + ';'
                    except Exception as e:
                        print (e)
                        list_date += u''

                    context = selector.xpath('//ul[@id="J_AttrUL"]/li')
                    print (list_date, len(context))
                    for i in range(len(context)):
                        # .encode('utf-8')
                        a = etree.tostring(context[i], encoding="utf-8")
                        b = a.split('>')
                        end = b[1].split('<')[0] + ';'
                        list_date += end
                    # print list_date.encode('utf8')
                    return list_date
                except Exception as e:
                    print ('页面解析失败')
                    return 0

            return list_date
        except:
            print ('页面解析失败')
            return 0

    elif(url.find("tmall") != -1 and first != 1):
        print ("检测为天猫页面，")
        try:
            driver = webdriver.Chrome()
            print ("正在获取详情页面,url为")
            #url ="https://item.taobao.com/item.htm?id=538287375253&abtest=10&rn=07abc745561bdfad6f726eb186dd990e&sid=46f938ba6d759f6e420440bf98b6caea"
            num_id = re.findall(
                'id=[0-9]+&', url)[0].replace('id=', '').replace('&', '')
            url = "https://detail.tmall.com/item.htm?id=" + str(num_id)
            print (url)
            driver.get(url)
            driver.implicitly_wait(40)  # 设置智能超时时间
            html = driver.page_source
            driver.quit()
        except Exception as e:
            print ("页面加载失败", e)
            return 0
        try:
            print ('正在解析页面')
            try:
                selector = etree.HTML(
                    html, parser=etree.HTMLParser(encoding='utf-8'))
            except Exception as e:
                print ("页面加载失败", e)
                return 0
            try:
                # 此部分用于采集每月销量的数据
                context = selector.xpath('//div[@class="tm-indcon"]')
                xiaoliang_date = u''
                for i in range(len(context)):
                    temp_date = etree.tostring(
                        context[i], encoding="utf-8")  # .encode('utf-8')
                    re_h = re.compile('</?\w+[^>]*>')  # 去除一切html标签
                    s = re_h.sub('', temp_date) + ','
                    xiaoliang_date += s
                list_date = xiaoliang_date + ';'
            except Exception as e:
                print (e)
                list_date = u''

            context = selector.xpath('//ul[@id="J_AttrUL"]/li')
            print (list_date, len(context))
            for i in range(len(context)):
                # .encode('utf-8')
                a = etree.tostring(context[i], encoding="utf-8")
                b = a.split('>')
                end = b[1].split('<')[0] + ';'
                list_date += end
            # print list_date.encode('utf8')
            return list_date
        except Exception as e:
            print ('页面解析失败')
            return 0

    # if (first):
    #     print "检测为天猫页面，"
    #     try:
    #         driver = webdriver.PhantomJS(r"C:\Users\Administrator\Desktop\phantomjs-2.1.1-windows\bin\phantomjs.exe")
    #         print "正在获取详情页面,url为"
    #         #url ="https://item.taobao.com/item.htm?id=538287375253&abtest=10&rn=07abc745561bdfad6f726eb186dd990e&sid=46f938ba6d759f6e420440bf98b6caea"
    #         #num_id = re.findall('id=[0-9]+&',url)[0].replace('id=','').replace('&','')
    #         #url = "https://detail.tmall.com/item.htm?id="+str(num_id)
    #         print url
    #         driver.implicitly_wait(40) #设置智能超时时间
    #         driver.get(url)
    #         html = driver.page_source.encode('utf-8')
    #         driver.quit()
    #     except Exception,e:
    #         print "页面加载失败",e
    #         return 0
    #     try:
    #         print '正在解析页面'
    #         selector=etree.HTML(html, parser=etree.HTMLParser(encoding='utf-8'))
    #         try:
    #         #此部分用于采集每月销量的数据
    #             context=selector.xpath('//div[@class="tm-indcon"]')
    #             xiaoliang_date = u''
    #             for i in range(len(context)):
    #                 temp_date = etree.tostring(context[i], encoding="utf-8")#.encode('utf-8')
    #                 re_h=re.compile('</?\w+[^>]*>')#去除一切html标签
    #                 s=re_h.sub('',temp_date)+','
    #                 xiaoliang_date += s
    #             list_date = xiaoliang_date+';'
    #         except Exception,e:
    #             print e
    #             list_date = u''
    #
    #         context=selector.xpath('//ul[@id="J_AttrUL"]/li')
    #         print list_date,len(context)
    #         for i in range(len(context)):
    #             a = etree.tostring(context[i], encoding="utf-8")#.encode('utf-8')
    #             b = a.split('>')
    #             end  = b[1].split('<')[0]+';'
    #             list_date += end
    #         #print list_date.encode('utf8')
    #         return list_date
    #     except Exception,e:
    #         print '页面解析失败'
    #         return 0



def install_table(date, work, i):
    '''导入数据列表存入表格中 '''
    str_list = ['B', 'C', 'D', 'E', 'F', 'G',
                'H', 'I', 'J', 'K', 'L', 'M', "N",'O','P','Q','R','S','T','U']
    #global worksheet1
    try:
        work.write('A' + str(i), int(i) - 1)
    except Exception as e:
        print ('无法写入')
        print( e)
    for now_str, now_date in zip(str_list, date):
        num = now_str + str(i)
        try:
            work.write(num, now_date)
        except Exception as e:
            print ("无法写入")
            print (e)


def img_download(id, url, mkname):
    '''导入图片url，文件夹名，以id为图片名'''
    try:
        print ("主图下载中")
        #img = requests.get(url).context()
        name = id
        r = requests.get(url, timeout=100)
        #name = int(time.time())
        f = open('./' + mkname + '/' + str(name) + '.jpg', 'wb')
        f.write(r.content)
        f.close()
    except:
        print ("主图下载失败")


def create_mkdir(name,datelist, prefix = 'Data'):
    '''创建文件夹'''
    try:
        print ("开始创建文件夹 ", name)
        os.mkdir(r'./'+ 'Data' )
    except Exception as e:
        print (e)
    try:
        os.mkdir(r'./'+ 'Data'+ '/'+ date_list[0] + date_list[1] + date_list[2] )
    except Exception as e:
        print (e)
    try:
        os.mkdir(r'./'+ 'Data'+ '/'+ date_list[0] + date_list[1] + date_list[2] + '/' + name )
    except Exception as e:
        print (e)
    try:
        os.mkdir(r'./' + 'Data'+ '/'+ date_list[0] + date_list[1] + date_list[2] +'/'+ name + "/pic")
    except Exception as e:
        print (e)


def create_table(name,date_list,prefix = 'Data'):
    ''' 导入表格名字，在当前目录下创建该表格'''
    check_list=['手机壳','电脑','手机','马克杯']
    subs_list = ['Mobile_case','Computer','Phone','Mug']
    filename  = name
    if filename in check_list:
    	filename = subs_list[check_list.index(filename)]

    try:
        name = './' + prefix + '/' + date_list[0] + date_list[1] + date_list[2] + '/' + name + '/' + filename + '_'+ date_list[0] + date_list[1] + date_list[2] + '.xls'

        workbook = xlsxwriter.Workbook(name)
        worksheet1 = workbook.add_worksheet()
        worksheet1.write('A1', 'ID')
        worksheet1.write('B1', u"商品名")
        worksheet1.write('C1', u'店铺')
        worksheet1.write('D1', u'付款人数')
        worksheet1.write('E1', u'当前价格')
        worksheet1.write('F1', u'原始价格')
        worksheet1.write('G1', u'优惠类型')
        worksheet1.write('H1', u'地区')
        worksheet1.write('I1', u'商品url_1')
        worksheet1.write('J1', u'商品url_2')
        worksheet1.write('K1', u'图片url')
        worksheet1.write('L1', u'date')
        worksheet1.write('M1', u'宝贝id')
        worksheet1.write('N1', u'店铺id')
        worksheet1.write('O1', u'店铺好评率')
        worksheet1.write('P1', u'描述符合')
        worksheet1.write('Q1', u'服务态度')
        worksheet1.write('R1', u'物流速度')
        worksheet1.write('S1', u'描述符合与同行相比')
        worksheet1.write('T1', u'服务态度与同行相比')
        worksheet1.write('U1', u'物流速度与同行相比')
        # workbook.close()
        print ('表格构建完成,name', name)
        return worksheet1, workbook
    except Exception as e:
        print( e)

def multi(i,date_list,page_num,key,thread):
    create_mkdir(i,date_list)
    work, workbook = create_table(i,date_list)
    # time.sleep(100)
    print ('开始采集请等待')
    # main(work,key,num,i)
    for now_page_num in range(1, page_num + 1):
        main(work, i, 1, i, '', '', now_page_num,thread)
        time.sleep(5)
    workbook.close()
    print ('采集完成')



if __name__ == '__main__':
    # print argv
    try:
        key = argv[1]
    except:
        print ('请指定关键词作为第一个参数')
        key = ''
    try:
        name = argv[2]
    except:
        print ("请指定输出文件名问第二个参数")
        name = ''
    try:
        num = argv[3]
        # print num ,star_price , end_price
    except:
        print ("请指定排序方式 1 为综合排序 2 为销量排序, 当前默认为综合排序")
        num = 1
    try:
        page_num = int(argv[4])
    except:
        print ('页码错误，默认值为1')
        page_num = 1
    try:
        star_price = argv[5]
        end_price = argv[6]
    except:
        star_price = ''
        end_price = ''

    #key = u'皮裤男'
    print ('启动采集，关键词为：', key, " 存入： ", name, "排序为 ", num, star_price, end_price)
    if (key == '' or name == '' or num == ''):
        print ('参数不正确')
        print ("请按顺序输入参数 关键词 输出文件名 排序方式（1或者2）页数 （价格区间）")
        print ("例如:python Search.py 皮裤男 皮裤男1 2 1")
        print ("先按照默认爬取模式随便爬爬")
        year,month,date = get_date()
        date_list = [year,month,date]
        name = ['手机壳','电脑','手机','马克杯']
        page_num = 20
        for thread,i in enumerate(name):
            t =threading.Thread(target=multi,args=(i,date_list,page_num,key,thread,))
            t.start()
    




    else:
        year,month,date = get_date()
        date_list = [year,month,date]
        create_mkdir(name,date_list)
        work, workbook = create_table(name,date_list)
        # time.sleep(100)
        print ('开始采集请等待')
        # main(work,key,num,name)
        for now_page_num in range(1, page_num + 1):
            main(work, key, num, name, star_price, end_price, now_page_num)
            time.sleep(5)
        workbook.close()
        print ('采集完成')