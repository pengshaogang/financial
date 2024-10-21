import time
import pandas as pd
df = pd.DataFrame()
import requests
from DrissionPage import ChromiumPage

def get_row_data_H(tr):
    tds = tr.eles('css:td')
    if not tds:
        return None
    link_element = tds[1].ele('css:a')
    link = 'https://s.askci.com/' + link_element.attr('href').replace('http:', 'https:') if link_element and link_element.attr('href').startswith('http:') else link_element.attr('href')
    # if tds[11].text == '汽车制造':
    return {
        '序号': tds[0].text,
        '股票代码': tds[1].text,
        '股票简称': tds[2].text,
        '公司名称': tds[3].text,
        # '注册地址': tds[4].text,
        # '主营业务收入(202312) ': tds[5].text,
        # '净利润(202312) ': tds[6].text,
        # '员工人数': tds[7].text,
        # '上市日期': tds[8].text,
        # '招股书': tds[9].text,
        # '公司财报': tds[10].text,
        '行业分类': tds[11].text,
        # '主营业务': tds[12].text,
        # # '市盈率': tds[13].text,
        # '股票链接': link
        }


# if __name__ == '__main__':
#     page = ChromiumPage()
#     initial_url = 'https://s.askci.com/stock/h/ci0000001208-0'  # 汽车整车
#     page.get(initial_url)
#     data_list = []
#     i = 0
#     while True:
#         trs = page.eles('css:#myTable04 > tbody > tr')
#         print(len(trs))
#         for tr_index in range(len(trs)):
#             trs = page.eles('css:#myTable04 > tbody > tr')
#             tr = trs[tr_index]
#             # print("tr = ", tr.text)
#
#             row_data = get_row_data_H(tr)
#             if row_data['行业分类'] == '汽车制造':
if __name__ == '__main__':
    A_stock = {

        # '300678': '中科信息',
        # '600797': '浙大网新',
        # '300418': '昆仑万维',
        # '002777': '久远银海',
        # '300212': '易华录',
        # '300440': '运达科技',
        # '002990': '盛视科技',
        # '300226': '上海钢联',
        # '002368': '太极股份',
        # '300231': '银信科技',
        # '300523': '辰安科技',
        #
        # '300766': '每日互动',
        #
        # '300872': '天阳科技',
        # '688636': '智明达',
        # '300250': '初灵信息',
        # '300295': '三六五网',



        # '002195': '岩山科技',
        #
        # '600131': '国网信通',
        # '300738': '奥飞数据',
        # '600756': '浪潮软件',
        # '300269': '联建光电',
        # '300044': '赛为智能',
        # '000555': '神州信息',
        # '300170': '汉得信息',
        # '002373': '千方科技',
        # '300687': '赛意信息',
        # '603887': '城地香江',
        # '300419': '浩丰科技',
        # '688158': '优刻得-W',
        # '301396': '宏景科技',
        # '600850': '电科数字',
        # '688568': '中科星图',
        # '300541': '先进数通',
        # '301270': '汉仪股份',
        # '300150': '世纪瑞尔',
        # '603869': 'ST智知',
        # '688509': '正元地信',
        # '300508': '维宏股份',
        # '002591': '恒大高新',
        # '688562': '航天软件',
        # '300785': '值得买',
        # '000409': '云鼎科技',
        # '600358': '国旅联合',
        # '002474': '榕基软件',
        # '300264': '佳创视讯',
        # '600718': '东软集团',
        # '300846': '首都在线',
        # '300448': '浩云科技',
        # '603171': '税友股份',
        # '900901': '云赛Ｂ股',--no
        # '200045': '深纺织Ｂ',--no

        # '002642': '荣联科技',
        # '002115': '三维通信',
        # '603918': '金桥信息',
        # '833030': '立方控股',
        # '688258': '卓易信息',
        # '300271': '华宇软件',
        # '002401': '中远海科',
        # '688004': '博汇科技',
        # '838924': '广脉科技',
        # '300168': '万达信息',
        # '688619': '罗普特',
        # '300078': '思创医惠',
        # '838227': '美登科技',
        # '605398': '新炬网络',
        # '688316': '青云科技',
        # '300288': '朗玛信息',
        '688051': '佳华科技',
        '301381': '赛维时代',
        '300895': '铜牛信息',
        '600070': 'ST富润',
        '688500': '慧辰股份',
        '300248': '新开普',
        '003005': '竞业达',
        '301380': '挖金客',
        '300792': '壹网壹创',
        '300096': 'ST易联众',
        '688228': '开普云',
        '600589': '广东榕泰',
        '300532': '今天国际',
        '300209': 'ST有树',
        '002354': '天娱数科',
        '688229': '博睿数据',
        '300245': '天玑科技',
        '300167': 'ST迪威',

        '688039': '当虹科技',
        '688292': '浩瀚深度',
        '300464': '星徽股份',
        '002315': '焦点科技',
        '688365': '光云科技',
        '600228': '返利科技',
        '300020': 'ST银江',
        '002912': '中新赛克',
        '603881': '数据港',
        '300682': '朗新集团',
        '300300': 'ST峡创',
        '600410': '华胜天成',
        '002331': '皖通科技',
        '300941': '创识科技',
        '688787': '海天瑞声',
        '002291': '遥望科技',
        '002609': '捷顺科技',
        '002264': '新华都',
        '002095': '生意宝',
        '002766': '索菱股份',
        '603636': '南威软件',
        '002316': 'ST亚联',
        '300592': '华凯易佰',
        '301299': '卓创资讯',
        '600271': '航天信息',
        '301428': '世纪恒通',
        '300634': '彩讯股份',
        '003010': '若羽臣',
        '002771': '真视通',
        '002657': '中科金财',
        '688088': '虹软科技',
        '002131': '利欧股份',
        '300277': '海联讯',
        '300324': '旋极信息',

        '300645': '正元智慧',
        '603613': '国联股份',
        '301110': '青木科技',
        '301085': '亚康股份',
        '300399': '天利科技',
        '300518': '新迅达',
        '301001': '凯淳股份',
        '300079': '数码视讯',
        '300017': '网宿科技',
        '300383': '光环新网',
        '002530': '金财互联',
        '000997': '新大陆',
        '301382': '蜂助手',
        '002649': '博彦科技',
        '301171': '易点天下',
        '600571': '信雅达',
        '603496': '恒为科技',
        '300609': '汇纳科技',
        '002232': '启明信息',
        '300166': '东方国信',
        '600446': '金证股份',
        '300442': '润泽科技',
        '002380': '科远智慧',
        '000676': '智度股份',
        '300287': '飞利信',
        '600845': '宝信软件',
        '300496': '中科创达',
        '000938': '紫光股份',
        '300290': '荣科科技',
        '300674': '宇信科技',
        '300059': '东方财富'
        # '600584': '长电科技',
        # '301050': '雷电微力',
        # '002371': '北方华创',
        #
        # '300077': '国民技术',
        # '688256': '寒武纪-U',
        # '688008': '澜起科技',
        # '688981': '中芯国际',
        # '688041': '海光信息'
        
    }

    # for i in range(1,10):       # 爬取全部187页数据，设置为200页，确保都覆盖
    #     # url = 'https://s.askci.com/stock/h/ci0000001208-0?reportTime=2023-12-31&pageNum={i}#QueryCondition'.format(i=i) #港股
    #     url = 'https://s.askci.com/stock/a/ci0000001523-0?reportTime=2024-03-31&pageNum={i}#QueryCondition'.format(i=i) #A股
    #     page = ChromiumPage()
    #     page.get(url)
    #     data_list = []
    #     i = 0
    #     trs = page.eles('css:#myTable04 > tbody > tr')
    #     # print(trs)
    #     for tr_index in range(len(trs)):
    #         tr = trs[tr_index]
    #         # print(tr)
    #         row_data = get_row_data_H(tr)
    #         if row_data['股票代码'] in A_stock:
    #             url_financial = 'https://s.askci.com/stock/financialreport/' + row_data['股票代码']
    #             url_profit = 'https://s.askci.com/stock/financialreport/' + row_data['股票代码'] + '/profit'
    #             url_cashflow = 'https://s.askci.com/stock/financialreport/' + row_data['股票代码'] + '/cashflow'
    #
    #             print(url_financial)
    #             financial_df = pd.read_html(url_financial)[0]
    #             profit_df = pd.read_html(url_profit)[0]
    #             print(url_cashflow)
    #             cashflow_df = pd.read_html(url_cashflow)[0]
    #
    #             stock_name = row_data['股票简称']
    #             stock_name = stock_name.replace('*', '')
    #             # 格式化文件名
    #             file_name = '汽车-{}历史数据.xlsx'.format(stock_name)
    #
    #
    #
    #
    #
    #             # 使用 ExcelWriter 保存数据到指定的 Excel 文件和多个工作表
    #             with pd.ExcelWriter(file_name) as writer:
    #                 financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(stock_name))
    #                 profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(stock_name))
    #                 cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(stock_name))



    for stock_code, company_name in A_stock.items():#A stock
        print(company_name)
        url_financial = 'https://s.askci.com/stock/financialreport/' + stock_code
        url_profit = 'https://s.askci.com/stock/financialreport/' + stock_code + '/profit'
        url_cashflow = 'https://s.askci.com/stock/financialreport/' + stock_code + '/cashflow'

        print(url_financial)
        financial_df = pd.read_html(url_financial)[0]
        profit_df = pd.read_html(url_profit)[0]
        cashflow_df = pd.read_html(url_cashflow)[0]

        file_name = '互联网-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





