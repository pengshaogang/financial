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

        # '603626': '科森科技',
        # '300793': '佳禾智能',
        # '002655': '共达电声',
        # '300602': '飞荣达',
        # '300686': '智动力',
        # '688036': '传音控股',
        # '002843': '泰嘉股份',
        # '603327': '福蓉科技',
        # '603629': '利通电子',
        # '688007': '光峰科技',
        # '002681': '奋达科技',
        #
        # '002981': '朝阳科技',
        #
        # '002855': '捷荣技术',
        # '002925': '盈趣科技',
        # '301387': '光大同创',
        # '002866': '传艺科技',
        #
        #
        #
        # '300976': '达瑞电子',
        #
        # '002577': '雷柏科技',
        # '300032': '金龙机电',
        # '002045': '国光电器',
        # '002861': '瀛通通讯',
        # '300279': '和晶科技',
        # '300679': '电连技术',
        # '002782': '可立克',
        # '300543': '朗科智能',
        # '300956': '英力股份',
        # '300822': '贝仕达克',
        # '301503': '智迪科技',
        # '831167': '鑫汇科',
        # '600898': 'ST美讯',
        # '301180': '万祥科技',
        # '833346': '威贸电子',
        # '301318': '维海德',
        # '300951': '博硕科技',
        # '300787': '海能实业',
        # '301626': '苏州天脉',
        # '301135': '瑞德智能',
        # '300866': '安克创新',
        # '002660': '茂硕电源',
        # '002993': '奥海科技',
        # '003028': '振邦智能',
        # '301486': '致尚科技',
        # '301086': '鸿富瀚',
        # '603595': '东尼电子',
        # '872190': '雷神科技',
        # '300916': '朗特智能',
        # '002881': '美格智能',
        # '001308': '康冠科技',
        # '832876': '慧为智能',
        # '300684': '中石科技',
        # '688678': '福立旺',
        #
        # '603380': '易德龙',
        # '301123': '奕东电子',
        # '603296': '华勤技术',
        # '688683': '莱尔科技',
        # '001314': '亿道信息',
        # '002937': '兴瑞科技',
        # '603890': '春秋电子',
        # '605277': '新亚电子',
        # '002351': '漫步者',
        # '300433': '蓝思科技',
        # '301067': '显盈科技',
        # '301578': '辰奕智能',
        # '603341': '龙旗科技',
        # '301606': '绿联科技',
        # '002369': '卓翼科技',
        # '002402': '和而泰',
        # '002139': '拓邦股份',
        # '301489': '思泉新材',
        # '300647': '超频三',
        # '300812': '易天股份',
        # '601231': '环旭电子',
        # '002888': '惠威科技',
        # '002841': '视源股份',
        # '688260': '昀冢科技',
        # '002947': '恒铭达',
        # '301567': '贝隆精密',
        # '002635': '安洁科技',
        # '300968': '格林精密',
        '002055': '得润电子',
        '600130': '波导股份',
        '300322': '硕贝德',
        '002885': '京泉华',
        '300843': '胜蓝股份',
        '300131': '英唐智控',

        '301326': '捷邦科技',
        '300136': '信维通信',
        '301182': '凯旺科技',
        '600745': '闻泰科技',
        '300256': '星星科技',
        '002241': '歌尔股份',
        '300847': '中船汉光',
        '300115': '长盈精密',
        '000021': '深科技',
        '600203': '福日电子',
        '002475': '立讯精密',
        '300857': '协创数据',
        '300709': '精研科技',
        '300735': '光弘科技',
        '601138': '工业富联',
        '002600': '领益智造'

        # '300410': '正业科技',
        # '301217': '铜冠铜箔',
        # '688103': '国力股份',
        # '603052': '可川科技',
        # '603328': '依顿电子',
        # '837821': '则成电子',
        # '688322': '奥比中光',
        # '301282': '金禄电子',
        # '605258': '协和电子',
        # '300516': '久之洋',
        # '300227': '光韵达',
        # '688025': '杰普特',
        # '300747': '锐科激光',
        # '300903': '科翔股份',
        # '603228': '景旺电子',
        # '301320': '豪江智能',
        # '300408': '三环集团',
        # '300739': '明阳电路',

        # '300814': '中富电路',
        # '001389': '广合科技',
        # '688183': '生益电子',
        # '301536': '星宸科技',
        # '002922': '伊戈尔',
        # '002729': '好利科技',
        # '002579': '中京电子',
        # '301517': '陕西华达',
        # '002436': '兴森科技',
        # '002179': '中航光电',
        # '002414': '高德红外',
        # '600601': '方正科技',
        # '603738': '泰晶科技',
        # '301571': '国科天成',
        # '002916': '深南电路',
        # '300657': '弘信电子',
        # '300475': '香农芯创',
        # '300124': '汇川技术',
        # '002938': '鹏鼎控股',
        # '603920': '世运电路',
        # '000988': '华工科技',
        # '300476': '胜宏科技',
        # '300474': '景嘉微',
        # '002463': '沪电股份',
        # '300287': '飞利信',
        # '600845': '宝信软件',
        # '300496': '中科创达',
        # '000938': '紫光股份',
        # '300290': '荣科科技',
        # '300674': '宇信科技',
        # '300059': '东方财富'
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

        file_name = '消费电子-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





