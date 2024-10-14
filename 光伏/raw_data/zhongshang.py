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
        '300274': '阳光电源',
        '603398': '沐邦高科',
        '688599': '天合光能',
        '688303': '大全能源',
        '002897': '意华股份',
        '688223': '晶科能源',
        '300716': '泉为科技',
        '002865': '钧达股份',
        '688408': '中信博',
        '688503': '聚和材料',
        '600151': '航天机电',
        '688598': '金博股份',
        '688560': '明冠新材',
        '002309': 'ST中利',
        '688717': '艾罗能源',
        '603396': '金辰股份',

        '301046': '能辉科技',
        '835985': '海泰新能',
        '600876': '凯盛新能',
        '002623': '亚玛顿',
        '600481': '双良节能',
        '603628': '清源股份',

        # '688726': '拉普拉斯',

        '688147': '微导纳米',
        '688033': '天宜上佳',
        '834770': '艾能聚',
        '301168': '通灵股份',

        '688429': '时创能源',
        '002056': '横店东磁',
        '301278': '快可电子',
        '603185': '弘元绿能',
        '839167': '同享科技',
        '300029': 'ST天龙',
        '002459': '晶澳科技',
        '300317': '珈伟新能',
        '603330': '天洋新材',
        '301266': '宇邦新材',
        '300345': '华民股份',

        '300776': '帝尔激光',
        '300051': '琏升科技',
        '603212': '赛伍技术',
        '688472': '阿特斯',
        '603778': '国晟科技',
        '600537': '亿晶光电',
        '601908': '京运通',
        '300842': '帝科股份',
        '688348': '昱能科技',

        '600207': '安彩高科',
        '688680': '海优新材',

        '001269': '欧晶科技',
        '300305': '裕兴股份',
        '603381': '永臻股份',
        '603105': '芯能科技',
        '002218': '拓日新能',
        '300093': '金刚光伏',
        '000821': '京山轻机',
        '300125': 'ST聆达',
        '688556': '高测股份',
        '300827': '上能电气',
        '300316': '晶盛机电',

        '300393': '中来股份',
        '688032': '禾迈股份',
        '300751': '迈为股份',
        '600732': '爱旭股份',
        '002506': '协鑫集成',
        '605117': '德业股份',
        '688390': '固德威',
        '002617': '露笑科技',
        '300724': '捷佳伟创',
        '688516': '奥特维',
        '601865': '福莱特',

        '300118': '东方日升',
        '300763': '锦浪科技',
        '600438': '通威股份',
        '603806': '福斯特',
        '002129': 'TCL中环',
        '601012': '隆基绿能'

        
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

        file_name = '光伏-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





