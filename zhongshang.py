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
        # '601127': '赛力斯',
        # '600733': '北汽蓝谷',
        # '601777': '力帆科技',
        # '600418': '江淮汽车',
        # '000572': '海马汽车',
        # '601238': '广汽集团',
        # '600104': '上汽集团',
        '002594': '比亚迪',
        '600006': '东风汽车',
        # '000868': '安凯客车',
        # '000625': '长安汽车',
        # '601633': '长城汽车',
        # '600166': '福田汽车',
        # '600213': 'ST亚星',
        # '000980': '众泰汽车',
        # '000550': '江铃汽车',
        # '600375': 'ST汉马',
        # '600303': 'ST曙光',
        # '301039': '中集车辆',
        # '000800': '一汽解放',
        # '000951': '中国重汽',
        # '600066': '宇通客车',
        # '600686': '金龙汽车',
        # '000957': '中通客车'
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

        file_name = '汽车-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





