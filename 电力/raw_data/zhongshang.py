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

        # '600900': '长江电力',
        # '601985': '中国核电',
        # '600886': '国投电力',
        # '600795': '国电电力',
        # '003816': '中国广核',
        # '600863': '内蒙华电',
        # '600905': '三峡能源',
        # '000027': '深圳能源',
        # '600642': '申能股份',
        # '600578': '京能电力',
        # '600157': '永泰能源',
        # '600023': '浙能电力',
        # '000966': '长源电力',
        # '001286': '陕西能源',
        # '600163': '中闽能源',
        # '000899': '赣能股份',
        # '600509': '天富能源',
        #
        # '601016': '节能风电',
        # '600011': '华能国际',
        # '000591': '太阳能',
        # '000767': '晋控电力',
        # '600719': '大连热电',
        # '600396': '华电辽能',
        # '600116': '三峡水利',
        # '605028': '世茂能源',
        # '600969': '郴电国际',
        # '600868': '梅雁吉祥',

        # '605580': '恒盛能源',
        #
        # '600505': '西昌电力',

        # '600032': '浙江新能',
        #
        # '000862': '银星能源',

        # '900937': '华电B股',---no

        # '200037': '深南电B',---no
        # '000722': '湖南发展',
        #
        # '000958': '电投产融',
        # '600149': '廊坊发展',

        # '900957': '凌云Ｂ股',---no

        # '001258': '立新能源',

        # '200539': '粤电力Ｂ',---no

        '600726': '华电能源',
        '600744': '华银电力',
        '600982': '宁波能源',
        '600027': '华电国际',
        '001289': '龙源电力',
        '601619': '嘉泽新能',
        '603693': '江苏新能',
        '001896': '豫能控股',
        '000040': 'ST旭蓝',
        '600212': '绿能慧充',
        '000601': '韶能股份',
        '002039': '黔源电力',

        '000993': '闽东电力',
        '000037': '深南电A',
        '002616': '长青集团',
        '000155': '川能动力',
        '002608': '江苏国信',
        '600644': '乐山电力',
        '600452': '涪陵电力',
        '600101': '明星电力',
        '000531': '穗恒运Ａ',
        '000883': '湖北能源',
        '000537': '中绿电',
        '600780': '通宝能源',
        '000690': '宝新能源',
        '600821': '金开新能',
        '000791': '甘肃能源',
        '000600': '建投能源',
        '600674': '川投能源',

        '600098': '广州发展',
        '600025': '华能水电',
        '600995': '南网储能',
        '000543': '皖能电力',
        '002015': '协鑫能科',
        '000875': '吉电股份',
        '600310': '广西能源',
        '600236': '桂冠电力',
        '000539': '粤电力Ａ',
        '601778': '晶科科技',
        '601991': '大唐发电',
        '600483': '福能股份',
        '600021': '上海电力'

        
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

        file_name = '电力-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





