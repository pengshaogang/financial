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
        # '603882': '金域医学',
        # '603108': '润达医疗',
        # '000516': '国际医学',
        # '301333': '诺思格',
        # '300759': '康龙化成',
        # '688222': '成都先导',
        # '300143': '盈康生命',
        # '301293': '三博脑科',
        # '301033': '迈普医学',
        # '301235': '华康医疗',
        # '301257': '普蕊斯',
        # '688621': '阳光诺和',
        # '301096': '百诚医药',
        # '300244': '迪安诊断',
        # '688315': '诺禾致源',
        # '301201': '诚达药业',
        #
        # '002524': '光正眼科',
        # '301103': '何氏眼科',
        # '605186': '健麾信息',

        # '200028': '一致Ｂ',

        '301267': '华厦眼科',
        '002172': '澳洋健康',



        '002622': '皓宸医疗',
        '300404': '博济医药',
        '688246': '嘉和美康',
        '600327': '大东方',

        '300149': '睿智医药',
        '301126': '达嘉维康',
        '600721': '百花医药',
        '002044': '美年健康',
        '301520': '万邦医药',
        '002219': '新里程',
        '301080': '百普赛斯',
        '000504': '南华生物',
        '688238': '和元生物',
        '301239': '普瑞眼科',
        '688202': '美迪西',

        '002173': '创新医疗',
        '688076': '诺泰生物',
        '603127': '昭衍新药',
        '300347': '泰格医药',
        '002821': '凯莱英',
        '603259': '药明康德',
        '600763': '通策医疗',
        '300015': '爱尔眼科',
        # '688348': '昱能科技',
        #
        # '600207': '安彩高科',
        # '688680': '海优新材',
        #
        # '001269': '欧晶科技',
        # '300305': '裕兴股份',
        # '603381': '永臻股份',
        # '603105': '芯能科技',
        # '002218': '拓日新能',
        # '300093': '金刚光伏',
        # '000821': '京山轻机',
        # '300125': 'ST聆达',
        # '688556': '高测股份',
        # '300827': '上能电气',
        # '300316': '晶盛机电',
        #
        # '300393': '中来股份',
        # '688032': '禾迈股份',
        # '300751': '迈为股份',
        # '600732': '爱旭股份',
        # '002506': '协鑫集成',
        # '605117': '德业股份',
        # '688390': '固德威',
        # '002617': '露笑科技',
        # '300724': '捷佳伟创',
        # '688516': '奥特维',
        # '601865': '福莱特',
        #
        # '300118': '东方日升',
        # '300763': '锦浪科技',
        # '600438': '通威股份',
        # '603806': '福斯特',
        # '002129': 'TCL中环',
        # '601012': '隆基绿能'

        
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

        file_name = '医疗-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





