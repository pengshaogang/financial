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

        # '000062': '深圳华强',
        # '002199': '东晶电子',
        # '002384': '东山精密',
        # '300936': '中英科技',
        # '603186': '华正新材',
        # '920008': '成电光信',--no
        # '002869': '金溢科技',
        # '600237': '铜峰电子',
        # '688629': '华丰科技',
        # '600183': '生益科技',
        # '688582': '芯动联科',
        #
        # '002025': '航天电器',
        #
        # '688511': '天微电子',
        # '600353': '旭光电子',
        # '301383': '天键股份',
        # '300868': '杰美特',
        #
        #
        #
        # '000733': '振华科技',
        #
        # '688800': '瑞可达',
        # '301251': '威尔高',
        # '603936': '博敏电子',
        # '300460': '惠伦晶体',
        # '300220': '金运激光',
        # '871857': '泓禧科技',
        # '871981': '晶赛科技',
        # '300656': '民德电子',
        # '688776': '国光电气',
        # '301577': '美信科技',
        # '002815': '崇达技术',
        # '688603': '天承科技',
        # '837212': '智新电子',
        # '301319': '唯特偶',
        # '688375': '国博电子',
        # '688496': '清越科技',
        # '002141': 'ST贤丰',
        # '688371': '菲沃泰',
        # '603375': '盛景微',
        # '688053': '思科瑞',
        # '301176': '逸豪新材',
        # '872374': '云里物里',
        # '605058': '澳弘电子',
        # '002137': '实益达',
        # '603678': '火炬电子',
        # '002138': '顺络电子',
        # '838701': '豪声电子',
        # '834950': '迅安科技',
        # '603115': '海星股份',
        # '301389': '隆扬电子',
        # '301280': '珠城科技',
        # '002636': '金安国纪',
        # '900938': '海科B',--no
        # '833914': '远航精密',
        #
        # '832110': '雷特科技',
        # '200413': 'ST东旭B',--no
        # '200020': '深华发Ｂ',--no
        # '200725': '京东方Ｂ',--no
        # '688210': '统联精密',
        # '200541': '粤照明Ｂ',--no
        '603989': '艾华集团',
        '300726': '宏达电子',
        '301328': '维峰电子',
        '430718': '合肥高科',
        '836395': '朗鸿科技',
        '688662': '富信科技',
        '688093': '世华科技',
        '301359': '东南电子',
        '301021': '英诺激光',
        '000636': '风华高科',
        '002134': '天津普林',
        '688655': '迅捷兴',
        '002724': '海洋王',
        '600563': '法拉电子',
        '002161': '远望谷',
        '300964': '本川智能',
        '870357': '雅葆轩',
        '001298': '好上好',
        '600884': '杉杉股份',
        '688020': '方邦股份',
        '002214': '大立科技',
        '301041': '金百泽',
        '002859': '洁美科技',
        '301566': '达利凯普',
        '600478': '科力远',
        '300852': '四会富仕',
        '002388': '新亚制程',
        '688489': '三未信安',

        '603386': '骏亚科技',
        '301132': '满坤科技',
        '688519': '南亚新材',
        '301329': '信音电子',
        '002484': '江海股份',
        '301285': '鸿日达',
        '688539': '高华科技',
        '688143': '长盈通',
        '000823': '超声电子',
        '600288': '大恒科技',
        '300991': '创益通',
        '301413': '安培龙',
        '301366': '一博科技',
        '002913': '奥士康',
        '603267': '鸿远电子',
        '688035': '德邦科技',
        '300410': '正业科技',
        '301217': '铜冠铜箔',
        '688103': '国力股份',
        '603052': '可川科技',
        '603328': '依顿电子',
        '837821': '则成电子',
        '688322': '奥比中光',
        '301282': '金禄电子',
        '605258': '协和电子',
        '300516': '久之洋',
        '300227': '光韵达',
        '688025': '杰普特',
        '300747': '锐科激光',
        '300903': '科翔股份',
        '603228': '景旺电子',
        '301320': '豪江智能',
        '300408': '三环集团',
        '300739': '明阳电路',

        '300814': '中富电路',
        '001389': '广合科技',
        '688183': '生益电子',
        '301536': '星宸科技',
        '002922': '伊戈尔',
        '002729': '好利科技',
        '002579': '中京电子',
        '301517': '陕西华达',
        '002436': '兴森科技',
        '002179': '中航光电',
        '002414': '高德红外',
        '600601': '方正科技',
        '603738': '泰晶科技',
        '301571': '国科天成',
        '002916': '深南电路',
        '300657': '弘信电子',
        '300475': '香农芯创',
        '300124': '汇川技术',
        '002938': '鹏鼎控股',
        '603920': '世运电路',
        '000988': '华工科技',
        '300476': '胜宏科技',
        '300474': '景嘉微',
        '002463': '沪电股份'
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

        file_name = '电子元件-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





