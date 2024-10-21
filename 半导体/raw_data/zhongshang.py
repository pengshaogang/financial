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

        # '002049': '紫光国微',
        # '688361': '中科飞测',
        # '688709': '成都华微',
        # '688270': '臻镭科技',
        # '603501': '韦尔股份',
        # '688498': '源杰科技',
        # '300604': '长川科技',
        # '688200': '华峰测控',
        # '688409': '富创精密',
        # '301369': '联动科技',
        # '430139': '华岭股份',
        #
        # '688213': '思特威-W',
        #
        # '688072': '拓荆科技',
        # '002185': '华天科技',
        # '688259': '创耀科技',
        # '688099': '晶晨股份',



        '688123': '聚辰股份',

        '603061': '金海通',
        '688352': '颀中科技',
        '688720': '艾森股份',
        '688332': '中科蓝讯',
        '688416': '恒烁股份',
        '002213': '大为股份',
        '603160': '汇顶科技',
        '688220': '翱捷科技',
        '688381': '帝奥微',
        '688486': '龙迅股份',
        '300373': '扬杰科技',
        '688689': '银河微电',
        '688061': '灿瑞科技',
        '688167': '炬光科技',
        '600360': 'ST华微',
        '688432': '有研硅',
        '688216': '气派科技',
        '688589': '力合微',
        '688153': '唯捷创芯',
        '301297': '富乐德',
        '688172': '燕东微',
        '688045': '必易微',
        '603933': '睿能科技',
        '688279': '峰岹科技',
        '688138': '清溢光电',
        '688173': '希荻微',
        '688286': '敏芯股份',
        '688130': '晶华微',
        '688325': '赛微微电',
        '688230': '芯导科技',
        '688512': '慧智微-U',
        '301095': '广立微',
        '688652': '京仪装备',
        '688391': '钜泉科技',

        '688591': '泰凌微',
        '688798': '艾为电子',
        '688135': '利扬芯片',
        '003026': '中晶科技',
        '688702': '盛科通信',
        '688711': '宏微科技',
        '688478': '晶升股份',
        '688530': 'XD欧莱新',
        '300456': '赛微电子',
        '688584': '上海合晶',
        '301348': '蓝箭电子',
        '688403': '汇成股份',
        '688484': '南芯科技',
        '688653': '康希通信',
        '688233': '神工股份',
        '688049': '炬芯科技',
        '001270': '铖昌科技',
        '688018': '乐鑫科技',
        '300327': '中颖电子',
        '688728': '格科微',
        '688620': '安凯微',
        '688458': '美芯晟',
        '688262': '国芯科技',
        '300706': '阿石创',
        '688766': '普冉股份',
        '688362': '甬矽电子',
        '688521': '芯原股份',
        '688107': '安路科技',
        '688693': '锴威特',
        '603068': '博通集成',
        '688401': '路维光电',
        '688234': '天岳先进',
        '002119': '康强电子',
        '000670': '盈方微',

        '688141': '杰华特',
        '688261': '东微半导',
        '001309': '德明利',
        '688002': '睿创微纳',
        '688082': '盛美上海',
        '688595': '芯海科技',
        '688439': '振华风光',
        '688699': '明微电子',
        '688209': '英集芯',
        '688661': '和林微纳',
        '688593': '新相微',
        '605588': '冠石科技',
        '688691': '灿芯股份',
        '688048': '长光华芯',
        '688052': '纳芯微',
        '688508': '芯朋微',
        '300042': '朗科科技',
        '688515': '裕太微-U',
        '688110': '东芯股份',
        '605358': '立昂微',
        '300458': '全志科技',
        '688120': '华海清科',
        '002077': '大港股份',
        '688601': '力芯微',
        '600877': '电科芯片',
        '688396': '华润微',
        '300666': '江丰电子',
        '688252': '天德钰',
        '688608': '恒玄科技',
        '605111': '新洁能',
        '688380': '中微半导',
        '300613': '富瀚微',
        '688536': '思瑞浦',
        '688469': '芯联集成',

        '688721': '龙图光罩',
        '600171': '上海贝岭',
        '300831': '派瑞股份',
        '688368': '晶丰明源',
        '688037': '芯源微',
        '688347': '华虹公司',
        '300046': '台基股份',
        '688535': '华海诚科',
        '300672': '国科微',
        '002079': '苏州固锝',
        '300053': '航宇微',
        '603005': '晶方科技',
        '688372': '伟测科技',
        '603290': '斯达半导',
        '603893': '瑞芯微',
        '688047': '龙芯中科',
        '688385': '复旦微电',
        '300661': '圣邦股份',
        '300671': '富满微',
        '300493': '润欣科技',
        '300223': '北京君正',
        '301308': '江波龙',
        '688249': '晶合集成',
        '600460': '士兰微',
        '300623': '捷捷微电',
        '002156': '通富微电',
        '603986': '兆易创新',
        '688126': '沪硅产业',
        '688012': '中微公司',
        '300782': '卓胜微',
        '688525': '佰维存储',
        '600584': '长电科技',
        '301050': '雷电微力',
        '002371': '北方华创',

        '300077': '国民技术',
        '688256': '寒武纪-U',
        '688008': '澜起科技',
        '688981': '中芯国际',
        '688041': '海光信息'
        
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

        file_name = '半导体-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





