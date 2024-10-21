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

        # '836239': '长虹能源',
        # '301152': '天力锂能',
        # '688779': '五矿新能',
        # '002805': '丰元股份',
        # '301511': '德福科技',
        # '002850': '科达利',
        # '603906': '龙蟠科技',
        # '601311': '骆驼股份',
        # '002245': '蔚蓝锂芯',
        # '002759': '天际股份',
        # '873152': '天宏锂电',
        #
        # '688275': '万润新能',
        #
        # '831627': '力王股份',
        # '301150': '中一科技',
        # '688573': '信宇人',
        # '835237': '力佳科技',
        #
        #
        #
        # '688388': '嘉元科技',
        #
        # '830809': '安达科技',
        # '605378': '野马电池',
        # '600847': '万里股份',
        # '688638': '誉辰智能',
        # '001283': '豪鹏科技',
        # '301222': '浙江恒威',
        # '688778': '厦钨新能',
        # '300953': '震裕科技',
        # '688184': '帕瓦股份',
        # '603032': '德新科技',
        # '688567': '孚能科技',
        # '600241': '时代万恒',
        '688345': '博力威',
        '603031': '安孚科技',
        '300530': '领湃科技',
        '835185': '贝特瑞',
        '301325': '曼恩斯特',
        '002580': '圣阳股份',
        '833523': '德瑞锂电',
        '301210': '金杨股份',
        '600152': '维科技术',
        '688148': '芳源股份',
        '688006': 'XD杭可科',
        '688707': '振华新材',
        '688499': '利元亨',
        '688339': '亿华通-U',
        '688772': '珠海冠宇',
        '300648': '星云股份',
        '000049': '德赛电池',
        '688063': '派能科技',
        '688005': '容百科技',
        '300619': '金银河',
        '300340': '科恒股份',
        '600110': '诺德股份',

        '301121': '紫建电子',
        '301587': '中瑞股份',
        '300890': '翔丰华',
        '688819': '天能股份',
        '301238': '瑞泰新材',
        '002733': '雄韬股份',
        '301358': '湖南裕能',
        '688155': '先惠技术',
        '300037': '新宙邦',
        '300919': '中伟股份',
        '300568': '星源材质',
        '300457': '赢合科技',
        '300207': '欣旺达',
        '002812': '恩捷股份',
        '301487': '盟固利',
        '603659': '璞泰来',
        '002074': '国轩高科',
        '300450': '先导智能',
        '300438': '鹏辉能源',
        '300769': '德方纳米',
        '300073': '当升科技',
        '300014': '亿纬锂能',
        '300068': '南都电源',
        '300750': '宁德时代'
        # '688766': '普冉股份',
        # '688362': '甬矽电子',
        # '688521': '芯原股份',
        # '688107': '安路科技',
        # '688693': '锴威特',
        # '603068': '博通集成',
        # '688401': '路维光电',
        # '688234': '天岳先进',
        # '002119': '康强电子',
        # '000670': '盈方微',
        #
        # '688141': '杰华特',
        # '688261': '东微半导',
        # '001309': '德明利',
        # '688002': '睿创微纳',
        # '688082': '盛美上海',
        # '688595': '芯海科技',
        # '688439': '振华风光',
        # '688699': '明微电子',
        # '688209': '英集芯',
        # '688661': '和林微纳',
        # '688593': '新相微',
        # '605588': '冠石科技',
        # '688691': '灿芯股份',
        # '688048': '长光华芯',
        # '688052': '纳芯微',
        # '688508': '芯朋微',
        # '300042': '朗科科技',
        # '688515': '裕太微-U',
        # '688110': '东芯股份',
        # '605358': '立昂微',
        # '300458': '全志科技',
        # '688120': '华海清科',
        # '002077': '大港股份',
        # '688601': '力芯微',
        # '600877': '电科芯片',
        # '688396': '华润微',
        # '300666': '江丰电子',
        # '688252': '天德钰',
        # '688608': '恒玄科技',
        # '605111': '新洁能',
        # '688380': '中微半导',
        # '300613': '富瀚微',
        # '688536': '思瑞浦',
        # '688469': '芯联集成',
        #
        # '688721': '龙图光罩',
        # '600171': '上海贝岭',
        # '300831': '派瑞股份',
        # '688368': '晶丰明源',
        # '688037': '芯源微',
        # '688347': '华虹公司',
        # '300046': '台基股份',
        # '688535': '华海诚科',
        # '300672': '国科微',
        # '002079': '苏州固锝',
        # '300053': '航宇微',
        # '603005': '晶方科技',
        # '688372': '伟测科技',
        # '603290': '斯达半导',
        # '603893': '瑞芯微',
        # '688047': '龙芯中科',
        # '688385': '复旦微电',
        # '300661': '圣邦股份',
        # '300671': '富满微',
        # '300493': '润欣科技',
        # '300223': '北京君正',
        # '301308': '江波龙',
        # '688249': '晶合集成',
        # '600460': '士兰微',
        # '300623': '捷捷微电',
        # '002156': '通富微电',
        # '603986': '兆易创新',
        # '688126': '沪硅产业',
        # '688012': '中微公司',
        # '300782': '卓胜微',
        # '688525': '佰维存储',
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

        file_name = '电池-{}历史数据.xlsx'.format(company_name)

        with pd.ExcelWriter(file_name) as writer:
            financial_df.to_excel(writer, index=False, sheet_name='{}资产负债表'.format(company_name))
            profit_df.to_excel(writer, index=False, sheet_name='{}利润表'.format(company_name))
            cashflow_df.to_excel(writer, index=False, sheet_name='{}现金流量表'.format(company_name))





