import time
import pandas as pd
from DrissionPage import ChromiumPage

def get_row_data(tr):
    tds = tr.eles('css:td')
    if not tds:
        return None
    link_element = tds[1].ele('css:a')
    link = link_element.attr('href').replace('http:', 'https:') if link_element and link_element.attr('href').startswith('http:') else link_element.attr('href')
    return {
        '序号': tds[0].text,
        '股票代码': tds[1].text,
        '股票名称': tds[2].text,
        '现价': tds[3].text,
        '涨跌幅': tds[4].text,
        '涨跌额': tds[5].text,
        '涨速 (%)': tds[6].text,
        '换手 (%)': tds[7].text,
        '量比': tds[8].text,
        '振幅': tds[9].text,
        '成交额': tds[10].text,
        '流通股': tds[11].text,
        '流通市值': tds[12].text,
        '市盈率': tds[13].text,
        '股票链接': link
    }

def fetch_company_details(page, link):
    page.get(link)
    time.sleep(2)
    dd_elements = page.eles('css:body > div.m_content > div:nth-child(2) > div:nth-child(3) > dl > dd')
    dt_elements = page.eles('css:body > div.m_content > div:nth-child(2) > div:nth-child(3) > dl > dt')
    dd_elements.pop(3)
    company_details = {}
    for dt, dd in zip(dt_elements, dd_elements):
        company_details[dt.text] = dd.text
    return company_details

def navigate_to_next_page(page):
    nextpage = page.ele('@@text():下一页@@class=changePage')
    if nextpage:
        nextpage.click(by_js=True)
        time.sleep(2)  # 等待页面加载
        return True
    return False


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
        '注册地址': tds[4].text,
        '主营业务收入(202312) ': tds[5].text,
        '净利润(202312) ': tds[6].text,
        '员工人数': tds[7].text,
        '上市日期': tds[8].text,
        '招股书': tds[9].text,
        '公司财报': tds[10].text,
        '行业分类': tds[11].text,
        '主营业务': tds[12].text,
        # '市盈率': tds[13].text,
        '股票链接': link
        }
    # else:
    #     return



if __name__ == '__main__':
    page = ChromiumPage()
    initial_url = 'https://q.10jqka.com.cn/thshy/detail/code/881125/'  # 汽车整车
    page.get(initial_url)
    data_list = []
    i = 0
    while True:
        trs = page.eles('css:#maincont > table > tbody > tr')
        print(len(trs))
        for tr_index in range(len(trs)):
            trs = page.eles('css:#maincont > table > tbody > tr')
            tr = trs[tr_index]
            print("tr = ", tr.text)
            row_data = get_row_data(tr)
            if row_data:
                company_details = fetch_company_details(page, row_data['股票链接'])
                page.back(1)
                for _ in range(i):
                    nextpage = page.ele('@@text():下一页@@class=changePage')
                    if nextpage:
                        nextpage.click(by_js=True)
                        time.sleep(2)
                combined_data = {**row_data, **company_details}
                data_list.append(combined_data)
        i += 1
        if not navigate_to_next_page(page):
            break
    df = pd.DataFrame(data_list)
    df.to_excel('stocks.xlsx', index=False)
    print("数据已保存到Excel文件。")

    # 清空数据列表，准备处理第二个页面
    data_list = []

    # # 定义第二个 URL 链接
    # initial_url_2 = 'https://s.askci.com/stock/h/ci0000001208-0'  # 第二个页面链接
    # myTable04 > tbody > tr:nth-child(8)
    # myTable04 > tbody > tr:nth-child(7)

    # 进入第二个页面
    page.get(initial_url)
    i = 0
    while True:
        trs = page.eles('css:#myTable04 > tbody > tr')
        print(len(trs))
        for tr_index in range(len(trs)):
            trs = page.eles('css:#myTable04 > tbody > tr')
            tr = trs[tr_index]
            print("tr = ", tr.text)

            row_data = get_row_data(tr)
            # if row_data['行业分类'] == '汽车制造':
            #     data_list.append(row_data)
            if row_data:
                company_details = fetch_company_details(page, row_data['股票链接'])
                page.back(1)
                for _ in range(i):
                    nextpage = page.ele('@@text():下一页@@class=changePage')
                    # nextpage = page.ele('css:.pageBtnWrap a[title="下一页"]')
                    # m-page > a:nth-child(3)

                    # kkpager > div:nth-child(1) > span.pageBtnWrap > a:nth-child(5)

                    if nextpage:
                        nextpage.click(by_js=True)
                        time.sleep(2)
                combined_data = {**row_data, **company_details}

                data_list.append(combined_data)
        i += 1
        if not navigate_to_next_page(page):
            break

    # 将第二个页面的数据存储到 Excel 文件
    df2 = pd.DataFrame(data_list)
    df2.to_excel('stocks_2.xlsx', index=False)
    print("第二个页面数据已保存到Excel文件。")

