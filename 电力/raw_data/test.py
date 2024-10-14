import requests
from bs4 import BeautifulSoup
import pandas as pd

def fetch_data(url):
    # 使用requests获取网页内容
    response = requests.get(url)
    if response.status_code == 200:
        # 使用Beautiful Soup解析HTML
        soup = BeautifulSoup(response.text, 'html.parser')
        # 寻找网页中的表格元素
        tables = soup.find_all('table')
        if tables:
            # 假设第一个表格是我们需要的
            return pd.read_html(str(tables[0]))[0]
        else:
            print("No tables found on the webpage.")
            return None
    else:
        print(f"Failed to retrieve the webpage: Status code {response.status_code}")
        return None

def main(stock_code):
    url_financial = f'https://s.askci.com/stock/financialreport/{stock_code}'
    url_profit = f'https://s.askci.com/stock/financialreport/{stock_code}/profit'
    url_cashflow = f'https://s.askci.com/stock/financialreport/{stock_code}/cashflow'

    # 打印生成的URLs
    print("URL for Financial Report:", url_financial)
    print("URL for Profit Report:", url_profit)
    print("URL for Cash Flow Report:", url_cashflow)

    # 获取数据
    financial_df = fetch_data(url_financial)
    profit_df = fetch_data(url_profit)
    cashflow_df = fetch_data(url_cashflow)

    # 输出数据表（如果存在）
    if financial_df is not None:
        print("Financial Report:")
        print(financial_df)
    if profit_df is not None:
        print("Profit Report:")
        print(profit_df)
    if cashflow_df is not None:
        print("Cash Flow Report:")
        print(cashflow_df)

# 替换'stock_code'为实际的股票代码
main('600032')
