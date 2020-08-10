import requests
from bs4 import BeautifulSoup
import traceback
import pandas as pd
import datetime
import os
import csv


# 일별 주가 읽기 함수
def parse_page(cd, pg):
    try:
        url = 'http://finance.naver.com/item/sise_day.nhn?code={code}&page={page}'.format(code=cd, page=pg)
        res = requests.get(url)
        _soap = BeautifulSoup(res.text, 'lxml')
        _df = pd.read_html(str(_soap.find("table")), header=0)[0]
        _df = _df.dropna()
        return _df
    except Exception as e:
        traceback.print_exc()
    return None


# 차익 계산시 기준 - 종가, 시가, 고가, 저가 등
GAIN_CALCULATE_BASIS = '종가'

# csv 파일로 주식 매매 내역 입력 받기
f = open('C:/Users/minjae/Downloads/stockdata.csv', 'r', encoding='utf-8')
rdr = csv.reader(f)

# 차익, 수익률 등 출력할 csv 파일
# today = datetime.datetime.strftime(datetime.datetime.today(), '%Y%m%d')
# fw = open('C:/Users/minjae/Downloads/stock_output_' + today + '.csv', 'w', encoding='utf-8', newline='')
# wr = csv.writer(fw)

result = []

# 주식 매매 내역 분석
for line in rdr:
    name = line[0]  # 매매자 이름
    stock_name = line[1]    # 종목이름
    stock_code = line[2]  # 종목코드
    buy_date = line[3]   # 매수 날짜
    buy_cnt = line[4]   # 매수 개수
    sell_date = line[5]  # 매도 날짜
    sell_cnt = line[6]  # 매도 개수

    buy_date_splited = buy_date.split('.')
    b_year = int(buy_date_splited[0])   # 매수 년
    b_month = int(buy_date_splited[1])  # 매수 월
    b_day = int(buy_date_splited[2])    # 매수 일

    sell_date_splited = sell_date.split('.')
    s_year = int(sell_date_splited[0])  # 매도 년
    s_month = int(sell_date_splited[1]) # 매도 월
    s_day = int(sell_date_splited[2])   # 매도 일

    b_cnt = int(buy_cnt)    # 매수 개수
    s_cnt = int(sell_cnt)   # 매도 개수

    # 네이버 금융에서 일별 주가 크롤링
    code = stock_code

    # 일별 주가 페이지 마지막 번호 가져오기
    url = 'http://finance.naver.com/item/sise_day.nhn?code={code}'.format(code=code)
    res = requests.get(url)
    res.encoding = 'utf-8'

    soap = BeautifulSoup(res.text, 'lxml')

    el_table_navi = soap.find("table", class_="Nnavi")
    el_td_last = el_table_navi.find("td", class_="pgRR")
    pg_last = el_td_last.a.get('href').rsplit('&')[1]
    pg_last = pg_last.split('=')[1]
    pg_last = int(pg_last)

    str_date_buy = datetime.datetime.strftime(datetime.datetime(year=b_year, month=b_month, day=b_day), '%Y.%m.%d')
    str_date_sell = datetime.datetime.strftime(datetime.datetime(year=s_year, month=s_month, day=s_day), '%Y.%m.%d')

    df = None
    df_buy = None
    df_sell = None

    # 매수, 매도한 날 일별 주가 데이터만 가져옴
    for page in range(1, pg_last + 1):
        _df = parse_page(code, page)

        _df_buy = _df[_df['날짜'] == str_date_buy]
        _df_sell = _df[_df['날짜'] == str_date_sell]

        if len(_df_buy) > 0:
            df_buy = _df_buy

        if len(_df_sell) > 0:
            df_sell = _df_sell

        if df_buy is not None and df_sell is not None:
            break

    # 매수한 날 주가 * 매수 개수
    cost = int(df_buy[GAIN_CALCULATE_BASIS]) * b_cnt
    # 매도한 날 주가 * 매도 개수
    income = int(df_sell[GAIN_CALCULATE_BASIS]) * s_cnt

    # 차익
    gain = income - cost

    # 수익률 = 100 * (차익 / 매수비용)
    rate = round(100 * float(gain)/float(cost), 3)

    result.append([name, stock_name, stock_code, GAIN_CALCULATE_BASIS,
                 buy_date, buy_cnt, cost,
                 sell_date, sell_cnt, income,
                 gain, rate,
                 int(df_buy['시가']), int(df_buy['종가']), int(df_buy['고가']), int(df_buy['저가']),
                 int(df_sell['시가']), int(df_sell['종가']), int(df_sell['고가']), int(df_sell['저가'])])


today = datetime.datetime.strftime(datetime.datetime.today(), '%Y%m%d')
base_dir = "C:/Users/minjae/Downloads"
file_nm = "stock_output_" + today + ".xlsx"

xlxs_dir = os.path.join(base_dir, file_nm)

rdf = pd.DataFrame(result, columns=[
    '이름', '종목명', '종목코드', '기준가',
    '매수날짜', '매수개수', '매수금액',
    '매도날짜', '매도개수', '매도금액',
    '차익', '수익률',
    '매수시가', '매수종가', '매수고가', '매수저가',
    '매도시가', '매도종가', '매도고가', '매도저가'])

rdf.to_excel(xlxs_dir,
             sheet_name=today, na_rep='NaN', header=True,
             index=False, index_label="번호",
             startrow=0, startcol=0, freeze_panes=(1, 0)
             )

    # # 결과값 csv 형태로 저장
    # wr.writerow([name, stock_code,
    #              buy_date, buy_cnt,
    #              sell_date, sell_cnt,
    #              gain, rate,
    #              int(df_buy['시가']), int(df_buy['고가']), int(df_buy['저가']),
    #              int(df_sell['시가']), int(df_sell['고가']), int(df_sell['저가'])])



