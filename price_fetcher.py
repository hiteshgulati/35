import quandl
from datetime import datetime
from datetime import timedelta
import pandas as pd
quandl.ApiConfig.api_key = 'NjFyz5PTC_XMypj9dL9f'
ticker = "IBULISL"
database = "NSE/"
quandl_code = database + "/" + ticker
yesterday = datetime.strftime(datetime.now() - timedelta(1), '%Y-%m-%d')
print(yesterday)
fetch_date = yesterday
# quandl.get(quandl_code, start_date = fetch_date, end_date = fetch_date).Close[0]
price_file = pd.read_excel("Price Reporter.xlsx", sheet_name = "Input")


slice = {"NSE_Ticker":"", "Remarks":"", "DOA": datetime(1,1,1), "DOA_Open":0,"DOA_Close":0,
         "Date1": datetime(1, 1, 1), "Date1_Open": 0, "Date1_Close": 0,
         "Date2": datetime(1, 1, 1), "Date2_Open": 0, "Date2_Close": 0,
         "Date3": datetime(1, 1, 1), "Date3_Open": 0, "Date3_Close": 0,
         "Date4": datetime(1, 1, 1), "Date4_Open": 0, "Date4_Close": 0,
         "Date5": datetime(1, 1, 1), "Date5_Open": 0, "Date5_Close": 0,
         "Date6": datetime(1, 1, 1), "Date6_Open": 0, "Date6_Close": 0,
         "Date7": datetime(1, 1, 1), "Date7_Open": 0, "Date7_Close": 0,
         "Date8": datetime(1, 1, 1), "Date8_Open": 0, "Date8_Close": 0,
         "Date9": datetime(1, 1, 1), "Date9_Open": 0, "Date9_Close": 0,
         "Date10": datetime(1, 1, 1), "Date10_Open": 0, "Date10_Close": 0}
outputdf = pd.DataFrame(columns=list(slice))
for i, item in price_file.iterrows():
    slice['NSE_Ticker'] = item.NSE_Ticker
    slice['DOA'] = item.Date_of_Action
    quandl_code = database + slice['NSE_Ticker']
    fetch_date = slice['DOA'].strftime('%Y-%m-%d')
    slice['DOA_Open'] = quandl.get(quandl_code, start_date = fetch_date, end_date = fetch_date).Open[0]
    slice['DOA_Close'] = quandl.get(quandl_code, start_date=fetch_date, end_date=fetch_date).Close[0]

    print(item.NSE_Ticker)

