import quandl
from datetime import datetime
from datetime import timedelta
import pandas as pd
quandl.ApiConfig.api_key = 'NjFyz5PTC_XMypj9dL9f'
database = "NSE/"

price_file = pd.read_excel("Price Reporter.xlsx", sheet_name = "Input")

def get_quote_from_quandl (quandl_code, fetch_date = datetime.today(), type = "Close"):
    fetch = fetch_date.strftime('%Y-%m-%d')
    try:
        quote = quandl.get(quandl_code, start_date=fetch, end_date=fetch)[type][0]
    except IndexError:
        quote = get_quote_from_quandl(quandl_code,fetch_date + timedelta(1), type)
    except quandl.errors.quandl_error.NotFoundError:
        quote = 0
    except:
        quote = -1
    return quote


def run():
    slice = {"NSE_Ticker":"", "Remarks":"", "DOA": datetime(1990,1,1), "DOA_Open":0,"DOA_Close":0, "Next_Day_Close": 0,
         "Third_Day_Close": 0, "Next_Week_Close": 0, "Next_Month_Close": 0, "Third_Month_Close": 0,
         "Date1": datetime(1990, 1, 1), "Date1_Open": 0, "Date1_Close": 0,
         "Date2": datetime(1990, 1, 1), "Date2_Open": 0, "Date2_Close": 0,
         "Date3": datetime(1990, 1, 1), "Date3_Open": 0, "Date3_Close": 0,
         "Date4": datetime(1990, 1, 1), "Date4_Open": 0, "Date4_Close": 0,
         "Date5": datetime(1990, 1, 1), "Date5_Open": 0, "Date5_Close": 0,
         "Date6": datetime(1990, 1, 1), "Date6_Open": 0, "Date6_Close": 0,
         "Date7": datetime(1990, 1, 1), "Date7_Open": 0, "Date7_Close": 0,
         "Date8": datetime(1990, 1, 1), "Date8_Open": 0, "Date8_Close": 0,
         "Date9": datetime(1990, 1, 1), "Date9_Open": 0, "Date9_Close": 0,
         "Date10": datetime(1990, 1, 1), "Date10_Open": 0, "Date10_Close": 0}
    outputdf = pd.DataFrame(columns=list(slice))
    print (outputdf.values)
    for i, item in price_file.iterrows():
        print ("Started: ", item.NSE_Ticker)
        slice['NSE_Ticker'] = item.NSE_Ticker
        slice['Remarks'] = item.Remarks
        slice['DOA'] = item.Date_of_Action
        quandl_code = database + slice['NSE_Ticker']
        fetch_date = slice['DOA']
        slice['DOA_Open'] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date, type = "Open")
        slice['DOA_Close'] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date)
        fetch_date = ( slice['DOA'] + timedelta(1))
        slice['Next_Day_Close'] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date)
        fetch_date = (slice['DOA'] + timedelta(3))
        slice['Third_Day_Close'] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date)
        fetch_date = (slice['DOA'] + timedelta(7))
        slice['Next_Week_Close'] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date)
        fetch_date = (slice['DOA'] + timedelta(365/12))
        slice['Next_Month_Close'] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date)
        fetch_date = (slice['DOA'] + timedelta(3*365/12))
        slice['Third_Month_Close'] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date)
        for j in range(1,item.Additional_Date_Count+1):
            print (j, "-->", item['Date'+str(j)])
            slice['Date'+str(j)] = item['Date'+str(j)]
            fetch_date = slice['Date'+str(j)]
            slice['Date' + str(j)+"_Open"] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date, type='Open')
            slice['Date' + str(j) + "_Close"] = get_quote_from_quandl(quandl_code, fetch_date=fetch_date, type='Close')
        print(slice)
        outputdf = outputdf.append(slice, ignore_index= True)
        print(item.NSE_Ticker, "Done!")
    outputdf.to_excel("Output File.xlsx", index=False)
    print ("Done making the output file")


if __name__ == '__main__':
    run()