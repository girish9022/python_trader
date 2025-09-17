import logging
import yaml
from NorenRestApiPy.NorenApi import NorenApi
import pandas as pd
from datetime import datetime

# This script will fetch historical data, calculate SMAs, and save to Excel.

class ShoonyaApp(NorenApi):
    def __init__(self):
        NorenApi.__init__(self, host='https://api.shoonya.com/NorenWClientTP/',
                        websocket='wss://api.shoonya.com/NorenWSTP/',
                        eodhost='https://api.shoonya.com/chartApi/getdata/')

    def login(self):
        # Load credentials from cred.yml
        with open('cred.yml') as f:
            cred = yaml.load(f, Loader=yaml.FullLoader)

        ret = super().login(userid=cred['user'], password=cred['pwd'], twoFA=cred['factor2'],
                         vendor_code=cred['vc'], api_secret=cred['apikey'], imei=cred['imei'])

        if ret is not None and ret.get('stat') == 'Ok':
            print("Login successful")
        else:
            print(f"Login failed: {ret}")
            exit()

    def get_time_series(self, exchange, token, starttime, endtime, interval):
        ret = self.get_time_price_series(exchange=exchange, token=token, starttime=starttime, endtime=endtime, interval=interval)

        if ret:
            df = pd.DataFrame(ret)
            df.rename(columns={'into': 'open', 'inth': 'high', 'intl': 'low', 'intc': 'close', 'intv': 'volume'}, inplace=True)
            df['time'] = pd.to_datetime(df['time'], format='%d-%m-%Y %H:%M:%S')
            numeric_cols = ['open', 'high', 'low', 'close', 'volume']
            df[numeric_cols] = df[numeric_cols].apply(pd.to_numeric, errors='coerce')
            return df
        return pd.DataFrame()

def aggregate_to_interval(df, interval_minutes=15, tz='Asia/Kolkata'):
    if df.empty:
        return df

    df = df.set_index('time')
    try:
        df.index = df.index.tz_localize(tz)
    except TypeError:
        # already localized
        pass


    agg_df = df.resample(f'{interval_minutes}T', label='right', closed='right').agg({
        'open': 'first',
        'high': 'max',
        'low': 'min',
        'close': 'last',
        'volume': 'sum'
    })

    agg_df.dropna(subset=['close'], inplace=True)

    return agg_df

def calculate_smas(df, sma_windows=(5, 21, 50, 100, 200)):
    if df is None or df.empty:
        return None

    # Aggregate to 15-minute candles
    agg_df = aggregate_to_interval(df, interval_minutes=15)

    # Calculate SMAs
    for w in sma_windows:
        agg_df[f'sma{w}'] = agg_df['close'].rolling(window=w, min_periods=1).mean()

    return agg_df

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO)

    # --- 1. Login ---
    api = ShoonyaApp()
    api.login()

    # --- 2. Define parameters ---
    exchange = 'NSE'
    symbol = 'RELIANCE-EQ' # Example symbol, change as needed
    output_file = 'sma_data.xlsx'

    # --- 3. Get token for the symbol ---
    logging.info(f"Searching for token for {symbol}")
    ret = api.searchscrip(exchange=exchange, searchtext=symbol)
    if not ret or not ret.get('values'):
        logging.error(f"Could not find token for {symbol}")
        exit()
    token = ret['values'][0]['token']
    logging.info(f"Found token: {token}")

    # --- 4. Fetch historical data ---
    endtime = datetime.now()
    starttime = endtime - pd.Timedelta(days=15)

    logging.info(f"Fetching 1-minute data from {starttime} to {endtime}")
    raw_data = api.get_time_series(exchange=exchange,
                                   token=token,
                                   starttime=int(starttime.timestamp()),
                                   endtime=int(endtime.timestamp()),
                                   interval=1)

    if raw_data.empty:
        logging.error("No historical data received.")
        exit()
    logging.info(f"Received {len(raw_data)} rows of 1-minute data.")

    # --- 5. Calculate SMAs on 15-min candles ---
    logging.info("Calculating SMAs on 15-minute aggregated data.")
    sma_data_15min = calculate_smas(raw_data)

    if sma_data_15min is None or sma_data_15min.empty:
        logging.error("Could not calculate SMAs.")
        exit()

    # --- 6. Save the 15-minute SMA data to Excel ---
    final_df_to_save = sma_data_15min.reset_index()

    final_df_to_save.to_excel(output_file, index=False)
    logging.info(f"15-minute SMA data saved to {output_file}")
    print("\n--- Script Finished ---")
    print(f"Latest SMA data:")
    print(final_df_to_save.tail())
