from gooey import Gooey, GooeyParser
import MetaTrader5 as mt5
from datetime import datetime, date
import pytz
import pandas as pd
import numpy as np
import signal

TIMEFRAMES = {"1m":mt5.TIMEFRAME_M1,
              "5m":mt5.TIMEFRAME_M5,
              "10m":mt5.TIMEFRAME_M10,
              "15m":mt5.TIMEFRAME_M15,
              "30m":mt5.TIMEFRAME_M30,
              "1h":mt5.TIMEFRAME_H1,
              "6h":mt5.TIMEFRAME_H6,
              "8h":mt5.TIMEFRAME_H8,
              "12h":mt5.TIMEFRAME_H12,
              "1d":mt5.TIMEFRAME_D1,
              "1w":mt5.TIMEFRAME_W1,
              "1M":mt5.TIMEFRAME_MN1
              }

TZ = pytz.timezone("Etc/UTC")

def get_ohlcv(symbol:str, timeframe:int, date_from:str, date_to:str, round_to=4) -> pd.DataFrame:
    global TZ
    # Note: In Gooey, the Dates are given as strings with format (YYYY-MM-DD).

    # Get Data from MT5
    temp_start = datetime.strptime(date_from, "%Y-%m-%d")
    start_date = datetime(temp_start.year, temp_start.month, temp_start.day, tzinfo=TZ)
    temp_end = datetime.strptime(date_to, "%Y-%m-%d")
    end_date = datetime(temp_end.year, temp_end.month, temp_end.day, tzinfo=TZ)
    raw_prices = mt5.copy_rates_range(symbol, timeframe, start_date, end_date) # Using UTC Timezone

    # Note: When it returns data, the time is Timestamp

    # Clean Data
    df = pd.DataFrame(raw_prices)
    df['time'] = pd.to_datetime(df['time'], unit='s')
    # df['time'] = pd.to_datetime(df['time'])
    df.rename(columns={"time": "Datetime", "open": "Open", "high":"High", "low":"Low", "close":"Close", "tick_volume":"Volume"}, inplace=True)
    df = df[["Datetime", "Open", "High", "Low", "Close", "Volume",]] # Only Select DOHLCV
    df.set_index("Datetime", inplace=True)
    df = df.round(decimals=round_to)
    # 1m is not an error, user need to set unlimited 'Max. bars in chart' settings in their MT5 Platform.

    # Return Data
    return df

def get_tick(symbol:str, date_from:str, date_to:str, round_to=4) -> pd.DataFrame:
    # Get Data
    temp_start = datetime.strptime(date_from, "%Y-%m-%d")
    start_date = datetime(temp_start.year, temp_start.month, temp_start.day, tzinfo=TZ)
    temp_end = datetime.strptime(date_to, "%Y-%m-%d")
    end_date = datetime(temp_end.year, temp_end.month, temp_end.day, tzinfo=TZ)
    raw_prices = mt5.copy_ticks_range(symbol, start_date, end_date, mt5.COPY_TICKS_ALL)

    # Clean Data
    df = pd.DataFrame(raw_prices)
    df['time'] = pd.to_datetime(df['time'], unit='s')
    df.rename(columns={"time": "Datetime", "bid": "Bid", "ask": "Ask"}, inplace=True) # Rename Columns
    df = df[["Datetime", "Bid", "Ask"]]
    df.set_index("Datetime", inplace=True)
    df = df.round(decimals=round_to) # Round to N Deminal Places

    # Return Data
    return df

@Gooey(program_name="MT5 Data Downloader", program_description="Utility software for downloading MetaTrader 5 Tick or OHLCV Data",
       navigation="SIDEBAR", tabbed_groups=True,
       default_size=(1000, 700),
       menu=[{"name":"Help",
              "items":[{"type":"Link", "menuTitle":"Setting Unlimited Max Bars", "url":"https://www.metatrader5.com/en/terminal/help/startworking/settings#max_bars"},
                       {"type":"MessageDialog", "menuTitle":"Timezone Info", "message":"The Datetime is Standardized to UTC.", "caption":"Timezone Info"}
                       ]
              }
             ],
       shutdown_signal=signal.CTRL_C_EVENT
       )
def parse_args():
    parser = GooeyParser()
    # subparsers = parser.add_subparsers(required=False)

    # main_subparser = subparsers.add_parser("Data Downloader")
    # subparser1 = subparsers.add_parser("Action2")

    login_options_group = parser.add_argument_group("Login Details")
    login_options_group.add_argument("account", metavar="Account Number", widget="TextField", help="MT5 Account Number") # Dont forget to convert to int
    login_options_group.add_argument("password", metavar="Password", widget="PasswordField", help="MT5 Account Password")
    login_options_group.add_argument("server", metavar="Server", widget="TextField", help="e.g. mt5-demo01.pepperstone.com")

    symbol_options_group = parser.add_argument_group("Symbol Options")
    symbol_options_group.add_argument("symbol", metavar="Symbol", widget="TextField", help="Valid Symbol from your MT5.")
    symbol_options_group.add_argument("start_date", metavar="Start Date", widget="DateChooser", help="Select the Start Date.")
    symbol_options_group.add_argument("end_date", metavar="End Date", widget="DateChooser", help="Select the End Date.")

    # symbol_options_group.add_argument("ohlcv_or_tick", metavar="OHLCV/Tick", widget="Dropdown", choices=["OHLCV", "Tick"], help="Type of Data to Download.")
    ohlcv_or_tick = symbol_options_group.add_mutually_exclusive_group("OHLCV or Tick", gooey_options=dict(title="Choose the Quote Type", required=True))
    ohlcv_or_tick.add_argument("--select_ohlcv", metavar="OHLCV", action="store_true", help="Select OHLCV Data")
    ohlcv_or_tick.add_argument("--select_tick", metavar="Tick", action="store_true", help="Select Tick Data")

    symbol_options_group.add_argument("export_directory", metavar="Directory", widget="DirChooser", help="Select Directory for Export.")
    symbol_options_group.add_argument("file_extension", metavar="File Extension", widget="Dropdown",
                                      help="Select the File Extension for export (.csv/.xlsx).", choices=[".csv", ".xlsx"])
    symbol_options_group.add_argument("--round_to", metavar="Round", widget="IntegerField", type=int, default=4,
                                      help="Round to decimal places")
    symbol_options_group.add_argument("--timeframe", metavar="Timeframe", widget="Dropdown", default="1d",
                                      help="Select the Timeframe for OHLCV Date. Must be selected for OHLCV Data.",
                                      choices=[t for t in TIMEFRAMES.keys()])

    # ohlcv_options = symbol_options_group.add_mutually_exclusive_group("OHLCV Options")
    # ohlcv_options.add_argument("--timeframe", metavar="Timeframe", widget="Dropdown", help="Select the Timeframe for OHLCV Date",
    #                            choices=[t for t in TIMEFRAMES.keys()])

    return parser.parse_args()


def main(args):
    mt5.initialize()
    authorized = mt5.login(int(args.account), password=args.password, server=args.server)
    if authorized:
        print("Connected: Connecting to MT5 Client")
        print(f"MT5 Terminal Version: {mt5.version()}")
    else:
        print("Failed to connect at account #{}, error code: {}".format(account, mt5.last_error()))
        raise ConnectionError("Cannot Connect to MT5 Terminal.")

    # if args.ohlcv_or_tick == "OHLCV":
    if args.select_ohlcv:
        if args.timeframe not in TIMEFRAMES.keys(): raise ValueError("Please select a timeframe for OHLCV Data")
        data = get_ohlcv(args.symbol, TIMEFRAMES[args.timeframe], args.start_date, args.end_date, round_to=args.round_to)
        filepath = fr'{args.export_directory}\{args.symbol}_{args.timeframe}_{args.file_extension}'
        if args.file_extension == ".csv":
            data.to_csv(filepath)
        if args.file_extension == ".xlsx":
            data.to_excel(filepath)

    # if args.ohlcv_or_tick == "Tick":
    if args.select_tick:
        data = get_tick(args.symbol, args.start_date, args.end_date)
        filepath = fr'{args.export_directory}\{args.symbol}_Tick{args.file_extension}'
        if args.file_extension == ".csv":
            data.to_csv(filepath)
        if args.file_extension == ".xlsx":
            data.to_excel(filepath)
    print(args)

if __name__ == "__main__":
    try:
        main(parse_args())
    except:
        mt5.shutdown()