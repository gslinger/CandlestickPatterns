# Candlestick Pattern Detection Based on OHLC Data

This module scans OHLC data and tests for CandleStick Patterns based on shape data. 

Currently the pattern scan only looks back at 3 lags, this may be improved later to detect longer trends

This is largely an educational project, so feel free to offer advice, criticism or contributions

The formulas i use to detect are based on the OHLC criteria and my main source is : https://www.candlescanner.com/patterns-dictionary/

I've been comparing my results to that of some trading software's candlescans and found my results to be the same for the most-part. 

However, any suggestions on improvements or optimizations, please let me know.

# Current Pattern List:
 - Doji
 - Bearish Engulfing
 - Dark Cloud Cover
 - Three Outside Down
 - Evening Star Doji
 - Bearish Harami

# To-Do 
- Add more candle stick patterns
- Make more user friendly
- Improve algorithms
- Add some customization/sensitivity options
- (maybe) plotting
- (maybe) testing
- (maybe) add functions to detect candle color, shape, trend etc.

# Usage 
- Add module to Excel Project
- You will need to set wb and ws variables to where your data is
- Change constants ColNo and RowNo to state where you want the output
- Change the ranges for O/H/L/C data in the CandleScan subroutine
- Run CandleScan, for now it will simply print in the assigned column if it fits criteria
- If two candlestick patterns are detected at the same date, for now it will just write them both


