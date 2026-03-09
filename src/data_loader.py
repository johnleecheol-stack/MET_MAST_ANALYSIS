"""
MET_MAST_ANALYSIS — src/data_loader.py
Parse NRG SymphoniePRO 10-min .txt data files into a clean DataFrame.
"""

import pandas as pd
import numpy as np
import math
from pathlib import Path


# Default column names for NRG SymphoniePRO 6-anemometer / 3-vane / T+P+H setup
DEFAULT_COLUMNS = [
    'Timestamp',
    'WS100_NE_Avg','WS100_NE_SD','WS100_NE_Min','WS100_NE_Max','WS100_NE_Gust',
    'WS100_SW_Avg','WS100_SW_SD','WS100_SW_Min','WS100_SW_Max','WS100_SW_Gust',
    'WS80_NE_Avg', 'WS80_NE_SD', 'WS80_NE_Min', 'WS80_NE_Max', 'WS80_NE_Gust',
    'WS80_SW_Avg', 'WS80_SW_SD', 'WS80_SW_Min', 'WS80_SW_Max', 'WS80_SW_Gust',
    'WS60_NE_Avg', 'WS60_NE_SD', 'WS60_NE_Min', 'WS60_NE_Max', 'WS60_NE_Gust',
    'WS60_SW_Avg', 'WS60_SW_SD', 'WS60_SW_Min', 'WS60_SW_Max', 'WS60_SW_Gust',
    'Temp_Avg','Temp_SD','Temp_Min','Temp_Max',
    'BP_Avg',  'BP_SD',  'BP_Min',  'BP_Max',
    'RH_Avg',  'RH_SD',  'RH_Min',  'RH_Max',
    'WD95_Avg','WD95_SD','WD95_Gust',
    'WD75_Avg','WD75_SD','WD75_Gust',
    'WD55_Avg','WD55_SD','WD55_Gust',
    'Extra1',  'Extra2',
]
