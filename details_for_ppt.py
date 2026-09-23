# -*- coding: utf-8 -*-
"""
Created on Tuesday 6 Jan 2026

@author: Imogen
"""



import pickle as pkl
#from datetime import time
from datetime import *
from pathlib import Path
import os
import time as ti
import numpy as np
import pandas as pd
import scipy as sp
from scipy.signal import savgol_filter as sgf
import matplotlib
import matplotlib.pyplot as plt
matplotlib.use('Agg')  # disables interactive GUI backends
from clinical_data_plotting.cdp_lite_v40 import CDP as CDP4
from clinical_data_plotting.cdp_lite_v40 import sqlite_to_df, csv_to_df
import math
#from fastdtw import fastdtw
from scipy.interpolate import interp1d
from scipy.signal import butter, filtfilt
from scipy.spatial.distance import euclidean
from scipy.stats import zscore
from random import randint
from pathlib import Path
import shutil
import itertools
import random
from datetime import datetime, timedelta

#from Cleveland_data_check import *
from time_sync_utils import *
from patient_metadata import *
from wave_ICP_tools import *
#from Cleveland_data_check import crosscorr

sys.path.append(str(Path(__file__).resolve().parent.parent))
if str(Path(__file__).resolve().parent.parent) == str(Path('C:/Users/imy1/Documents/GitHub')): # User is IS
    sys.path.append(str(Path('/Users/imy1/Documents/Github/Cleveland-ICU/CRD-269')))
    sys.path.append(str(Path('/Users/imy1/Documents/Github/2022-sheep-study')))
else: # from repo structure
    sys.path.append(str(Path('/Users/imy1/Documents/Github/Cleveland-ICU/CRD-269')))
    

code_dir = str(Path('C:/Users/imy1/Documents/GitHub/Cleveland-ICU/CRD-269'))

# Set up data paths and extract patient IDs

data_fp = Path("C:/Users/imy1/Downloads/Clean_Data")
patient_metadata711 = new_patients(file_path=data_fp)
patient_metadata16 = patient_metadata(data_fp)
pids1 = list(patient_metadata16.keys())
pids2 = list(patient_metadata711.keys())
pickle_dir = 'CIP_Pickles'


nice_list = ['S1D1', 'S2D1','S2D2', 'S2D3a','S2D3b', 'S3D1', 'S3D2', 'S3D3', 'S4D1', 'S5D1','S5D2', 'S6D1','S7D1', 'S7D2', 'S8D2', 'S8D3', 'S9D1','S9D2','S9D3', 'S10D1', 'S11D1','S11D2']


pids_arr = np.array(nice_list)
pids_arr_neut = np.array(pids1+pids2) # all pids, not just nice looking data


# Set directories from which downsampled data extracted

dir_name = 'ICM_data_downsampled_and_aligned' # data to be extracted
dir_name_icm = 'ICM' # contains ICM+ data
dir_name_cyb = 'Cyban' # contains Cyban data

# Directory of RR_v3 data
RR_v3_dir = 'RR_v3_Data'

# set seed
np.random.seed(3010)

# Define which patient IDs have 'Skin' data instead of LBrain
skin_pids = ['S3D2', 'S9D3','S5D1', 'S5D2', 'S5D3', 'S6D1', 'S6D2', 'S6D3', 'S11D2']
pid_lim = len(pids_arr)


# Set ideal number of samples for each bin and 'buffer' for flatline ranges to make sure no flatlines included
num_samples = 25 # usually 25
buffer = 5 # usually 10

block_time = 13

highpass_thresh = 0.5
fps = 100

# Factor in different sensor names
times = []
total_time = 0
for pid in pids_arr:
    if pid in skin_pids:
        Sensor1 = 'Skin'
        Sensor2 = 'RBrain'
    else:
        Sensor1 = 'LBrain'
        Sensor2 = 'RBrain'

    # Load aligned data - note, these have same sample rate AND same start/end times. 
    icm_df = pd.read_csv(Path(dir_name+'/'+dir_name_icm+'/'+'icm_data_downsampled_aligned_'+pid+'.csv'))
    cyban_df = pd.read_csv(Path(dir_name+'/'+dir_name_cyb+'/'+'cyban_data_downsampled_aligned_'+pid+'.csv'))

    time = datetime.strptime(icm_df['Time'].iloc[-1],'%Y-%m-%d %H:%M:%S.%f') - datetime.strptime(icm_df['Time'].iloc[0], '%Y-%m-%d %H:%M:%S.%f')
    times.append(time.total_seconds())
    total_time = total_time + time.total_seconds()

print(times)
    
    
    
print(f'Analysed Patients have total time {total_time} seconds')