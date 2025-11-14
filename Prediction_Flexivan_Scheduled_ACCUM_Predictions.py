import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")  # Non-interactive backend, figures won't pop up
import matplotlib.pyplot as plt
import copy
import os
import re
import sys

from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score
from xgboost import XGBClassifier
from xgboost import XGBRegressor
from datetime import datetime, timedelta
from pathlib import Path
from tqdm import tqdm
from colorama import Fore, Style, init

from sklearn.model_selection import train_test_split

from docx.shared import Inches
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import nsdecls
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import warnings
warnings.filterwarnings("ignore")

import Flexivan_Prediction_Package


# DESCRIPTION
# This script uses the building blocks developed in Prediction_Flexivan_REPO_Analysis.py and runs the scheduled task for production
# The scheduled task takes the latest samples CSV file and runs the prediction for both RETURN and PICKUP dates. It then calculates
# the dates for each prediction and updates the ACCUM file by lot for the ACCUMULATED sum of each returns and pickups

# ARGS:
#   [1] - Name of the ACCUM file
#   [2] - Name of newest CSV file with chassis samples to anaylze
#   [3] - Window size
#   [4] - Test fraction

#region Parse input variables

Input_Vars_No = 5

if len(sys.argv) < Input_Vars_No:
    print(f'{len(sys.args)} input variables provided but {Input_Vars_No-len(sys.argv)}')
    print('These are the expected parameters:')
    print('\t[1] - Name of the ACCUM file')
    print('\t[2] - Name of newest CSV file with chassis samples to anaylze')
    print('\t[3] - Window size')        
    print('\t[4] - Test fraction')        
      
    sys.exit(1)
else:
    ACCUM_File_Fullpath = sys.argv[1]
    Newest_CSV_Fullpath = sys.argv[2]
    Train_Window_Size = sys.argv[3]
    Test_Fracrion = sys.argv[4]

#endregion

#region Load the newest CSV file

NEWEST_CSV_FILE = pd.read_csv(Newest_CSV_Fullpath)

#endregion

#region Analyze for RETURNS prediction
Columns_2_Drop_From_Training = ['CHS ID', 'CTR Trip Id', 'CHS Return Dt', 'CHS Return LOC', 'CHS Pickup Date', 'CTR pick Dt', 'CTR Return Dt']
Step = int(Test_Fracrion * Train_Window_Size)
Results_DF, DATA_ORIG, DATA, Analysis_Info_DICT, y_test, y_pred, RESULTS_DETAILED_DICT = \
    Flexivan_Prediction_Package.Analyze_Data_File(Newest_CSV_Fullpath, Columns_2_Drop_From_Training, Train_Window_Size, Step, 
                                                  Error_Threshold=20, test_frac=.2, Sorting_Field='CHS Pickup Date')


#endregion


#region Load the ACCUM file into DataFrame

ACCUM_FILE_DF = pd.read_csv(ACCUM_File_Fullpath)

#endregion

#region Add the RETURNS and PICKUPS predicted by dates (fill future dates baskets)





#endregion

# Analyze for PICKUPS prediction (fill future dates baskets)

#region Save ACCUM DataFrame into a CSV file for Flexivan to consume

ACCUM_FILE_DF.to_csv(ACCUM_File_Fullpath)

#endregion