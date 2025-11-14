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


def Divide_ARR_2_Arrays_by_Range(DF, Column_Name, Ranges):
    # This function takes a DataFrame and Ranges = [0, 10, 20, ..., 100] any selected values
    # and returns a list of arrays that contains the values between those values

    DF_LIST = []
    RANGES = []

    for ii in range(len(Ranges)-1):
        DF_LIST.append(DF[(DF[Column_Name]>=Ranges[ii]) & (DF[Column_Name]<Ranges[ii+1])])
        RANGES.append([Ranges[ii], Ranges[ii+1]])

    return DF_LIST, RANGES

def train_and_test_xgboost(df, target_column, test_size, RegressionORPrediction=0, random_state=42):
    # 1. Separate features and target
    X = df.drop(columns=[target_column])
    y = df[target_column]

    # 2. Train/test split
    Split_Index = int((1-test_size) * len(df))
    X_train = df.iloc[0:Split_Index].drop(columns=[target_column])
    y_train = df.iloc[0:Split_Index][target_column]
    X_test = df.iloc[Split_Index:].drop(columns=[target_column])
    y_test = df.iloc[Split_Index:][target_column]

    # X_train, X_test, y_train, y_test = train_test_split(
    #     X, y, test_size=test_size, random_state=random_state
    # )

    # 3. Initialize XGBoost REGRESSOR model
    if RegressionORPrediction == 0:
        model = XGBClassifier(use_label_encoder=False, eval_metric='logloss')
    else:       
        model = XGBRegressor(use_label_encoder=False, eval_metric='logloss')

    # 4. Train the model
    model.fit(X_train, y_train)

    # 5. Predict and evaluate
    y_pred = model.predict(X_test)
    # accuracy = accuracy_score(y_test, y_pred)

    # print(f"Test accuracy: {accuracy:.4f}")

    return model, y_test, y_pred
    
def Display_CDF(ARR, ax=None):
    # This function displays the CDF of a selected Numpy array - what percentage of the data is found under which value

    # Sort the data
    sorted_data = np.sort(ARR)

    # Compute CDF values
    CDF = (100 *np.arange(1, len(sorted_data) + 1) / len(sorted_data))[::-1]

    # Plot the CDF
    if ax is None:
        fig, ax = plt.subplots()
        
    ax.plot(sorted_data, CDF, linewidth=2)
    plt.grid(True, linestyle='--', alpha=0.6)
    plt.show()

    return CDF, ax
    
def enumerate_columns(df, column_name):
    """
    Replace values in each column with integer codes representing
    the unique values in that column.
    """
    unique_vals = {v: i for i, v in enumerate(df[column_name].unique())}
    df[column_name] = df[column_name].map(unique_vals)
    
    return df

def sliding_xgb_window_eval(df, target_col, window_size, step, test_frac=0.2,
                            error_threshold=20.0, xgb_params=None, random_state=42,
                            min_test_samples=2, show_progress=False, Classifier_or_Regressor=0):
    """
    Slide a fixed-size window over df, train XGBRegressor on the first (1-test_frac)
    fraction and test on the last test_frac fraction. For each window compute the
    percentage of test rows whose percentage error <= error_threshold.

    Returns a DataFrame with columns:
    window_idx, start, end, n_test, pct_under_threshold
    """
    import math
    from tqdm import tqdm
    from xgboost import XGBRegressor, XGBClassifier

    if xgb_params is None:
        xgb_params = {'n_estimators': 100, 'random_state': random_state, 'verbosity': 0}

    results = []
    n = len(df)
    if window_size > n:
        window_size = n
        print("window_size larger than dataframe length. Setting one windw")

    indices = range(0, n - window_size + 1, step)
    iterator = tqdm(indices) if show_progress else indices
    wi = 0
    RESULTS_DETAILED_DF = {}            # Holds the detailed results for each window for results CSV purposes. Keys are the start index,
                                        # and the content is a list of [[y_test], [y_pred]]

    print('Training and predicting for sliding windows...')
    for start in tqdm(iterator):
        try:
            end = start + window_size  # exclusive
            window = df.iloc[start:end].copy()

            # drop rows missing target
            window = window.dropna(subset=[target_col])
            if len(window) < 2:
                continue

            # prepare features (dummies on full window to keep alignment)
            X_all = pd.get_dummies(window.drop(columns=[target_col]), drop_first=True)
            y_all = window[target_col].astype(float).values
            # y_all = window[target_col].astype(int).values

            split_idx = int(math.floor((1.0 - test_frac) * len(window)))
            # ensure at least min_test_samples in test
            if len(window) - split_idx < min_test_samples or split_idx < 1:
                # skip window too small
                continue

            X_train = X_all.iloc[:split_idx, :].values
            y_train = y_all[:split_idx]
            X_test = X_all.iloc[split_idx:, :].values
            y_test = y_all[split_idx:]

            # train
            if Classifier_or_Regressor:
                xgb_params = {'n_estimators': 100, 'random_state': random_state, 'verbosity': 0, 'num_class': len(set(y_train))}
                model = XGBClassifier(**xgb_params)
                y_test = [int(x) for x in y_test]
                y_train = [int(x) for x in y_train]
            else:
                model = XGBRegressor(**xgb_params)

            # try:
            #     model.fit(X_train, y_train)   
            # except Exception as e:
            #     pass
            model.fit(X_train, np.array(y_train))
            y_pred = model.predict(X_test)

            # model, y_test, y_pred = train_and_test_xgboost(window, target_col, test_size, RegressionORPrediction=1, random_state=42)

            # compute percent error safely:
            abs_diff = np.abs(y_test - y_pred)
            # If true value is zero, define percent error as 0 if prediction equals, else 100.
            pct_err = np.where(
                np.isclose(y_test, 0.0),
                np.where(np.isclose(abs_diff, 0.0), 0.0, 100.0),
                100.0 * abs_diff / np.abs(y_test)
            )

            RESULTS_DETAILED_DF[start] = [y_test, y_pred, start + np.arange(split_idx, split_idx+len(y_test))]

            pct_under = 100.0 * float(np.sum(pct_err <= error_threshold)) / len(pct_err)
        except:
            pass

        results.append({
            'window_idx': wi,
            'start': start,
            'end': end,
            'n_test': len(pct_err),
            'pct_under_threshold': pct_under
        })

        wi += 1

    print('DONE.')

    return pd.DataFrame(results).sort_values('window_idx').reset_index(drop=True), y_test, y_pred, RESULTS_DETAILED_DF

def Analyze_RAW_Windows_Results(DATA, RESULTS_DETAILED_DICT):
    # This function takes the RESULTS_DETAILED_DF yielded by the sliding_xgb_window_eval function and contains the the results y_test-y_pred
    # and generates a DataFrame that contains pickup date (retrieved by iloc index from DATA) and the results of the prediction vs. GT.
    # Second part of the analysis, takes the repeating indexes (table index) and chooses the best prediction result
    # RESULTS_DETAILED_DICT = [y_test, y_pred, Original_Indexes]            (Original_Indexes are in the full DATA DF cleaned from CSV file)

    Window_Size = None
    RESULTS_DETAILED_DF = None
    DATA_COLUMNS = list(DATA.columns)

    # Collect all results from all windows
    for key in RESULTS_DETAILED_DICT.keys():
        if Window_Size is None:
            Window_Size = len(RESULTS_DETAILED_DICT[key][0])

        Col_index = DATA_COLUMNS.index('CHS Pickup Date')

        Pickup_Dates = DATA.iloc[RESULTS_DETAILED_DICT[key][2], Col_index]

        RESULTS_DETAILED_DF_TEMP = pd.DataFrame({
            'Pickup Date': np.array(Pickup_Dates).ravel(),
            'Return Date Diff': np.array(RESULTS_DETAILED_DICT[key][0]).ravel(),
            'Return Date Diff Predicted': np.array(RESULTS_DETAILED_DICT[key][1]).ravel(),
            'Diff % (ABS)': 100*abs(np.array(RESULTS_DETAILED_DICT[key][1])-np.array(RESULTS_DETAILED_DICT[key][0]))/np.array(RESULTS_DETAILED_DICT[key][0])
        })

        if RESULTS_DETAILED_DF is None:
            RESULTS_DETAILED_DF = copy.deepcopy(RESULTS_DETAILED_DF_TEMP)
        else:
            RESULTS_DETAILED_DF = pd.concat([RESULTS_DETAILED_DF, RESULTS_DETAILED_DF_TEMP])

    RESULTS_DETAILED_DF_REFINED = copy.deepcopy(RESULTS_DETAILED_DF)

    return RESULTS_DETAILED_DF_REFINED
        
def Analyze_Data_File(Fullpath, Columns_2_Drop_From_Training, window_size, step, Error_Threshold=20, test_frac=.2, Sorting_Field='CHS Pickup Date'):
    # This function analyzes data in CSV found in Fullpath. It uses the sliding window XGBoost model (test and train over a portion of the data)
    # The process is repeated for sliding windows of size (window_size with step of step samples)

    Folder = os.path.dirname(Fullpath)
    Filename = os.path.basename(Fullpath)

    Analysis_Info_DICT = {}

    #region Load data from file
    DATA = pd.read_csv(Folder + '/' + Filename)

    DATA['CHS Pickup Date'] = pd.to_datetime(DATA['CHS Pickup Date'], errors='coerce')
    DATA['CHS Return Dt'] = pd.to_datetime(DATA['CHS Return Dt'], errors='coerce')
    DATA_ORIG = copy.deepcopy(DATA)

    Analysis_Info_DICT['Total_Lines'] = len(DATA_ORIG)

    #endregion

    #region Data cleanup

    DATA_LEN = len(DATA)

    for col in DATA.select_dtypes(include=['object']).columns:
        DATA[col] = DATA[col].map(lambda x: x.strip() if isinstance(x, str) else x)

    # Normalize common missing markers and empty strings to NaN
    DATA.replace(['', 'NA', 'N/A', 'na', 'n/a'], np.nan, inplace=True)

    # Drop any row that contains at least one NaN
    DATA.dropna(axis=0, how='any', inplace=True)

    print(f'Sorting by field {Sorting_Field}...', end='')
    DATA = DATA.sort_values(by=Sorting_Field)
    print('DONE.')

    DATA.reset_index(drop=True, inplace=True)

    print(f"Rows after cleaning: " + Fore.YELLOW + f'{len(DATA)}' + Fore.RESET + ' (' + Fore.GREEN + f'{int(100*len(DATA)/DATA_LEN)}' + Fore.RESET + ')% remained after cleaning')
    print(f'Sorted by ' + Fore.YELLOW + f'{Sorting_Field}' + Fore.RESET + ' field')

    Analysis_Info_DICT['Rows_After_Cleaning'] = len(DATA)

    #endregion

    #region Calculating differences and addting new columns
    
    Units = 'Days'
    Diff_Col_Name = f'Pickup_Return_Time_Diff_{Units}'
    Analysis_Info_DICT['Time_Diff_Units'] = Units

    if Units == 'Hours':
        DATA[Diff_Col_Name] = (DATA['CHS Return Dt'] - DATA['CHS Pickup Date']).dt.total_seconds() / 3600
        
    elif Units == 'Days':
        DATA[Diff_Col_Name] = (DATA['CHS Return Dt'] - DATA['CHS Pickup Date']).dt.total_seconds() / (3600*24)

    #endregion

    #region Enumerate data

    Enumerated_Columns_LIST = ['CHS Pickup Loc', 'CHS Return Loc', 'CHS pickup MCO', 'CTR Trip MCO', 'O Customer', 'Customer', 'DC Loc', 'CTR Pickup Term', 'CTR Return Term', 
                           'pgkey', 'CTR Trip Loc Type Pattern', 'CTR Trip Pattern']
    
    for column in DATA.columns:
        if column in Enumerated_Columns_LIST:
            DATA = enumerate_columns(DATA, column)

    for column in Columns_2_Drop_From_Training:
        if column in DATA.columns:
            DATA.drop(columns=[column], inplace=True)
    try:
        DATA['CHS Return Loc'] = DATA['CHS Return Loc'].astype(int)
    except:
        pass

    #endregion

    #region Analyzing data file    

    # Predicting return time diff
    step=int(test_frac * window_size)      # Force so there will be no overlapping results (more than one prediction per sample)
    Results_DF, y_test, y_pred, RESULTS_DETAILED_DICT = sliding_xgb_window_eval(DATA, Diff_Col_Name, window_size, step, test_frac,
                                                                              Error_Threshold, xgb_params=None, random_state=42,
                                                                              min_test_samples=2, show_progress=False, Classifier_or_Regressor=0)
    # Predicting return LOT
    __, y_test_Return_LOT, y_pred_Return_LOT, RESULTS_DETAILED_DICT_LOT = sliding_xgb_window_eval(DATA, 'CHS Return Loc', window_size, step, test_frac,
                                                                              Error_Threshold, xgb_params=None, random_state=42,
                                                                              min_test_samples=2, show_progress=False, Classifier_or_Regressor=0)
    y_pred_Return_LOT = [int(x) for x in y_pred_Return_LOT]

    # # Analyze RESULTS_DETAILED_DF - raw results from all windows of the XGBoost. Takes the best accuracy for each pickup date index
    # RESULTS_DETAILED_REFINED = Analyze_RAW_Windows_Results(DATA_ORIG, RESULTS_DETAILED_DICT)
    
    Analysis_Info_DICT['Model_Window_Size'] = window_size
    Analysis_Info_DICT['Model_Window_Step'] = step
    Analysis_Info_DICT['Model_Erro_THR'] = Error_Threshold

    #endregion

    return Results_DF, DATA_ORIG, DATA, Analysis_Info_DICT, y_test, y_pred, RESULTS_DETAILED_DICT, RESULTS_DETAILED_DICT_LOT

def extract_datetimes_from_filenames(filenames):
    """
    Extracts datetime objects from filenames of the form 'Latest_Test_<DATE>'.
    Supports formats like YYYY-MM-DD, YYYYMMDD, or YYYY_MM_DD.
    """
    datetimes = []

    for name in filenames:
        # Try to find the date pattern
        match = re.search(r'(\d{4}[-_]\d{2}[-_]\d{2}|\d{8})', name)
        if match:
            date_str = match.group(1)
            
            # Normalize formats like YYYY_MM_DD → YYYY-MM-DD
            date_str = date_str.replace("_", "-")
            
            # Try multiple possible formats
            for fmt in ("%Y-%m-%d", "%Y%m%d"):
                try:
                    dt = datetime.strptime(date_str, fmt)
                    datetimes.append(dt)
                    break
                except ValueError:
                    continue

    return datetimes

def set_thick_outside_borders(table, size=12, color="000000"):
    """
    Apply thick borders to the outside of a table.
    - size: border thickness (12 ~ 2pt)
    - color: hex color code (default black)
    """
    tbl = table._tbl

    # Ensure tblPr exists
    if tbl.tblPr is None:
        tbl.tblPr = OxmlElement('w:tblPr')

    tblBorders = OxmlElement('w:tblBorders')

    for border_name in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), str(size))
        border.set(qn('w:color'), color)
        tblBorders.append(border)

    # Disable inside borders
    for inside in ['insideH', 'insideV']:
        border = OxmlElement(f'w:{inside}')
        border.set(qn('w:val'), 'nil')
        tblBorders.append(border)

    # Append borders to table properties
    tbl.tblPr.append(tblBorders)

def Add_DICT_2_Table(doc, Results_DICT, HEADERS):
    # This function takes a dictionary that contains one data item in each key and adds it to a table in doc.
    # The function returns the DOCX document object with the added table

    # Add table with header row
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'  # You can also try 'Light Grid Accent 1', etc.

    # Add header cells
    hdr_cells = table.rows[0].cells
    for ii, header in enumerate(HEADERS):
        # hdr_cells[ii].text = header
        para = hdr_cells[ii].paragraphs[0]
        run = para.add_run(header)
        run.bold = True
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run.font.size = Pt(12)
   
    for key in Results_DICT.keys():
        row_cells = table.add_row().cells
        row_cells[0].text = str(key)
        row_cells[1].text = str(np.round(Results_DICT[key], 2)) + '%'

    set_thick_outside_borders(table)

    return doc

def Generate_Return_Dates_from_DIFF(Pickup_Dates, DIFFs):
    # Genetraes an array of return dates (datetime format) from datetimes in Pickup_Dates using float values in DIFFs.
    # This is used for the ACCUM file generation of the final result

    Return_Dates = []

    for ii, pickup_date in enumerate(Pickup_Dates):
        return_date = pickup_date + timedelta(days=DIFFs[ii])
        Return_Dates.append(return_date)

    return Return_Dates

def Add_File_Results_2_ACCUM_Results_DF(DATA_ORIG, Results_File_ACCUM_DF, Detailed_Pred_Results_DF):
    # DATA_ORIG - the original cleaned version of the data file. This is used to get information about a sample that
    #             does not exist in the Detailed_Pred_Results_DF table (LOT and PU date)
    # Results_File_ACCUM_DF - this is the daily report to maintain with ACCUM figure for returns and pickups
    # Detailed_Pred_Results_DF = [[y_pred], [y_test]], [Indexes_in_ORIG], [Predicted_LOT]]

    if Results_File_ACCUM_DF is None:
        Results_File_ACCUM_DF = pd.DataFrame(columns=['LOT', 'Date', 'Predicted_Pickups', 'Predicted_Returns'])
    
    #region Adding the RETURNS PREDICTION indormation into the ACCUM report

    for jj, key in enumerate(Detailed_Pred_Results_DF.keys()):
        Epoch_DATA = Detailed_Pred_Results_DF[key]      # [[y_pred], [y_test]], [Indexes_in_ORIG]]

        y_test = Epoch_DATA[0]
        y_pred = Epoch_DATA[1]
        Indexes_ORIG = Epoch_DATA[2]
        LOT = 0             # TODO: complete the LOT prediction model
        # LOTS = Epoch_DATA[3]
                                    # These are all aligned in indexes
    
        PickUp_Dates = DATA_ORIG.iloc[Indexes_ORIG]['CHS Pickup Date']
        Display_Factor = 100

        for ii, DATE in enumerate(PickUp_Dates):
            if ii%Display_Factor==0:
                print(f'Epoch No.{jj+1} out of {len(Detailed_Pred_Results_DF.keys())} - {ii} out of {len(PickUp_Dates)} Dates completed', end='\r')

            # LOT = LOTS[ii]
            try:
                DATE = datetime.strptime(DATE, "%Y-%m-%d %H:%M:%S")
            except:
                pass

            Future_Date = DATE + timedelta(days=float(y_pred[ii]))
            Future_Date = DATE.date()
            Future_Date = DATE.strftime("%Y-%m-%d")

            if len(Results_File_ACCUM_DF)==0:
                Results_File_ACCUM_DF.loc[len(Results_File_ACCUM_DF)] = [LOT, Future_Date, 0, 1]
                Results_File_ACCUM_DF.set_index('Date', inplace=True)
            else:
                try:
                    Results_File_ACCUM_DF.loc[Future_Date, 'Predicted_Returns'] += 1
                    pass
                except:
                    Results_File_ACCUM_DF.loc[Future_Date] = [LOT, 0, 1]
    #endregion

    return Results_File_ACCUM_DF

#endregion