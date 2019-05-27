'''
Compares the files generated by macros
with the improvement plans submitted by users
'''
#Import statements starts
import os
import configparser
import pandas as pd
import numpy as np
import win32com.client as win32
#Import statements ends

#Function declaration Phase Starts
def diff_pd(df1, df2, plan, imp_plan_section):
    """Identify differences between two pandas DataFrames"""
    assert (df1.columns == df2.columns).all(), \
        print(
            "DataFrame column names are different for the section:" +
            str(imp_plan_section) + " in the file:" + str(plan)
            )
    if any(df1.dtypes != df2.dtypes):
        print(
            "Data Types are different for the section:" +
            str(imp_plan_section) + " in the file:" +
            str(plan) + ", trying to convert"
            )
        df2 = df2.astype(df1.dtypes)
    if df1.equals(df2):
        return pd.DataFrame(
            {
                'plan': plan, 'section': imp_plan_section,
                'Result': 'Match', 'from': '', 'to': ''
                },
            index=['First']
            )
    if len(df1.columns) != len(df2.columns):
        return pd.DataFrame(
            {
                'plan': plan, 'section': imp_plan_section,
                'Result': 'Missmatch column count',
                'from': str(len(df1.columns))+' rows', 'to': str(len(df2.columns))+' rows'
                },
            index=['First']
            )
    if len(df1) != len(df2):
        #df1["identifierOld"] = "Yes"
        #df2["identifierNew"] = "Yes"
        #dfMerge = pd.merge(
        # df1,df2,how="outer",left_on=["Tooling Info"],right_on=["Tooling Info"]
        # )
        #dfMerge['MismatchRows'] = (
        # df1["identifierOld"]).astype(str) + "-" + (df2["identifierNew"]
        # ).astype(str)
        #dfMerge = dfMerge[~dfMerge['MismatchRows'].str.contains("Yes-Yes", na=False)]
        return pd.DataFrame(
            {
                'plan': NEWFILENAME, 'section': imp_plan_section,
                'Result': 'Missmatch row count',
                'from': str(len(df1))+' rows', 'to': str(len(df2))+' rows'
                },
            index=['First']
            )
    diff_mask = (df1 != df2) & ~(df1.isnull() & df2.isnull())
    ne_stacked = diff_mask.stack()
    changed = ne_stacked[ne_stacked]
    changed.index.names = ['id', 'col']
    difference_locations = np.where(diff_mask)
    changed_from = df1.values[difference_locations]
    changed_to = df2.values[difference_locations]
    return pd.DataFrame(
        {
            'plan': NEWFILENAME, 'section': imp_plan_section,
            'Result': 'Missmatch in data',
            'from': changed_from, 'to': changed_to
            },
        index=changed.index
        )

def check_sheet_protection(file_path, plan, sheet):
    '''Sheet protection check'''
    excel_file = win32.gencache.EnsureDispatch('Excel.Application')
    work_book = excel_file.Workbooks.Open(file_path)
    excel_file.Visible = True
    work_sheet = work_book.Worksheets(sheet)
    try:
        work_sheet.Range("A1").Value = "Cell A1"
        return pd.DataFrame(
            {
                'plan': plan, 'sheet': sheet,
                'Result': 'Sheet Not Protected'
            },
            index=['FIRST']
        )
    except:
        return pd.DataFrame(
            {
                'plan': plan, 'sheet': sheet,
                'Result': 'Sheet Protected'
            },
            index=['FIRST']
        )
    finally:
        excel_file.DisplayAlerts = False
        work_book.Close(False)
        excel_file.Application.Quit()
#Function declaration Phase Ends

#Preprocessing Preparation Phase Starts
CONFIG = configparser.ConfigParser()
CONFIG.read('ExcelComparatorConfigFile.ini')

IMPPLANOLDFILEPATH = CONFIG['ConnectionString OldFilePath']['sourcefilepath']
IMPPLANNEWFILEPATH = CONFIG['ConnectionString NewFilePath']['sourcefilepath']

ALLOUTPUTFILES = list(filter(lambda x: x.endswith('.xlsx'), os.listdir(IMPPLANNEWFILEPATH)))
CONFIGSHEETREADER = list(filter(lambda x: x.startswith('Readsheet'), CONFIG.sections()))

#NEWFILENAME = ALLOUTPUTFILES[randint(0,47)]
NEWFILENAME = ALLOUTPUTFILES[28]
OLDFILENAME = NEWFILENAME[:-4] + "xlsm"  #Once migration is done this line will change

OLDFILEPATH = IMPPLANOLDFILEPATH + OLDFILENAME
NEWFILEPATH = IMPPLANNEWFILEPATH + NEWFILENAME
COMPARERESULTAPPENDED = pd.DataFrame()
SHEETPROTECTIONRESULTAPPENDED = pd.DataFrame()
#Preprocessing Preparation Phase Ends

#Data Reading & Testing Phase Starts
for configSheetReaderItem in CONFIGSHEETREADER:
    oldData = pd.read_excel(
        OLDFILEPATH, sheet_name=CONFIG[configSheetReaderItem]['SheetName'],
        skiprows=int(CONFIG[configSheetReaderItem]['skiprows']),
        nrows=int(CONFIG[configSheetReaderItem]['nrows']),
        usecols=CONFIG[configSheetReaderItem]['parse_cols']
        )
    if configSheetReaderItem in ('Readsheet ProjectList2019', 'Readsheet ProjectList2020'):
        newData = pd.read_excel(
            NEWFILEPATH, sheet_name=CONFIG[configSheetReaderItem]['SheetName'],
            skiprows=int(CONFIG[configSheetReaderItem]['skiprows'])+4,
            nrows=int(CONFIG[configSheetReaderItem]['nrows']),
            usecols=CONFIG[configSheetReaderItem]['parse_cols']
            )
    else:
        newData = pd.read_excel(
            NEWFILEPATH, sheet_name=CONFIG[configSheetReaderItem]['SheetName'],
            skiprows=int(CONFIG[configSheetReaderItem]['skiprows']),
            nrows=int(CONFIG[configSheetReaderItem]['nrows']),
            usecols=CONFIG[configSheetReaderItem]['parse_cols']
            )

    dfOldData = pd.DataFrame(oldData)
    dfOldData = dfOldData[dfOldData[CONFIG[configSheetReaderItem]['nullCheck']].notnull()]
    #Section to be removed post migration
    if configSheetReaderItem in ('Readsheet ProjectList2019', 'Readsheet ProjectList2020'):
        dfOldData = dfOldData[~dfOldData[CONFIG[configSheetReaderItem]['nullCheck']].str.contains(
            '<Project', na=False)]
        dfOldData = dfOldData[~dfOldData[CONFIG[configSheetReaderItem]['RankCheck']].apply(
            lambda x: isinstance(x, (int, np.int64, float, np.float64)))]
        dfOldData[CONFIG[configSheetReaderItem]['RankCheck']][
            dfOldData[CONFIG[configSheetReaderItem]['nullCheck']] == "Other"
            ] = "AAAA"
        dfOldData = dfOldData.sort_values(
            [
                CONFIG[configSheetReaderItem]['RankCheck'],
                CONFIG[configSheetReaderItem]['nullCheck']
                ],
            ascending=[True, True]
            )
        dfOldData[CONFIG[configSheetReaderItem]['RankCheck']] = (
            dfOldData[CONFIG[configSheetReaderItem]['RankCheck']]
            ).str.replace("AAAA", "None")
        dfOldData = dfOldData.reset_index(drop=True)
    #Section to be removed post migration
    DFNEWDATA = pd.DataFrame(newData)
    DFNEWDATA = DFNEWDATA[DFNEWDATA[CONFIG[configSheetReaderItem]['nullCheck']].notnull()]
    CompareResult = diff_pd(dfOldData, DFNEWDATA, NEWFILENAME, configSheetReaderItem)
    COMPARERESULTAPPENDED = COMPARERESULTAPPENDED.append(
        CompareResult,
        ignore_index=True,
        sort=True
        )

    SheetProtectionResult = check_sheet_protection(
        NEWFILEPATH, NEWFILENAME, CONFIG[configSheetReaderItem]['SheetName']
        )
    SHEETPROTECTIONRESULTAPPENDED = SHEETPROTECTIONRESULTAPPENDED.append(
        SheetProtectionResult,
        ignore_index=True,
        sort=True
        )

COMPARERESULTAPPENDED = COMPARERESULTAPPENDED[['plan', 'section', 'Result', 'from', 'to']]
SHEETPROTECTIONRESULTAPPENDED = SHEETPROTECTIONRESULTAPPENDED[['plan', 'sheet', 'Result']]
print(COMPARERESULTAPPENDED)
print(SHEETPROTECTIONRESULTAPPENDED)
#Data Reading & Testing Phase Ends

#Data Output write Phase starts
WRITER = pd.ExcelWriter('TestResults.xlsx', engine='xlsxwriter')
COMPARERESULTAPPENDED.to_excel(WRITER, sheet_name='CompareResult', index=False)
SHEETPROTECTIONRESULTAPPENDED.to_excel(WRITER, sheet_name='SheetProtectionResult', index=False)

# Close the Pandas Excel WRITER and output the Excel file.
WRITER.save()
#Data Output write Phase ends