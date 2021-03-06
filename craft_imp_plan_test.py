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
from craft_imp_plan_migration_related_changes import migration_related_changes
#Import statements ends

#Function declaration Phase Starts
def identify_column_diff(old_list, new_list):
    """
    Compares Column Header values and data types
    """
    old_val = []
    new_val = []
    for index, (old, new) in enumerate(zip(old_list, new_list)):
        if old != new:
            old_val.append(old)
            new_val.append(new)
            print(index)
    return str(old_val), str(new_val)

def make_comparison_dataframe(plan_section, result_message, from_value, to_value, index_value):
    """
    Creates output dataframe for comparision result
    """
    return pd.DataFrame(
        {
            'plan - section': plan_section,
            'Result': result_message, 'from': from_value, 'to': to_value
            },
        index=index_value
    )

def make_protection_dataframe(plan_sheet, result_message, index_value):
    """
    Creates output dataframe for sheet protection result
    """
    return pd.DataFrame(
        {
            'plan - sheet': plan_sheet,
            'Result': result_message
            },
        index=index_value
    )

def compare_improvement_plans(old_plan, new_plan, plan, imp_plan_section):
    """
    Compares improvement plans for differences
    """
    if old_plan.equals(new_plan):
        result_of_compare = make_comparison_dataframe(
            plan + ': ' + imp_plan_section, 'Match', '', '', ['First'])

    elif len(old_plan) != len(new_plan):
        result_of_compare = make_comparison_dataframe(
            plan + ': ' + imp_plan_section, 'Missmatch row count',
            str(len(old_plan))+' rows', str(len(new_plan))+' rows', ['First'])

    elif len(old_plan.columns) != len(new_plan.columns):
        result_of_compare = make_comparison_dataframe(
            plan + ': ' + imp_plan_section, 'Missmatch column count',
            str(len(old_plan.columns))+' columns', str(len(new_plan.columns))+' columns',
            ['First'])

    elif (old_plan.columns == new_plan.columns).all() and any(old_plan.dtypes != new_plan.dtypes):
        old_col_dtypes = list(map(list, zip(list(old_plan.columns.values), list(old_plan.dtypes))))
        new_col_dtypes = list(map(list, zip(list(new_plan.columns.values), list(new_plan.dtypes))))
        old_val_str, new_val_str = identify_column_diff(old_col_dtypes, new_col_dtypes)
        result_of_compare = make_comparison_dataframe(
            plan + ': ' + imp_plan_section, 'Missmatch data type of same columns',
            old_val_str, new_val_str, ['First'])

    elif any(old_plan.columns != new_plan.columns):
        old_plan_columns = list(old_plan.columns.values)
        new_plan_columns = list(new_plan.columns.values)
        old_val_str, new_val_str = identify_column_diff(old_plan_columns, new_plan_columns)
        result_of_compare = make_comparison_dataframe(
            plan + ': ' + imp_plan_section, 'Different Columns Names',
            old_val_str, new_val_str, ['First'])
    else:
        diff_mask = (old_plan != new_plan) & ~(old_plan.isnull() & new_plan.isnull())
        ne_stacked = diff_mask.stack()
        changed = ne_stacked[ne_stacked]
        changed.index.names = ['id', 'col']
        difference_locations = np.where(diff_mask)
        result_of_compare = make_comparison_dataframe(
            plan + ': ' + imp_plan_section, 'Missmatch in data',
            old_plan.values[difference_locations],
            new_plan.values[difference_locations],
            changed.index)
    return result_of_compare

def check_sheet_protection(file_path, plan, sheet):
    '''Sheet protection check'''
    excel_file = win32.gencache.EnsureDispatch('Excel.Application')
    work_book = excel_file.Workbooks.Open(file_path)
    excel_file.Visible = True
    work_sheet = work_book.Worksheets(sheet)
    try:
        work_sheet.Range("A1").Value = "Cell A1"
        result_of_sheetprotection = make_protection_dataframe(
            plan + ': ' + sheet, 'Sheet Not Protected', ['First'])
    except: # pylint: disable=W0702
        result_of_sheetprotection = make_protection_dataframe(
            plan + ': ' + sheet, 'Sheet Protected', ['First'])
    finally:
        excel_file.DisplayAlerts = False
        work_book.Close(False)
        excel_file.Application.Quit()
    return result_of_sheetprotection

def prepare_data_for_test(config_sheet_reader_item, old_path, new_path):
    """
    Prepares improvement plan data for test
    """
    old_plan_path = old_path
    new_plan_path = new_path
    config_sheet_reader_item_key = config_sheet_reader_item
    old_data = pd.read_excel(
        old_plan_path, sheet_name=CONFIG[config_sheet_reader_item_key]['SheetName'],
        skiprows=int(CONFIG[config_sheet_reader_item_key]['skiprows']),
        nrows=int(CONFIG[config_sheet_reader_item_key]['nrows']),
        usecols=CONFIG[config_sheet_reader_item_key]['parse_cols']
        )
    if config_sheet_reader_item_key in ('Readsheet ProjectList2019', 'Readsheet ProjectList2020'):
        new_data = pd.read_excel(
            new_plan_path, sheet_name=CONFIG[config_sheet_reader_item_key]['SheetName'],
            skiprows=int(CONFIG[config_sheet_reader_item_key]['skiprows'])+4,
            nrows=int(CONFIG[config_sheet_reader_item_key]['nrows']),
            usecols=CONFIG[config_sheet_reader_item_key]['parse_cols']
            )
    else:
        new_data = pd.read_excel(
            new_plan_path, sheet_name=CONFIG[config_sheet_reader_item_key]['SheetName'],
            skiprows=int(CONFIG[config_sheet_reader_item_key]['skiprows']),
            nrows=int(CONFIG[config_sheet_reader_item_key]['nrows']),
            usecols=CONFIG[config_sheet_reader_item_key]['parse_cols']
            )

    df_old_data = pd.DataFrame(old_data)
    df_old_data = df_old_data[
        df_old_data[
            CONFIG[config_sheet_reader_item_key]['nullCheck']].notnull()]
    #Section to be removed post migration
    if config_sheet_reader_item_key in ('Readsheet ProjectList2019', 'Readsheet ProjectList2020'):
        df_old_data = migration_related_changes(
            df_old_data, CONFIG[config_sheet_reader_item_key]['nullCheck'],
            CONFIG[config_sheet_reader_item_key]['RankCheck']
            )
        df_old_data = df_old_data[df_old_data.columns[:-1]]
        df_old_data = df_old_data.reset_index(drop=True)
    #Section to be removed post migration
    df_new_data = pd.DataFrame(new_data)
    df_new_data = df_new_data[
        df_new_data[CONFIG[config_sheet_reader_item_key]['nullCheck']].notnull()]
    return df_old_data, df_new_data

def read_test_files(config_sheet_reader, old_file_path, new_file_path, new_file_name):
    """
    Reads improvement plans and create test results
    """
    compare_result_appended = pd.DataFrame()
    sheet_protection_result_appended = pd.DataFrame()
    for config_sheet_reader_item in config_sheet_reader:
        old_imp_plan, new_imp_plan = prepare_data_for_test(
            config_sheet_reader_item,
            old_file_path,
            new_file_path
            )
        compare_result = compare_improvement_plans(
            old_imp_plan, new_imp_plan, new_file_name, config_sheet_reader_item)
        compare_result_appended = compare_result_appended.append(
            compare_result,
            ignore_index=True,
            sort=True
            )
        sheet_protection_result = check_sheet_protection(
            NEWFILEPATH, new_file_name, CONFIG[config_sheet_reader_item]['SheetName']
            )
        sheet_protection_result_appended = sheet_protection_result_appended.append(
            sheet_protection_result,
            ignore_index=True,
            sort=True
            )
    compare_result_appended = compare_result_appended[['plan - section', 'Result', 'from', 'to']]
    sheet_protection_result_appended = sheet_protection_result_appended[['plan - sheet', 'Result']]
    return compare_result_appended, sheet_protection_result_appended

def write_output(compare_test_result, sheetprotect_test_result):
    """
    Writes test results output
    """
    #Data Output write Phase starts
    output_writer = pd.ExcelWriter('TestResults.xlsx', engine='xlsxwriter')
    compare_test_result.to_excel(
        output_writer,
        sheet_name='CompareResult', index=False)
    sheetprotect_test_result.to_excel(
        output_writer,
        sheet_name='SheetProtectionResult', index=False)

    # Close the Pandas Excel WRITER and output the Excel file.
    output_writer.save()
    #Data Output write Phase ends
    return "Test result generated.  Please check."

#Function declaration Phase Ends

#Preprocessing Preparation Phase Starts
CONFIG = configparser.ConfigParser()
CONFIG.read('ImpPlanReaderConfig.ini')

IMPPLANOLDFILEPATH = CONFIG['ConnectionString SourceFilePath']['sourcefilepath']
IMPPLANNEWFILEPATH = CONFIG['ConnectionString NewFilePath']['sourcefilepath']

ALLOUTPUTFILES = list(filter(lambda x: x.endswith('.xlsm'), os.listdir(IMPPLANNEWFILEPATH)))
CONFIGSHEETREADER = list(filter(lambda x: x.startswith('Readsheet'), CONFIG.sections()))

#NEWFILENAME = ALLOUTPUTFILES[randint(0,47)]
NEWFILENAME = ALLOUTPUTFILES[9]
OLDFILENAME = NEWFILENAME  #Once migration is done this line will change

OLDFILEPATH = IMPPLANOLDFILEPATH + OLDFILENAME
NEWFILEPATH = IMPPLANNEWFILEPATH + NEWFILENAME

COMPARE_RESULT, SHEET_PROTECT_RESULT = read_test_files(
    CONFIGSHEETREADER, OLDFILEPATH, NEWFILEPATH, NEWFILENAME)
OUTPUT_WRITE_EXCEL = write_output(
    COMPARE_RESULT, SHEET_PROTECT_RESULT)
print(OUTPUT_WRITE_EXCEL)
