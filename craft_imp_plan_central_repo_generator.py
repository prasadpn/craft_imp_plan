"""
Consolidates the data from all the improvement plans and then loads the data
into one central repository (excel file in this case)
Takes input from the configuration file
"""
#Import statements starts
import os
import configparser
import pandas as pd
import numpy as np
from craft_imp_plan_migration_related_changes import migration_related_changes
#Import statements ends

#Function declaration Phase Starts
def config_string_converter(config_string_input):
    """
    Takes the string as input.  Helps it convert into a list.
    """
    config_string_output = config_string_input[2:-2]
    config_string_output = config_string_output.split(",")
    return config_string_output

def filter_out_column_based_on_values(df_name, column_name, filter_value, filtercriteria):
    """
    Filters value on the the criteria
    """
    if filtercriteria == 'containsstr':
        filtered_df = DFS[df_name][~DFS[df_name][column_name].str.contains(
            filter_value, na=False
            )]
    elif filtercriteria == 'intdatatype':
        filtered_df = DFS[df_name][~DFS[df_name][column_name].apply(
            lambda x: isinstance(x, (int, np.int64, float, np.float64)))]
    elif filtercriteria == 'isin':
        filtered_df = DFS[df_name][DFS[df_name][column_name].isin(
            filter_value
            )]
    else:
        filtered_df = DFS[df_name]
    return filtered_df

#Function declaration Phase Ends

#Preprocessing Preparation Phase Starts
CONFIG = configparser.ConfigParser()
CONFIG.read('ImpPlanReaderConfig.ini')
IMPPLANSOURCEFILEPATH = CONFIG['ConnectionString SourceFilePath']['sourcefilepath']
DFNAMES = CONFIG['DF Names']['DFNames']
DFNAMES = DFNAMES.split(",")
CENTRALDATAOUTPUTSHEETNAMES = config_string_converter(
    CONFIG['Write Output CentralData']['SheetName']
    )
DFS = {}
DFS_TOPRINT = {}
DFS_TOEXCEPTION = DFNAMES
ALL_XLSM_FILES = list(filter(lambda x: x.endswith('.xlsm'), os.listdir(IMPPLANSOURCEFILEPATH)))
CONFIGSHEETREADER = list(filter(lambda x: x.startswith('Readsheet'), CONFIG.sections()))
CONFIGDATACHECKS = list(filter(lambda x: x.startswith('DataCheck'), CONFIG.sections()))
#Preprocessing Preparation Phase Ends

#Data Reading Phase Starts
for configSheetReaderItem in CONFIGSHEETREADER:
    df_appended = []
    for impPlanFileReaderItem in ALL_XLSM_FILES:
        FilePath = IMPPLANSOURCEFILEPATH + impPlanFileReaderItem
        data = pd.read_excel(
            FilePath, sheet_name=CONFIG[configSheetReaderItem]['SheetName'],
            skiprows=int(CONFIG[configSheetReaderItem]['skiprows']),
            nrows=int(CONFIG[configSheetReaderItem]['nrows']),
            usecols=CONFIG[configSheetReaderItem]['parse_cols']
            )
        columnNames = config_string_converter(
            CONFIG[configSheetReaderItem]['columns']
            )
        df = pd.DataFrame(
            data, columns=columnNames
            )
        df = df[df[CONFIG[configSheetReaderItem]['nullCheck']].notnull()]
        #Section to be removed once the migration of the template is done
        if configSheetReaderItem in ('Readsheet ProjectList2019', 'Readsheet ProjectList2020'):
            df = migration_related_changes(
                df,
                CONFIG[configSheetReaderItem]['nullCheck'],
                CONFIG[configSheetReaderItem]['RankCheck']
                )
        #Section to be removed once the migration of the template is done
        df['Source.Name'] = impPlanFileReaderItem
        df_appended.append(df)
    df_appended = pd.concat(df_appended)
    DFS[DFNAMES[int(CONFIG[configSheetReaderItem]['DFNum'])]] = df_appended
    del df
    del df_appended
#Data Reading Phase Ends

#Data Manipulation Phase Starts
#DFS specific data manipulation
DFS['Cover'] = DFS['Cover'].pivot(
    index='Source.Name', columns='Org Structure', values='Name'
    )

DFS['Cover']['Business Group'] = (
    (DFS['Cover']['Business Group'].str.upper()).str.rstrip()
    ).str.lstrip()
DFS['Cover']['Business Unit'] = (
    (DFS['Cover']['Business Unit'].str.upper()).str.rstrip()
    ).str.lstrip()

DFS['Cover']['BGBUID'] = DFS['Cover']['Business Group'] + " - " + DFS['Cover']['Business Unit']

DFS['ProjectList2019'] = filter_out_column_based_on_values(
    'ProjectList2019', 'Project Name/Team Name', '<Project', 'containsstr'
    )
DFS['ProjectList2020'] = filter_out_column_based_on_values(
    'ProjectList2020', 'Project Name/Team Name.1', '<Project', 'containsstr'
    )
DFS['ProjectList2019'] = filter_out_column_based_on_values(
    'ProjectList2019', 'Profile', None, 'intdatatype'
    )
DFS['ProjectList2020'] = filter_out_column_based_on_values(
    'ProjectList2020', 'Profile.1', None, 'intdatatype'
    )

#Actions - addition of ParentAttribute & consolidating the 3 DFS into 1
PARENTATTRIBUTECQ = CONFIG['ParentAttribute Name']['ParentAttributeCQ']
PARENTATTRIBUTETA = CONFIG['ParentAttribute Name']['ParentAttributeTA']
PARENTATTRIBUTECICD = CONFIG['ParentAttribute Name']['ParentAttributeCICD']
DFS['CodeQualityHygiene']['ParentAttribute'] = PARENTATTRIBUTECQ[2:-2]
DFS['TestAutomation']['ParentAttribute'] = PARENTATTRIBUTETA[2:-2]
DFS['ContinuousIntegration']['ParentAttribute'] = PARENTATTRIBUTECICD[2:-2]

FRAMES = [DFS['CodeQualityHygiene'], DFS['TestAutomation'], DFS['ContinuousIntegration']]
DFS['Actions2019'] = DFS.pop('CodeQualityHygiene')
DFS['Actions2019'] = pd.concat(FRAMES)

#Deletes the Continuous Integration DataFrame
del DFS['TestAutomation']
#Deletes the Test Automation DataFrame
del DFS['ContinuousIntegration']

print('Filtering based on Column values for data FRAMES to be printed starts')

DFS_TOPRINT = DFS
#configDataChecksItemCounter = 0
for configDataChecksItem in CONFIGDATACHECKS:
    DFSNam = CENTRALDATAOUTPUTSHEETNAMES[int(
        CONFIG[configDataChecksItem]['DFNumber']
        )]
    DFSCol = CONFIG[configDataChecksItem]['Column']
    DFSColVal = config_string_converter(
        CONFIG[configDataChecksItem]['ColVal']
        )
    DFS_TOPRINT[DFSNam] = filter_out_column_based_on_values(
        DFSNam, DFSCol, DFSColVal, 'isin'
        )

MERGEDFCOUNTER = len(CONFIGSHEETREADER)

for x in range(MERGEDFCOUNTER-3):
    x = x + 1
    DFS_TOPRINT[CENTRALDATAOUTPUTSHEETNAMES[x]] = pd.merge(
        DFS_TOPRINT[CENTRALDATAOUTPUTSHEETNAMES[x]],
        DFS_TOPRINT['Cover'],
        on='Source.Name', how='inner'
        )

#To get from config file once all the columns are decided
DFS_TOPRINT['TargetScore2019']['Joinkey'] = (
    DFS_TOPRINT['TargetScore2019']['BGBUID'] +
    DFS_TOPRINT['TargetScore2019']['Attribute']
    )
DFS_TOPRINT['TargetScore2020']['Joinkey'] = (
    DFS_TOPRINT['TargetScore2020']['BGBUID'] +
    DFS_TOPRINT['TargetScore2020']['Attribute.1']
    )
DFS_TOPRINT['ProjectList2019']['Joinkey'] = (
    DFS_TOPRINT['ProjectList2019']['BGBUID'] +
    (
        (DFS_TOPRINT['ProjectList2019']['Rank']).astype(int)).astype(str)
    )
DFS_TOPRINT['ProjectList2020']['Joinkey'] = (
    DFS_TOPRINT['ProjectList2020']['BGBUID'] +
    (
        (DFS_TOPRINT['ProjectList2020']['Rank']).astype(int)).astype(str)
    )
DFS_TOPRINT['Actions2019']['Joinkey'] = (
    DFS_TOPRINT['Actions2019']['BGBUID'] +
    DFS_TOPRINT['Actions2019']['ParentAttribute'] +
    (
        (DFS_TOPRINT['Actions2019']['Sl No']).astype(int)).astype(str)
    )
DFS_TOPRINT['Rationale']['Joinkey'] = (
    DFS_TOPRINT['Rationale']['BGBUID'] +
    DFS_TOPRINT['Rationale']['Attribute'] + DFS_TOPRINT['Rationale']['Profile']
    )
DFS_TOPRINT['DefectDensity']['Joinkey'] = (
    DFS_TOPRINT['DefectDensity']['BGBUID'] +
    DFS_TOPRINT['DefectDensity']['Major Release Defect Density']
    )
DFS_TOPRINT['CodeSize']['Joinkey'] = (
    DFS_TOPRINT['CodeSize']['BGBUID'] +
    DFS_TOPRINT['CodeSize']['Code Size']
    )
DFS_TOPRINT['SafeAgile']['Joinkey'] = (
    DFS_TOPRINT['SafeAgile']['BGBUID'] +
    DFS_TOPRINT['SafeAgile']['SAFE/Agile Data']
    )
DFS_TOPRINT['ToolInfo']['Joinkey'] = (
    DFS_TOPRINT['ToolInfo']['BGBUID'] +
    DFS_TOPRINT['ToolInfo']['Tooling Info']
    )
DFS_TOPRINT['SecurityActions']['Joinkey'] = (
    DFS_TOPRINT['SecurityActions']['BGBUID'] +
    DFS_TOPRINT['SecurityActions']['Security']
    )
DFS_TOPRINT['DevOpsActions']['Joinkey'] = (
    DFS_TOPRINT['DevOpsActions']['BGBUID'] +
    DFS_TOPRINT['DevOpsActions']['DevOps']
    )
DFS_TOPRINT['TechDebtActions']['Joinkey'] = (
    DFS_TOPRINT['TechDebtActions']['BGBUID'] +
    DFS_TOPRINT['TechDebtActions']['Tech Debt']
    )
DFS_TOPRINT['ReqBacklogActions']['Joinkey'] = (
    DFS_TOPRINT['ReqBacklogActions']['BGBUID'] +
    DFS_TOPRINT['ReqBacklogActions']['Requirements and Backlogs']
    )
DFS_TOPRINT['ApprovalTable']['Joinkey'] = (
    DFS_TOPRINT['ApprovalTable']['BGBUID'] +
    DFS_TOPRINT['ApprovalTable']['Role']
    )

DFS_TOPRINT['TargetScore2020'].rename(
    columns={'Attribute.1':'Attribute', 'Minimum.1':'Minimum', 'Best.1':'Best'},
    inplace=True
    )
DFS_TOPRINT['ProjectList2020'].rename(
    columns={'Project Name/Team Name.1':'Project Name/Team Name',
             'FTE/People.1':'FTE/People', 'Profile.1':'Profile'},
    inplace=True
    )
DFS_TOPRINT['TechDebtActions'].rename(
    columns={'Action.1':'Action', 'Owner.1':'Owner', 'End Date.1':'End Date'},
    inplace=True
    )
DFS_TOPRINT['ReqBacklogActions'].rename(
    columns={'Action.1':'Action', 'Owner.1':'Owner', 'End Date.1':'End Date'},
    inplace=True
    )

#Data Manipulation Phase Ends

#Data Output write Phase starts

print('Output writing to excel starts for Central Repo')
# Create a Pandas Excel writer using XlsxWriter as the engine.
WRITER = pd.ExcelWriter('CentralDataStore.xlsx', engine='xlsxwriter')

WRITEFCOUNTER = len(CENTRALDATAOUTPUTSHEETNAMES)

# Write each dataframe to a different worksheet.
for y in range(WRITEFCOUNTER):
    #worksheet = workbook.add_worksheet(CENTRALDATAOUTPUTSHEETNAMES[0])
    DFS_TOPRINT[CENTRALDATAOUTPUTSHEETNAMES[y]].to_excel(
        WRITER,
        sheet_name=CENTRALDATAOUTPUTSHEETNAMES[y], index=False
        )

# Close the Pandas Excel writer and output the Excel file.
WRITER.save()
#Data Output write Phase Ends
