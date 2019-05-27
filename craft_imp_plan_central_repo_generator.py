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
#Import statements ends

#Function declaration Phase Starts
def config_string_converter(config_string_input):
    """
    Takes the string as input.  Helps it convert into a list.
    """
    config_string_output = config_string_input[2:-2]
    config_string_output = config_string_output.split(",")
    return config_string_output

def filter_out_column_based_on_values(df_number, column_name, filter_value, filtercriteria):
    """
    Filters value based on the the criteria provided
    """
    if filtercriteria == 'containsstr':
        filtered_df = DFS[df_number][~DFS[df_number][column_name].str.contains(
            filter_value, na=False
            )]
    elif filtercriteria == 'intdatatype':
        filtered_df = DFS[df_number][~DFS[df_number][column_name].apply(
            lambda x: isinstance(x, (int, np.int64, float, np.float64)))]
    elif filtercriteria == 'isin':
        filtered_df = DFS[df_number][DFS[df_number][column_name].isin(
            filter_value
            )]
    else:
        filtered_df = DFS[df_number]
    return filtered_df

#Function declaration Phase Ends

#Preprocessing Preparation Phase Starts
CONFIG = configparser.ConfigParser()
CONFIG.read('ImpPlanReaderConfig.ini')
IMPPLANSOURCEFILEPATH = CONFIG['ConnectionString SourceFilePath']['sourcefilepath']
DFNAMES = CONFIG['DF Names']['DFNames']
DFNAMES = DFNAMES.split(",")
DFS = DFNAMES
DFS_TOPRINT = []
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
            df = df[~df[CONFIG[configSheetReaderItem]['nullCheck']].str.contains(
                '<Project', na=False
                )]
            df = df[~df[CONFIG[configSheetReaderItem]['RankCheck']].apply(
                lambda x: isinstance(x, (int, np.int64, float, np.float64)))]
            df[CONFIG[configSheetReaderItem]['RankCheck']][
                df[CONFIG[configSheetReaderItem]['nullCheck']] == "Other"
                ] = "AAAA"
            df = df.sort_values(
                [
                    CONFIG[configSheetReaderItem]['RankCheck'],
                    CONFIG[configSheetReaderItem]['nullCheck']
                    ],
                ascending=[True, True]
                )
            df['Rank'] = (
                df[CONFIG[configSheetReaderItem]['RankCheck']]+
                df[CONFIG[configSheetReaderItem]['nullCheck']]
                ).rank()
            df[CONFIG[configSheetReaderItem]['RankCheck']] = (
                df[CONFIG[configSheetReaderItem]['RankCheck']]
                ).str.replace("AAAA", "None")
        #Section to be removed once the migration of the template is done
        df['Source.Name'] = impPlanFileReaderItem
        df_appended.append(df)
    df_appended = pd.concat(df_appended)
    DFS[int(CONFIG[configSheetReaderItem]['DFNum'])] = df_appended
    del df
    del df_appended
#Data Reading Phase Ends

#Data Manipulation Phase Starts
#DFS specific data manipulation
DFS[0] = DFS[0].pivot(
    index='Source.Name', columns='Org Structure', values='Name'
    )

DFS[0]['Business Group'] = (
    (DFS[0]['Business Group'].str.upper()).str.rstrip()
    ).str.lstrip()
DFS[0]['Business Unit'] = (
    (DFS[0]['Business Unit'].str.upper()).str.rstrip()
    ).str.lstrip()

DFS[0]['BGBUID'] = DFS[0]['Business Group'] + " - " + DFS[0]['Business Unit']

DFS[3] = filter_out_column_based_on_values(
    3, 'Project Name/Team Name', '<Project', 'containsstr'
    )
DFS[4] = filter_out_column_based_on_values(
    4, 'Project Name/Team Name.1', '<Project', 'containsstr'
    )
DFS[3] = filter_out_column_based_on_values(
    3, 'Profile', None, 'intdatatype'
    )
DFS[4] = filter_out_column_based_on_values(
    4, 'Profile.1', None, 'intdatatype'
    )

#Actions - addition of ParentAttribute & consolidating the 3 DFS into 1
PARENTATTRIBUTECQ = CONFIG['ParentAttribute Name']['ParentAttributeCQ']
PARENTATTRIBUTETA = CONFIG['ParentAttribute Name']['ParentAttributeTA']
PARENTATTRIBUTECICD = CONFIG['ParentAttribute Name']['ParentAttributeCICD']
DFS[5]['ParentAttribute'] = PARENTATTRIBUTECQ[2:-2]
DFS[6]['ParentAttribute'] = PARENTATTRIBUTETA[2:-2]
DFS[7]['ParentAttribute'] = PARENTATTRIBUTECICD[2:-2]

FRAMES = [DFS[5], DFS[6], DFS[7]]
DFS[5] = pd.concat(FRAMES)

#Deletes the Continuous Integration DataFrame
del DFS[7]
#Deletes the Test Automation DataFrame
del DFS[6]

print('Filtering based on Column values for data FRAMES to be printed starts')

DFS_TOPRINT = DFS
#configDataChecksItemCounter = 0
for configDataChecksItem in CONFIGDATACHECKS:
    DFSNum = int(
        CONFIG[configDataChecksItem]['DFNumber']
        )
    DFSCol = CONFIG[configDataChecksItem]['Column']
    DFSColVal = config_string_converter(
        CONFIG[configDataChecksItem]['ColVal']
        )
    DFS_TOPRINT[DFSNum] = filter_out_column_based_on_values(
        DFSNum, DFSCol, DFSColVal, 'isin'
        )

MERGEDFCOUNTER = len(CONFIGSHEETREADER)

for x in range(MERGEDFCOUNTER-3):
    x = x + 1
    DFS_TOPRINT[x] = pd.merge(DFS_TOPRINT[x], DFS_TOPRINT[0], on='Source.Name', how='inner')

#To get from config file once all the columns are decided
DFS_TOPRINT[1]['Joinkey'] = (
    DFS_TOPRINT[1]['BGBUID'] +
    DFS_TOPRINT[1]['Attribute']
    )
DFS_TOPRINT[2]['Joinkey'] = (
    DFS_TOPRINT[2]['BGBUID'] +
    DFS_TOPRINT[2]['Attribute.1']
    )
DFS_TOPRINT[3]['Joinkey'] = (
    DFS_TOPRINT[3]['BGBUID'] +
    (
        (DFS_TOPRINT[3]['Rank']).astype(int)).astype(str)
    )
DFS_TOPRINT[4]['Joinkey'] = (
    DFS_TOPRINT[4]['BGBUID'] +
    (
        (DFS_TOPRINT[4]['Rank']).astype(int)).astype(str)
    )
DFS_TOPRINT[5]['Joinkey'] = (
    DFS_TOPRINT[5]['BGBUID'] +
    DFS_TOPRINT[5]['ParentAttribute'] +
    (
        (DFS_TOPRINT[5]['Sl No']).astype(int)).astype(str)
    )
DFS_TOPRINT[6]['Joinkey'] = (
    DFS_TOPRINT[6]['BGBUID'] +
    DFS_TOPRINT[6]['Attribute'] + DFS_TOPRINT[6]['Profile']
    )
DFS_TOPRINT[7]['Joinkey'] = (
    DFS_TOPRINT[7]['BGBUID'] +
    DFS_TOPRINT[7]['Major Release Defect Density']
    )
DFS_TOPRINT[8]['Joinkey'] = (
    DFS_TOPRINT[8]['BGBUID'] +
    DFS_TOPRINT[8]['Code Size']
    )
DFS_TOPRINT[9]['Joinkey'] = (
    DFS_TOPRINT[9]['BGBUID'] +
    DFS_TOPRINT[9]['SAFE/Agile Data']
    )
DFS_TOPRINT[10]['Joinkey'] = (
    DFS_TOPRINT[10]['BGBUID'] +
    DFS_TOPRINT[10]['Tooling Info']
    )
DFS_TOPRINT[11]['Joinkey'] = (
    DFS_TOPRINT[11]['BGBUID'] +
    DFS_TOPRINT[11]['Security']
    )
DFS_TOPRINT[12]['Joinkey'] = (
    DFS_TOPRINT[12]['BGBUID'] +
    DFS_TOPRINT[12]['DevOps']
    )
DFS_TOPRINT[13]['Joinkey'] = (
    DFS_TOPRINT[13]['BGBUID'] +
    DFS_TOPRINT[13]['Tech Debt']
    )
DFS_TOPRINT[14]['Joinkey'] = (
    DFS_TOPRINT[14]['BGBUID'] +
    DFS_TOPRINT[14]['Requirements and Backlogs']
    )
DFS_TOPRINT[15]['Joinkey'] = (
    DFS_TOPRINT[15]['BGBUID'] +
    DFS_TOPRINT[15]['Role']
    )

DFS_TOPRINT[2].rename(
    columns={'Attribute.1':'Attribute', 'Minimum.1':'Minimum', 'Best.1':'Best'},
    inplace=True
    )
DFS_TOPRINT[4].rename(
    columns={'Project Name/Team Name.1':'Project Name/Team Name',
             'FTE/People.1':'FTE/People', 'Profile.1':'Profile'},
    inplace=True
    )
DFS_TOPRINT[13].rename(
    columns={'Action.1':'Action', 'Owner.1':'Owner', 'End Date.1':'End Date'},
    inplace=True
    )
DFS_TOPRINT[14].rename(
    columns={'Action.1':'Action', 'Owner.1':'Owner', 'End Date.1':'End Date'},
    inplace=True
    )

#Data Manipulation Phase Ends

#Data Output write Phase starts
CENTRALDATAOUTPUTSHEETNAMES = config_string_converter(
    CONFIG['Write Output CentralData']['SheetName']
    )

print('Output writing to excel starts for Central Repo')
# Create a Pandas Excel writer using XlsxWriter as the engine.
WRITER = pd.ExcelWriter('CentralDataStore.xlsx', engine='xlsxwriter')

WRITEFCOUNTER = len(CENTRALDATAOUTPUTSHEETNAMES)

# Write each dataframe to a different worksheet.
for y in range(WRITEFCOUNTER):
    #worksheet = workbook.add_worksheet(CENTRALDATAOUTPUTSHEETNAMES[0])
    DFS_TOPRINT[y].to_excel(WRITER, sheet_name=CENTRALDATAOUTPUTSHEETNAMES[y], index=False)

# Close the Pandas Excel writer and output the Excel file.
WRITER.save()
#Data Output write Phase Ends
