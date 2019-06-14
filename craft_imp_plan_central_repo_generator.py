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

def filter_out_column_containsstr(df_name, column_name, filter_value):
    """
    Filters value on the the criteria
    """
    filtered_df = DFS[df_name][~DFS[df_name][column_name].str.contains(
        filter_value, na=False
        )]
    return filtered_df

def filter_out_column_intdata(df_name, column_name):
    """
    Filters value on the the criteria
    """
    filtered_df = DFS[df_name][~DFS[df_name][column_name].apply(
        lambda x: isinstance(x, (int, np.int64, float, np.float64)))]
    return filtered_df

def filter_out_column_isin(df_name, column_name, filter_value):
    """
    Filters value on the the criteria
    """
    filtered_df = DFS[df_name][DFS[df_name][column_name].isin(
        filter_value
        )]
    return filtered_df

def create_source_dict(config_sheet_reader, all_files, imp_plan_source_file_path):
    """
    Reads & consolidates all the improvement plans
    into a dictionary
    """
    func_config_sheet_reader = config_sheet_reader
    func_all_files = all_files
    func_imp_plan_source_file_path = imp_plan_source_file_path
    dict_src_data = {}
    for config_sheet_reader_item in func_config_sheet_reader:
        df_appended = []
        for imp_plan_file_reader_item in func_all_files:
            file_path = func_imp_plan_source_file_path + imp_plan_file_reader_item
            data = pd.read_excel(
                file_path, sheet_name=CONFIG[config_sheet_reader_item]['SheetName'],
                skiprows=int(CONFIG[config_sheet_reader_item]['skiprows']),
                nrows=int(CONFIG[config_sheet_reader_item]['nrows']),
                usecols=CONFIG[config_sheet_reader_item]['parse_cols']
                )
            column_names = config_string_converter(
                CONFIG[config_sheet_reader_item]['columns']
                )
            imp_plan_df = pd.DataFrame(
                data, columns=column_names
                )
            imp_plan_df = imp_plan_df[
                imp_plan_df[
                    CONFIG[config_sheet_reader_item]['nullCheck']
                    ].notnull()]
            #Section to be removed once the migration of the template is done
            if config_sheet_reader_item in (
                    'Readsheet ProjectList2019',
                    'Readsheet ProjectList2020'
                ):
                imp_plan_df = migration_related_changes(
                    imp_plan_df,
                    CONFIG[config_sheet_reader_item]['nullCheck'],
                    CONFIG[config_sheet_reader_item]['RankCheck']
                    )
            #Section to be removed once the migration of the template is done
            imp_plan_df['Source.Name'] = imp_plan_file_reader_item
            df_appended.append(imp_plan_df)
        df_appended = pd.concat(df_appended)
        dict_src_data[DFNAMES[int(CONFIG[config_sheet_reader_item]['DFNum'])]] = df_appended
        del imp_plan_df
        del df_appended
    return dict_src_data

def create_master_dict(config_masterdata_reader, masterdata_source_file_path):
    """
    Reads & consolidates all the improvement plans
    into a dictionary
    """
    func_config_masterdata_reader = config_masterdata_reader
    func_masterdata_source_file_path = masterdata_source_file_path
    dict_master_data = {}
    for config_masterdata_reader_item in func_config_masterdata_reader:
        data = pd.read_excel(
            func_masterdata_source_file_path,
            sheet_name=CONFIG[config_masterdata_reader_item]['SheetName'],
            skiprows=int(CONFIG[config_masterdata_reader_item]['skiprows']),
            nrows=int(CONFIG[config_masterdata_reader_item]['nrows']),
            usecols=CONFIG[config_masterdata_reader_item]['parse_cols']
            )
        column_names = config_string_converter(
            CONFIG[config_masterdata_reader_item]['columns']
            )
        masterdata_df = pd.DataFrame(
            data, columns=column_names
            )
        masterdata_df = masterdata_df[
            masterdata_df[
                CONFIG[config_masterdata_reader_item]['nullCheck']
                ].notnull()]
        dict_master_data[
            MASTERDFNAMES[int(CONFIG[config_masterdata_reader_item]['DFNum'])]
            ] = masterdata_df
        del masterdata_df
    return dict_master_data

def prepare_cover_tab(cover_dict, bgbulist_dict):
    """
    Prepares the data for the cover tab
    """
    cover_dict = cover_dict.pivot(
        index='Source.Name', columns='Org Structure', values='Name')
    cover_dict['Business Group'] = (
        (cover_dict['Business Group'].str.upper()).str.rstrip()).str.lstrip()
    cover_dict['Business Unit'] = (
        (cover_dict['Business Unit'].str.upper()).str.rstrip()).str.lstrip()
    cover_dict['Cluster'] = (
        (cover_dict['Cluster'].str.upper()).str.rstrip()).str.lstrip()
    bgbulist_dict['Business Group'] = (
        (bgbulist_dict['Business Group'].str.upper()).str.rstrip()
        ).str.lstrip()
    bgbulist_dict['Business Unit'] = (
        (bgbulist_dict['Business Unit'].str.upper()).str.rstrip()
        ).str.lstrip()
    bgbulist_dict['Cluster'] = (
        (bgbulist_dict['Cluster'].str.upper()).str.rstrip()
        ).str.lstrip()
    cover_dict['Source.Name'] = cover_dict.index
    cover_dict = pd.merge(cover_dict, bgbulist_dict, on='ID', how='inner')
    cover_dict['Business Unit_x'] = cover_dict['Business Unit_y']
    cover_dict['Business Group_x'] = cover_dict['Business Group_y']
    cover_dict['Cluster_x'] = cover_dict['Cluster_y']
    cover_dict.rename(columns={
        'Business Unit_x':'Business Unit',
        'Business Group_x':'Business Group',
        'Cluster_x':'Cluster'}, inplace=True)
    cover_dict['BGBUID'] = cover_dict['Business Group'] + " - " + cover_dict['Business Unit']
    cover_dict = cover_dict[[
        'ID', 'Source.Name', 'Template Version',
        'BGBUID', 'Cluster', 'Business Group', 'Business Unit'
        ]]
    return cover_dict

def frame_specific_manipulation(config_data_checks, central_data_output_sheet_names):
    """
    Data Manipulation
    """
    write_output_dict = {}
    DFS['Cover'] = prepare_cover_tab(DFS['Cover'], DFS['BGBUList'])

    DFS['ProjectList2019'] = filter_out_column_containsstr(
        'ProjectList2019', 'Project Name/Team Name', '<Project'
        )
    DFS['ProjectList2020'] = filter_out_column_containsstr(
        'ProjectList2020', 'Project Name/Team Name.1', '<Project'
        )
    DFS['ProjectList2019'] = filter_out_column_intdata(
        'ProjectList2019', 'Profile'
        )
    DFS['ProjectList2020'] = filter_out_column_intdata(
        'ProjectList2020', 'Profile.1'
        )
    #Actions - addition of ParentAttribute & consolidating the 3 DFS into 1
    parent_attribute_cq = CONFIG['ParentAttribute Name']['ParentAttributeCQ']
    parent_attribute_ta = CONFIG['ParentAttribute Name']['ParentAttributeTA']
    parent_attribute_cicd = CONFIG['ParentAttribute Name']['ParentAttributeCICD']
    DFS['CodeQualityHygiene']['ParentAttribute'] = parent_attribute_cq[2:-2]
    DFS['TestAutomation']['ParentAttribute'] = parent_attribute_ta[2:-2]
    DFS['ContinuousIntegration']['ParentAttribute'] = parent_attribute_cicd[2:-2]
    intermediate_frame = [
        DFS['CodeQualityHygiene'],
        DFS['TestAutomation'],
        DFS['ContinuousIntegration']
        ]
    DFS['Actions2019'] = DFS.pop('CodeQualityHygiene')
    DFS['Actions2019'] = pd.concat(intermediate_frame)
    #Deletes the Continuous Integration DataFrame
    del DFS['TestAutomation']
    #Deletes the Test Automation DataFrame
    del DFS['ContinuousIntegration']
    write_output_dict = DFS
    #configDataChecksItemCounter = 0
    for config_data_checks_item in config_data_checks:
        dfs_name = central_data_output_sheet_names[int(
            CONFIG[config_data_checks_item]['DFNumber']
            )]
        dfs_col = CONFIG[config_data_checks_item]['Column']
        dfs_col_val = config_string_converter(
            CONFIG[config_data_checks_item]['ColVal']
            )
        write_output_dict[dfs_name] = filter_out_column_isin(
            dfs_name, dfs_col, dfs_col_val
            )
    merged_counter = len(CONFIGSHEETREADER)
    for output_sheet_counter in range(merged_counter-3):
        output_sheet_counter = output_sheet_counter + 1
        write_output_dict[central_data_output_sheet_names[output_sheet_counter]] = pd.merge(
            write_output_dict[central_data_output_sheet_names[output_sheet_counter]],
            write_output_dict['Cover'][['Source.Name', 'Business Group',
                                        'Business Unit', 'Cluster', 'BGBUID']],
            on='Source.Name', how='inner'
            )
    #To get from config file once all the columns are decided
    write_output_dict['TargetScore2019']['Joinkey'] = (
        write_output_dict['TargetScore2019']['BGBUID'] +
        write_output_dict['TargetScore2019']['Attribute']
        )
    write_output_dict['TargetScore2020']['Joinkey'] = (
        write_output_dict['TargetScore2020']['BGBUID'] +
        write_output_dict['TargetScore2020']['Attribute.1']
        )
    write_output_dict['ProjectList2019']['Joinkey'] = (
        write_output_dict['ProjectList2019']['BGBUID'] +
        (
            (write_output_dict['ProjectList2019']['Rank']).astype(int)).astype(str)
        )
    write_output_dict['ProjectList2020']['Joinkey'] = (
        write_output_dict['ProjectList2020']['BGBUID'] +
        (
            (write_output_dict['ProjectList2020']['Rank']).astype(int)).astype(str)
        )
    write_output_dict['Actions2019']['Joinkey'] = (
        write_output_dict['Actions2019']['BGBUID'] +
        write_output_dict['Actions2019']['ParentAttribute'] +
        (
            (write_output_dict['Actions2019']['Sl No']).astype(int)).astype(str)
        )
    write_output_dict['Rationale']['Joinkey'] = (
        write_output_dict['Rationale']['BGBUID'] +
        write_output_dict['Rationale']['Attribute'] + write_output_dict['Rationale']['Profile']
        )
    write_output_dict['DefectDensity']['Joinkey'] = (
        write_output_dict['DefectDensity']['BGBUID'] +
        write_output_dict['DefectDensity']['Major Release Defect Density']
        )
    write_output_dict['CodeSize']['Joinkey'] = (
        write_output_dict['CodeSize']['BGBUID'] +
        write_output_dict['CodeSize']['Code Size']
        )
    write_output_dict['SafeAgile']['Joinkey'] = (
        write_output_dict['SafeAgile']['BGBUID'] +
        write_output_dict['SafeAgile']['SAFE/Agile Data']
        )
    write_output_dict['ToolInfo']['Joinkey'] = (
        write_output_dict['ToolInfo']['BGBUID'] +
        write_output_dict['ToolInfo']['Tooling Info']
        )
    write_output_dict['SecurityActions']['Joinkey'] = (
        write_output_dict['SecurityActions']['BGBUID'] +
        write_output_dict['SecurityActions']['Security']
        )
    write_output_dict['DevOpsActions']['Joinkey'] = (
        write_output_dict['DevOpsActions']['BGBUID'] +
        write_output_dict['DevOpsActions']['DevOps']
        )
    write_output_dict['TechDebtActions']['Joinkey'] = (
        write_output_dict['TechDebtActions']['BGBUID'] +
        write_output_dict['TechDebtActions']['Tech Debt']
        )
    write_output_dict['ReqBacklogActions']['Joinkey'] = (
        write_output_dict['ReqBacklogActions']['BGBUID'] +
        write_output_dict['ReqBacklogActions']['Requirements and Backlogs']
        )
    write_output_dict['ApprovalTable']['Joinkey'] = (
        write_output_dict['ApprovalTable']['BGBUID'] +
        write_output_dict['ApprovalTable']['Role']
        )
    write_output_dict['TargetScore2020'].rename(
        columns={'Attribute.1':'Attribute', 'Minimum.1':'Minimum', 'Best.1':'Best'},
        inplace=True
        )
    write_output_dict['ProjectList2020'].rename(
        columns={'Project Name/Team Name.1':'Project Name/Team Name',
                 'FTE/People.1':'FTE/People', 'Profile.1':'Profile'},
        inplace=True
        )
    write_output_dict['TechDebtActions'].rename(
        columns={'Action.1':'Action', 'Owner.1':'Owner', 'End Date.1':'End Date'},
        inplace=True
        )
    write_output_dict['ReqBacklogActions'].rename(
        columns={'Action.1':'Action', 'Owner.1':'Owner', 'End Date.1':'End Date'},
        inplace=True
        )
    return write_output_dict

def write_output(dict_toprint, output_sheet_names):
    """
    Writes dictionary output to excel
    """
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    output_sheet = output_sheet_names
    sheet_writer = pd.ExcelWriter('CentralDataStore.xlsx', engine='xlsxwriter')
    writer_counter = len(output_sheet)
    # Write each dataframe to a different worksheet.
    for sheet_increment in range(writer_counter):
        #worksheet = workbook.add_worksheet(CENTRALDATAOUTPUTSHEETNAMES[0])
        dict_toprint[output_sheet[sheet_increment]].to_excel(
            sheet_writer,
            sheet_name=output_sheet[sheet_increment], index=False
            )
    # Close the Pandas Excel writer and output the Excel file.
    sheet_writer.save()
    #Data Output write Phase Ends
    return "Excel generated.  Please check."

#Function declaration Phase Ends

CONFIG = configparser.ConfigParser()
CONFIG.read('ImpPlanReaderConfig.ini')
IMPPLANSOURCEFILEPATH = CONFIG['ConnectionString SourceFilePath']['sourcefilepath']
MASTERDATASOURCEFILEPATH = CONFIG['ConnectionString MasterDataFilePath']['sourcefilepath']
DFNAMES = CONFIG['DF Names']['DFNames']
DFNAMES = DFNAMES.split(",")
MASTERDFNAMES = CONFIG['MasterDF Names']['DFNames']
MASTERDFNAMES = MASTERDFNAMES.split(",")
CENTRALDATAOUTPUTSHEETNAMES = config_string_converter(
    CONFIG['Write Output CentralData']['SheetName']
    )
DFS = {}
DFS_TOPRINT = {}
DFS_TOEXCEPTION = DFNAMES
ALL_XLSM_FILES = list(filter(lambda x: x.endswith('.xlsm'), os.listdir(IMPPLANSOURCEFILEPATH)))
CONFIGSHEETREADER = list(filter(lambda x: x.startswith('Readsheet'), CONFIG.sections()))
CONFIGMASTERDATAREADER = list(filter(lambda x: x.startswith('Master Readsheet'), CONFIG.sections()))
CONFIGDATACHECKS = list(filter(lambda x: x.startswith('DataCheck'), CONFIG.sections()))


IMPPLANDFS = create_source_dict(CONFIGSHEETREADER, ALL_XLSM_FILES, IMPPLANSOURCEFILEPATH)
MASTERDATADFS = create_master_dict(CONFIGMASTERDATAREADER, MASTERDATASOURCEFILEPATH)
DFS = {**IMPPLANDFS, **MASTERDATADFS}
DFS_TOPRINT = frame_specific_manipulation(CONFIGDATACHECKS, CENTRALDATAOUTPUTSHEETNAMES)
OUTPUT_WRITE_EXCEL = write_output(DFS_TOPRINT, CENTRALDATAOUTPUTSHEETNAMES)
print(OUTPUT_WRITE_EXCEL)
