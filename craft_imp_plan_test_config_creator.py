'''
Config file creator
'''
import configparser
CONFIG = configparser.ConfigParser()
CONFIG['ConnectionString OldFilePath'] = {
    'sourcefilepath':
        'C:\\Users\\320014135\\OneDrive - Philips\\SW CoE\\2019\\dotCraft\\SP To OneDrive\\'
}
CONFIG['ConnectionString NewFilePath'] = {
    'sourcefilepath':
        'C:\\Users\\320014135\\OneDrive - Philips\\SW CoE\\2019\\dotCraft\\\
AssessmentFileGeneratorAutomation\\Output\\'
}

CONFIG['Readsheet Score Card 2019'] = {'skiprows': 6,
                                       'nrows': 16,
                                       'columns': ['Attribute,Minimum,Best'],
                                       'SheetName': 'Score Card',
                                       'parse_cols': 'B:D',
                                       'nullCheck': 'Attribute',
                                       'DFNum': 1}
CONFIG['Readsheet Score Card 2020'] = {'skiprows': 6,
                                       'nrows': 16,
                                       'columns': ['Attribute.1,Minimum.1,Best.1'],
                                       'SheetName': 'Score Card',
                                       'parse_cols': 'I:K',
                                       'nullCheck': 'Attribute.1',
                                       'DFNum': 2}
CONFIG['Readsheet ProjectList2019'] = {'skiprows': 27,
                                       'nrows': 100,
                                       'columns': ['Project Name/Team Name,\
FTE/People,Profile'],
                                       'SheetName': 'Score Card',
                                       'parse_cols': 'B:D',
                                       'nullCheck': 'Project Name/Team Name',
                                       'RankCheck': 'Profile',
                                       'DFNum': 3}
CONFIG['Readsheet ProjectList2020'] = {'skiprows': 27,
                                       'nrows': 100,
                                       'columns': ['Project Name/Team Name.1,\
FTE/People.1,Profile.1'],
                                       'SheetName': 'Score Card',
                                       'parse_cols': 'I:K',
                                       'nullCheck': 'Project Name/Team Name.1',
                                       'RankCheck': 'Profile.1',
                                       'DFNum': 4}
CONFIG['Readsheet Code Quality Hygiene'] = {'skiprows': 14,
                                            'columns': ['Sl No,Attribute,Action,\
Owner,End Date,Status,Remarks'],
                                            'nrows': 5000,
                                            'SheetName': 'Code Quality Hygiene',
                                            'parse_cols': 'A:H',
                                            'nullCheck': 'Action',
                                            'DFNum': 5}
CONFIG['Readsheet Test Automation'] = {'skiprows': 14,
                                       'columns': ['Sl No,Attribute,Action,\
Owner,End Date,Status,Remarks'],
                                       'nrows': 5000,
                                       'SheetName': 'Test Automation',
                                       'parse_cols': 'A:H',
                                       'nullCheck': 'Action',
                                       'DFNum': 6}
CONFIG['Readsheet Continuous Integration'] = {'skiprows': 14,
                                              'columns': ['Sl No,Attribute,Action,\
Owner,End Date,Status,Remarks'],
                                              'nrows': 5000,
                                              'SheetName': 'Continuous Integration',
                                              'parse_cols': 'A:H',
                                              'nullCheck': 'Action',
                                              'DFNum': 7}
CONFIG['Readsheet Rationale'] = {'skiprows': 6,
                                 'columns': ['Attribute,Profile,\
Rationale for exceptions,Link to Score Card,Reviewed with SW CoE'],
                                 'nrows': 5000,
                                 'SheetName': 'Rationale',
                                 'parse_cols': 'E:I',
                                 'nullCheck': 'Rationale for exceptions',
                                 'DFNum': 8}
CONFIG['Readsheet Code&Tooling Info DefectDensity'] = {'skiprows': 10,
                                                       'columns': ['Major Release Defect Density,\
Rel. 1,Rel. 2,Rel. 3'],
                                                       'nrows': 3,
                                                       'SheetName': 'Code&Tooling Info',
                                                       'parse_cols': 'C:F',
                                                       'nullCheck': 'Major Release Defect Density',
                                                       'DFNum': 9}
CONFIG['Readsheet Code&Tooling Info CodeSize'] = {'skiprows': 6,
                                                  'columns': ['Code Size,KLOCs,\
Targets for reducing overall code size'],
                                                  'nrows': 3,
                                                  'SheetName': 'Code&Tooling Info',
                                                  'parse_cols': 'M:O',
                                                  'nullCheck': 'Code Size',
                                                  'DFNum': 10}
CONFIG['Readsheet Code&Tooling Info SafeAgile'] = {'skiprows': 26,
                                                   'columns': ['SAFE/Agile Data,Status'],
                                                   'nrows': 3,
                                                   'SheetName': 'Code&Tooling Info',
                                                   'parse_cols': 'C:D',
                                                   'nullCheck': 'SAFE/Agile Data',
                                                   'DFNum': 11}
CONFIG['Readsheet Code&Tooling Info SecurityActions'] = {'skiprows': 28,
                                                         'columns': ['Security,Action,\
Owner,End Date'],
                                                         'nrows': 5000,
                                                         'SheetName': 'Code&Tooling Info',
                                                         'parse_cols': 'M:P',
                                                         'nullCheck': 'Action',
                                                         'DFNum': 13}
CONFIG['Readsheet Technical Debt DevOpsActions'] = {'skiprows': 21,
                                                    'columns': ['DevOps,Action,Owner,End Date'],
                                                    'nrows': 5000,
                                                    'SheetName': 'Technical Debt',
                                                    'parse_cols': 'C:F',
                                                    'nullCheck': 'Action',
                                                    'DFNum': 14}
CONFIG['Readsheet Technical Debt TechDebtActions'] = {'skiprows': 21,
                                                      'columns': ['Tech Debt,Action.1,\
Owner.1,End Date.1'],
                                                      'nrows': 5000,
                                                      'SheetName': 'Technical Debt',
                                                      'parse_cols': 'I:L',
                                                      'nullCheck': 'Action.1',
                                                      'DFNum': 15}
CONFIG['Readsheet Technical Debt ReqBacklogActions'] = {'skiprows': 21,
                                                        'columns': ['Requirements and Backlogs,\
Action.2,Owner.2,End Date.2'],
                                                        'nrows': 5000,
                                                        'SheetName': 'Technical Debt',
                                                        'parse_cols': 'O:R',
                                                        'nullCheck': 'Action.2',
                                                        'DFNum': 16}
CONFIG['Readsheet ApprovalTable'] = {'skiprows': 10,
                                     'columns': ['Role,Name,Approved'],
                                     'nrows': 6,
                                     'SheetName': 'Cover',
                                     'parse_cols': 'B:D',
                                     'nullCheck': 'Role',
                                     'DFNum': 17}
CONFIG['Readsheet Code&Tooling Info ToolInfo'] = {'skiprows': 22,
                                                  'columns': ['Tooling Info,Tools'],
                                                  'nrows': 5000,
                                                  'SheetName': 'Code&Tooling Info',
                                                  'parse_cols': 'I:J',
                                                  'nullCheck': 'Tooling Info',
                                                  'DFNum': 12}
with open('ExcelComparatorConfigFile.ini', 'w') as configfile:
    CONFIG.write(configfile)
