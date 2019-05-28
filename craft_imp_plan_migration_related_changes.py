'''
Migration related alterations
'''
import numpy as np
def migration_related_changes(base_dataframe, nullchk, rankchk):
    '''
    Rank for each project
    '''
    migration_df = base_dataframe[~base_dataframe[nullchk].str.contains(
                '<Project', na=False
                )]
    migration_df = migration_df[~migration_df[rankchk].apply(
        lambda x: isinstance(x, (int, np.int64, float, np.float64)))]
    migration_df[rankchk][
        migration_df[nullchk] == "Other"
        ] = "AAAA"
    migration_df = migration_df.sort_values(
        [
            rankchk,
            nullchk
            ],
        ascending=[True, True]
        )
    migration_df['Rank'] = (
        migration_df[rankchk]+
        migration_df[nullchk]
        ).rank()
    migration_df[rankchk] = (
        migration_df[rankchk]
        ).str.replace("AAAA", "None")
    return migration_df
