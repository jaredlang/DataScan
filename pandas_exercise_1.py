import re
import pandas as pd

sde_xlsx = pd.ExcelFile(r'H:\nbProjects\LDrive2SDE_Data_Path.xlsx')
sde_df = sde_xlsx.parse(sde_xlsx.sheet_names[0], header=0, index_col=None)

re_LNG = re.compile('\\\\Houston\\\\Mozambique LNG\\\\')
is_LNG = sde_df['Data Source'].str.contains(re_LNG)

re_GIS = re.compile('\\\\Houston\\\\IntlDeepW\\\\MOZAMBIQUE\\\\MOZGIS\\\\')
is_GIS = sde_df['Data Source'].str.contains(re_GIS)

re_GIS_VENDOR = re.compile('\\\\Houston\\\\IntlDeepW\\\\MOZAMBIQUE\\\\MOZGIS\\\\VENDOR\\\\')
is_GIS_VENDOR = sde_df['Data Source'].str.contains(re_GIS_VENDOR)

sde_df[is_LNG]['Source Type'].value_counts()
sde_df[is_GIS]['Source Type'].value_counts()

sde_df_LNG = sde_df[is_LNG]
sde_df_GIS = sde_df[is_GIS]

### vendor data
sde_df[is_LNG | is_GIS_VENDOR]['Source Type'].value_counts()
### apc internal data
sde_df[is_GIS & ~is_GIS_VENDOR]['Source Type'].value_counts()


# define the snapshot paths for replacement
LNG_snapshot_path_ex = '\\\\Houston\\\\Mozambique LNG\\\\~snapshot\\\\jk_keep\\\\'
GIS_snapshot_path_ex = '\\\\Houston\\\\IntlDeepW\\\\MOZAMBIQUE\\\\MOZGIS\\\\~snapshot\\\\jkmigrate_12272017_keep\\\\'

LNG_snapshot_path_NEW = sde_df[is_LNG]['Data Source'].str.replace('\\\\Houston\\\\Mozambique LNG\\\\', LNG_snapshot_path_ex)
GIS_snapshot_path_NEW = sde_df[is_GIS]['Data Source'].str.replace('\\\\Houston\\\\IntlDeepW\\\\MOZAMBIQUE\\\\MOZGIS\\\\', GIS_snapshot_path_ex)

sde_df['Snapshot Path'] = sde_df['Data Source']

sde_df.loc[is_LNG]['Snapshot Path'] = LNG_snapshot_path_NEW
sde_df.loc[is_GIS]['Snapshot Path'] = GIS_snapshot_path_NEW

sde_df.to_excel(r'H:\nbProjects\LDrive2SDE_Data_Path_w_snp-paths_2.xlsx', index=False)

