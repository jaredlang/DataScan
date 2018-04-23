
def convert_to_snapshot(ds_path):
	LNG_source_path = u'\\Houston\\Mozambique LNG\\'
	GIS_source_path = u'\\Houston\\IntlDeepW\\MOZAMBIQUE\\MOZGIS\\'

	LNG_snapshot_path = u'\\Houston\\Mozambique LNG\\~snapshot\\jk_keep\\'
	GIS_snapshot_path = u'\\Houston\\IntlDeepW\\MOZAMBIQUE\\MOZGIS\\~snapshot\\jkmigrate_12272017_keep\\'

	snp_path = None
	if ds_path.find(GIS_source_path) > -1:
		snp_path = ds_path.replace(GIS_source_path, GIS_snapshot_path)
	elif ds_path.find(LNG_source_path) > -1:
		snp_path = ds_path.replace(LNG_source_path, LNG_snapshot_path)

	return snp_path


import pandas as pd

sde_xlsx = pd.ExcelFile(r'H:\nbProjects\LDrive2SDE_Data_Path.xlsx')
sde_df = sde_xlsx.parse(sde_xlsx.sheet_names[0], header=0, index_col=None)

sde_df['Snapshot Path'] = sde_df['Data Source'].apply(convert_to_snapshot)

sde_df.to_excel(r'H:\nbProjects\LDrive2SDE_Data_Path_w_snp-paths_2.xlsx', index=False)
