echo off
echo SCAN+LOAD iMap script START %date% %time%

echo scan mxd of iMap %date% %time%
python H:\MXD_Scan\scan_mxds.py -m \\anadarko.com\world\AppsData\Houston\iMaps\Mxds_InWork\MOZ\Portal_20170221 -x H:\MXD_Scan\xlsx\iMap\mxds > H:\MXD_Scan\xlsx\iMap\scan_mxds_iMap.log 2>&1
python H:\MXD_Scan\merge_iMap_data.py -x H:\MXD_Scan\xlsx\iMap\mxds -m H:\MXD_Scan\xlsx\iMap\IMaps_DataSource_Latest_WORKING.xlsx > H:\MXD_Scan\xlsx\iMap\merge_iMap_data.log 2>&1
xcopy H:\MXD_Scan\xlsx\iMap\mxds\* H:\MXD_Scan\xlsx\iMap\mxds_bak\ /s /e /y

echo load data of iMap %date% %time%
python H:\MXD_Scan\import_data_to_sde.py -x H:\MXD_Scan\xlsx\iMap\mxds > H:\MXD_Scan\xlsx\iMap\import_data_to_sde_iMap.log 2>&1
xcopy H:\MXD_Scan\xlsx\iMap\mxds\* H:\MXD_Scan\xlsx\iMap\mxds_load\ /s /e /y
xcopy H:\MXD_Scan\xlsx\iMap\mxds_load\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all\ /s /e /y

rmdir /s /q H:\MXD_Scan\xlsx\iMap\mxds

echo SCAN+LOAD iMap script END %date% %time%
echo on
