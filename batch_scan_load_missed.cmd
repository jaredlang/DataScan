echo off
echo SCAN+LOAD (Missed) script START %date% %time%

rmdir /s /q H:\MXD_Scan\xlsx\arcgis_files\mxds

echo scan mxd of Missed %date% %time%
call H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_missed.cmd > H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_missed.log 2>&1

xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_scan_missed\ /c /e /y /q
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds_scan_missed\* H:\MXD_Scan\xlsx\arcgis_files\mxds_scan_all\ /c /e /y /q

echo load data of Missed %date% %time%
python H:\MXD_Scan\import_data_to_sde.py -x H:\MXD_Scan\xlsx\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\import_data_to_sde_missed.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_missed\ /c /e /y /q
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds_load_missed\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all\ /c /e /y /q

rmdir /s /q H:\MXD_Scan\xlsx\arcgis_files\mxds

echo SCAN+LOAD (Missed) script END %date% %time%
echo on
