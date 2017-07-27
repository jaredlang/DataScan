echo off
echo batch script START %date% %time%

echo scan mxd of 2014 %date% %time%
python H:\MXD_Scan\scan_mxds.py -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds -x H:\MXD_Scan\xlsx\arcgis_files\mxds_t14 -f "2014/2/19<2014/4/18" > H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_2014.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds_t14\* H:\MXD_Scan\xlsx\arcgis_files\mxds_bak_2014\ /s /e /y

echo load data of 2014 %date% %time%
python H:\MXD_Scan\import_data_to_sde.py -x H:\MXD_Scan\xlsx\arcgis_files\mxds_t14 > H:\MXD_Scan\xlsx\arcgis_files\import_data_to_sde_2014.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds_t14\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_2014\ /s /e /y

rmdir /s /q H:\MXD_Scan\xlsx\arcgis_files\mxds_t14

echo batch script END %date% %time%
echo on
