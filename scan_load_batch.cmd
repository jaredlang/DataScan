echo off
echo batch script START %date% %time%

echo scan mxd of 2016 %date% %time%
python scan_mxds.py -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds -x H:\MXD_Scan\xlsx\arcgis_files\mxds -f "2016/2/19<" > H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_2016.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_bak_2016\ /s /e /y

echo load data of 2016 %date% %time%
python import_data_to_sde.py -x H:\MXD_Scan\xlsx\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\import_data_to_sde_2016.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_2016\ /s /e /y

rmdir /s /q H:\MXD_Scan\xlsx\arcgis_files\mxds

echo scan mxd of 2015 %date% %time%
python scan_mxds.py -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds -x H:\MXD_Scan\xlsx\arcgis_files\mxds -f "2015/2/19<2016/2/18" > H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_2015.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_bak_2015\ /s /e /y

echo load data of 2015 %date% %time%
python import_data_to_sde.py -x H:\MXD_Scan\xlsx\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\import_data_to_sde_2015.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_2015\ /s /e /y

rmdir /s /q H:\MXD_Scan\xlsx\arcgis_files\mxds

echo scan mxd of 2014 %date% %time%
python scan_mxds.py -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds -x H:\MXD_Scan\xlsx\arcgis_files\mxds -f "2014/2/19<2015/2/18" > H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_2014.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_bak_2014\ /s /e /y

echo load data of 2014 %date% %time%
python import_data_to_sde.py -x H:\MXD_Scan\xlsx\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\import_data_to_sde_2014.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_2014\ /s /e /y

rmdir /s /q H:\MXD_Scan\xlsx\arcgis_files\mxds

echo scan mxd of 2013 %date% %time%
python scan_mxds.py -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds -x H:\MXD_Scan\xlsx\arcgis_files\mxds -f "<2014/2/18" > H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_2013.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_bak_2013\ /s /e /y

echo load data of 2013 %date% %time%
python import_data_to_sde.py -x H:\MXD_Scan\xlsx\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\import_data_to_sde_2013.log 2>&1
xcopy H:\MXD_Scan\xlsx\arcgis_files\mxds\* H:\MXD_Scan\xlsx\arcgis_files\mxds_load_2013\ /s /e /y

rmdir /s /q H:\MXD_Scan\xlsx\arcgis_files\mxds

echo batch script END %date% %time%
echo on
