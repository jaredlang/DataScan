echo off
echo MXD UPDATE script START %date% %time%

REM echo update mxd of iMap %date% %time%
REM python H:\MXD_Scan\update_mxds.py -x H:\MXD_Scan\xlsx\iMap\mxds_updateDDD -n H:\MXD_Scan\xlsx\iMap\mxds_load -m \\anadarko.com\world\AppsData\Houston\iMaps\Mxds_InWork\MOZ\Portal_20170221 > H:\MXD_Scan\xlsx\iMap\update_mxds_iMap.log 2>&1

echo update mxd of 2013 %date% %time%
python H:\MXD_Scan\update_mxds.py -f "<2014/2/18" -x H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all -n H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_2013.log 2>&1

echo update mxd of 2014 %date% %time%
python H:\MXD_Scan\update_mxds.py -f "2014/2/19<2015/2/18" -x H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all -n H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_2014.log 2>&1

echo update mxd of 2015 %date% %time%
python H:\MXD_Scan\update_mxds.py -f "2015/2/19<2016/2/18" -x H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all -n H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_2015.log 2>&1

echo update mxd of 2016 %date% %time%
python H:\MXD_Scan\update_mxds.py -f "2016/2/19<2017/2/18" -x H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all -n H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_2016.log 2>&1

echo update mxd of 2017 %date% %time%
python H:\MXD_Scan\update_mxds.py -f "2017/2/19<" -x H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all -n H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_2017.log 2>&1

echo MXD UPDATE script END %date% %time%
echo on
