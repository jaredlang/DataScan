call H:\MXD_Scan\batch_scan_load.cmd

call H:\MXD_Scan\batch_scan_load_missed.cmd

python H:\MXD_Scan\find_missed_mxds.py -t xlsx -x H:\MXD_Scan\xlsx\arcgis_files\mxds_scan_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\scan_mxds_missed_NEW.log 2>&1

python H:\MXD_Scan\find_missed_mxds.py -t xlsx -x H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\load_mxds_missed_NEW.log 2>&1

REM call H:\MXD_Scan\compile_data_folders.cmd

call H:\MXD_Scan\batch_update_mxds.cmd

python H:\MXD_Scan\find_missed_mxds.py -t mxd -x H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_missed_files_NEW.txt 2>&1

python H:\MXD_Scan\make_mxd_update_batch.py -b H:\MXD_Scan\batch_update_missed_mxds.cmd -s H:\MXD_Scan\xlsx\arcgis_files\update_mxds_missed_files_NEW.txt

echo MXD(missed) UPDATE script START %date% %time%
call H:\MXD_Scan\batch_update_missed_mxds.cmd > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_missed.log 2>&1
echo MXD(missed) UPDATE script END %date% %time%

python H:\MXD_Scan\find_missed_mxds.py -t mxd -x H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all -m \\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds > H:\MXD_Scan\xlsx\arcgis_files\update_mxds_missed_files_NEW2.txt 2>&1
