echo off
echo FIND REF_DATA_FOLDER script START %date% %time%

python.exe H:\MXD_Scan\compile_data_folders.py -x H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all -o H:\MXD_Scan\xlsx\arcgis_files\REF_data_folders.xlsx > H:\MXD_Scan\xlsx\arcgis_files\compile_data_folders.log 2>&1

echo FIND REF_DATA_FOLDER script END %date% %time%
echo on
