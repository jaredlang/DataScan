#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      kdb086
#
# Created:     20/05/2016
# Copyright:   (c) kdb086 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os

def analyze_mxd_scan(src_file, tgt_file, key_value):
    with open(tgt_file, 'w') as tgtf:
        with open(src_file, 'r') as srcf:
            lines = srcf.readlines()
            for ln in lines:
                i = ln.find(key_value)
                while i > -1:
                    tgtf.write(ln[i:].replace('\n', '').strip() + '\n')


if __name__ == '__main__':
    root_folder = r'C:\Users\kdb086\Documents\DataScan'

    analyze_mxd_scan(os.path.join(root_folder, 'listFolderMXD_MOZGIS_arcgis.log'),
        os.path.join(root_folder, 'mxd_input_files.txt'),
        r'Input: "')

    analyze_mxd_scan(os.path.join(root_folder, 'listFolderMXD_MOZGIS_arcgis.log'),
        os.path.join(root_folder, 'xref_to_Mozambique_LNG.txt'),
        r'* SOURCE: "L:\SharedData\Houston\Mozambique LNG')
