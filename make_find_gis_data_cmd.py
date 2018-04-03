import os
import sys
import argparse
from datetime import datetime

SYS_DIRS = ['Thumbs.db']

SCAN_CMD = r"python C:\Users\kdb086\Documents\scan_all\find_gis_data.py "

##SRC_FOLDERS = [{
##    "name": "MOZLNG",
##    "path": r"R:\~snapshot\jkmigrate_12272017_keep"
##}, {
##    "name": "MOZGIS",
##    "path": r"N:\~snapshot\daily.2017-12-28_0010"
##}]

def make_scan_cmd(srcFolder, xlsFolder, srcPrefix):
    cmdlines = []
    # list the first two sub-directories below srcPath
    for dirname in os.listdir(srcFolder):
        if dirname not in SYS_DIRS:
            subdir = os.path.join(srcFolder, dirname)
            for dirname2 in os.listdir(subdir):
                xlsFileName = srcPrefix + "_" + dirname + "_" + dirname2 + ".xlsx"
                cmdlines.append('%s -d "%s" -o "%s"\n' % (SCAN_CMD, os.path.join(subdir, dirname2), os.path.join(xlsFolder, xlsFileName)))

    return cmdlines


def make_batch_scan_cmd(srcFolder, batFile, xlsFolder, srcPrefix):
    cmdlines = make_scan_cmd(srcFolder, xlsFolder, srcPrefix)
    with open(batFile, "w") as bat:
        for scanCmd in cmdlines:
            bat.write(scanCmd)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Make batch commands with given parameters in files')

    parser.add_argument('-s','--src', help='A root directory for making a batch file', required=True)
    parser.add_argument('-b','--bat', help='Command batch file (output)', required=True)
    parser.add_argument('-x','--xls', help='A direcotry for the command output files', required=True)
    parser.add_argument('-a','--action', help='Action Options (update, scan, copy)', required=False, default='scan')
    parser.add_argument('-p','--pfx', help='A prefix in the command for the root directory', required=False, default="MOZ")

    params = parser.parse_args()

    if params.action == "scan":
        make_batch_scan_cmd(params.src, params.bat, params.xls, params.pfx)
        print('###### Completed making data scan batch [%s]'%params.bat)
    else:
        print 'Error: unknown action [%s] for importing' % params.action
