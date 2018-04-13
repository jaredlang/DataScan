import os
import sys
import re
import argparse
from datetime import datetime

SYS_DIRS = ['Thumbs\.db', ".*\.gdb", ".*\.zip", ".*\.doc", ".*\.docx", ".*\.pdf", ".*\.mxd"]

SCAN_CMD = r"python C:\Users\kdb086\Documents\scan_all\find_gis_data.py "

##SRC_FOLDERS = [{
##    "name": "MOZLNG",
##    "path": r"R:\~snapshot\jkmigrate_12272017_keep"
##}, {
##    "name": "MOZGIS",
##    "path": r"N:\~snapshot\daily.2017-12-28_0010"
##}]


def is_sys_dir(dirname):
    for ptn in SYS_DIRS:
        if re.match(ptn, dirname, re.IGNORECASE) is not None:
            return True
    return False


def make_scan_cmd(srcFolder, xlsFolder, srcType):
    cmdlines = []
    # list the first two sub-directories below srcPath
    for dirname in os.listdir(srcFolder):
        if not is_sys_dir(dirname):
            subdir = os.path.join(srcFolder, dirname)
            if not os.path.isdir(subdir):
                continue
            for dirname2 in os.listdir(subdir):
                if not is_sys_dir(dirname2):
                    subdir2 = os.path.join(subdir, dirname2)
                    if not os.path.isdir(subdir2):
                        continue
                    if srcType == "MOZGIS":
                        # Two-level sub-directories are enough
                        xlsFileName = srcType + "_" + dirname + "_" + dirname2 + ".xlsx"
                        cmdlines.append('%s -d "%s" -o "%s"\n' % (SCAN_CMD, subdir2, os.path.join(xlsFolder, xlsFileName)))
                    elif srcType == "MOZLNG":
                        # Go down one more level for MOZLNG
                        for dirname3 in os.listdir(subdir2):
                            subdir3 = os.path.join(subdir2, dirname3)
                            if not os.path.isdir(subdir3):
                                continue
                            if not is_sys_dir(dirname3):
                                # Two-level sub-directories are enough
                                xlsFileName = srcType + "_" + dirname + "_" + dirname2 + "_" + dirname3 + ".xlsx"
                                cmdlines.append('%s -d "%s" -o "%s"\n' % (SCAN_CMD, subdir3, os.path.join(xlsFolder, xlsFileName)))

    return cmdlines


def make_batch_scan_cmd(srcFolder, batFile, xlsFolder, srcType):
    cmdlines = make_scan_cmd(srcFolder, xlsFolder, srcType)
    with open(batFile, "w") as bat:
        for scanCmd in cmdlines:
            bat.write(scanCmd)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Make batch commands with given parameters in files')

    parser.add_argument('-s','--src', help='A root directory for making a batch file', required=True)
    parser.add_argument('-b','--bat', help='Command batch file (output)', required=True)
    parser.add_argument('-x','--xls', help='A direcotry for the command output files', required=True)
    parser.add_argument('-a','--action', help='Action Options (update, scan, copy)', required=False, default='scan')
    parser.add_argument('-t','--type', help='A prefix in the command for the root directory', required=False, default="MOZLNG")

    params = parser.parse_args()

    if params.action == "scan":
        make_batch_scan_cmd(params.src, params.bat, params.xls, params.type)
        print('###### Completed making data scan batch [%s]'%params.bat)
    else:
        print 'Error: unknown action [%s] for importing' % params.action
