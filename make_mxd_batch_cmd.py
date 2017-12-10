import os
import sys
import argparse
from datetime import datetime

SCAN_CMD = r"python H:\MXD_Scan\scan_mxds.py -a single "
UPDATE_CMD = r"python H:\MXD_Scan\update_mxds.py -a single "
COPY_CMD = "copy /B /Y "

SRC_FOLDER = r"\\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds"
TGT_FOLDER = r"H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all"
XTS_FOLDER = r"H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all"

XLS_FOLDER = r"H:\MXD_Scan\xlsx\arcgis_files\mxds_scan_missed"


def make_copy_cmd(srcPath):
    tgtPath = srcPath.replace(SRC_FOLDER, TGT_FOLDER)

    return "%s %s %s\n" % (COPY_CMD, srcPath, tgtPath)


def make_batch_copy_cmd(srcFile, batFile):

    with open(batFile, "w") as bat:
        with open(srcFile, "r") as src:
            srcPaths = src.readlines()
            for srcPath in srcPaths:
                copyCmd = make_copy_cmd(srcPath.replace('\n', ''))
                bat.write(copyCmd)


def make_scan_cmd(srcPath):
    return "%s -m %s -x %s -s %s\n" % (SCAN_CMD, SRC_FOLDER, XLS_FOLDER, srcPath)


def make_batch_scan_cmd(srcFile, batFile):

    with open(batFile, "w") as bat:
        with open(srcFile, "r") as src:
            srcPaths = src.readlines()
            for srcPath in srcPaths:
                scanCmd = make_scan_cmd(srcPath.replace('\n', ''))
                bat.write(scanCmd)


def make_update_cmd(srcPath):
    tgtPath = srcPath.replace(SRC_FOLDER, TGT_FOLDER)
    xtsPath = srcPath.replace(SRC_FOLDER, XTS_FOLDER)
    hasQuote = xtsPath[-1] in ['"', "'"]
    if hasQuote:
        xtsPath = xtsPath[:-1] + ".xlsx" + xtsPath[-1]
    else:
        xtsPath = xtsPath + ".xlsx"

    return "%s -m %s -n %s -x %s\n" % (UPDATE_CMD, srcPath, tgtPath, xtsPath)


def make_batch_update_cmd(srcFile, batFile):

    with open(batFile, "w") as bat:
        with open(srcFile, "r") as src:
            srcPaths = src.readlines()
            for srcPath in srcPaths:
                updateCmd = make_update_cmd(srcPath.replace('\n', ''))
                bat.write(updateCmd)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Make batch commands with given parameters in files')

    parser.add_argument('-s','--src', help='A file containing a list of mxd file paths', required=True)
    parser.add_argument('-b','--bat', help='Command batch file (output)', required=True)
    parser.add_argument('-a','--action', help='Action Options (update, scan, copy)', required=False, default='update')

    params = parser.parse_args()

    if params.action == "update":
        make_batch_update_cmd(params.src, params.bat)
        print('###### Completed making mxd update batch [%s]'%params.bat)
    elif params.action == "scan":
        make_batch_scan_cmd(params.src, params.bat)
        print('###### Completed making mxd scan batch [%s]'%params.bat)
    elif params.action == "copy":
        make_batch_copy_cmd(params.src, params.bat)
        print('###### Completed making mxd copy batch [%s]'%params.bat)
    else:
        print 'Error: unknown action [%s] for importing' % params.action
