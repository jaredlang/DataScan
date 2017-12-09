import os
import sys
import argparse
from datetime import datetime

CMD = "python H:\MXD_Scan\update_mxds.py -a single "
SRC_FOLDER = r"\\anadarko.com\world\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS\arcgis_files\mxds"
TGT_FOLDER = r"H:\MXD_Scan\xlsx\arcgis_files\mxds_updateDDD_all"
XTS_FOLDER = r"H:\MXD_Scan\xlsx\arcgis_files\mxds_load_all"

def make_batch_cmd(exeCmd, srcFile, tgtFile, xtsFile, cmdFile):

    with open(cmdFile, "w") as cmd:
        with open(srcFile, "r") as src:
            with open(tgtFile, "r") as tgt:
                with open(xtsFile, "r") as xts:
                    srcLines = src.readlines()
                    for srcLine in srcLines:
                        tgtLine = tgt.readline()
                        xtsLine = xts.readline()
                        cmd.write(" ".join([exeCmd,
                            srcLine.replace('\n', ''),
                            tgtLine.replace('\n', ''),
                            xtsLine.replace('\n', '')]) + "\n")


def make_update_cmd(srcPath):
    tgtPath = srcPath.replace(SRC_FOLDER, TGT_FOLDER)
    xtsPath = srcPath.replace(SRC_FOLDER, XTS_FOLDER)

    return "%s -m %s -n %s -x %s\n" % (CMD, srcPath, tgtPath, xtsPath)


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

    params = parser.parse_args()

    make_batch_update_cmd(params.src, params.bat)
    print('###### Completed making mxd update batch [%s]'%params.bat)

'''
    parser.add_argument('-s','--src', help='Source parameter file (input)', required=True)
    parser.add_argument('-t','--tgt', help='Target parameter file (input)', required=True)
    parser.add_argument('-x','--xts', help='Option parameter file (input)', required=True)
    parser.add_argument('-c','--cmd', help='Command batch file (output)', required=True)

    params = parser.parse_args()

    make_batch_cmd(CMD, params.src, params.tgt, params.xts, params.cmd)
'''
