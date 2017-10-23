import os
import sys
import argparse
from datetime import datetime

def find_missed_mxds(mxdFolder, xlsFolder, fext):
    missedMxds = []
    # walk through all files
    for root, dirs, files in os.walk(mxdFolder):
        for fname in files:
            if fname.endswith(".mxd"):
                mxdPath = os.path.join(root, fname)
                modFname = fname
                if not modFname.endswith("." + fext):
                    modFname = modFname + "." + fext
                xlsPath = os.path.join(root.replace(mxdFolder, xlsFolder), modFname)
                if not os.path.exists(xlsPath):
                    mxdModifiedTime = datetime.fromtimestamp(os.stat(mxdPath).st_mtime)
                    missedMxd = {
                        "Path": mxdPath,
                        "ModTime": mxdModifiedTime.strftime('"%Y-%m-%d"')
                    }
                    missedMxds.append(missedMxd)

    if len(missedMxds) == 0:
        print "****** No missed file"
    else:
        missedMxds = sorted(missedMxds, key=lambda missedMxd: missedMxd['ModTime'])
        for missedMxd in missedMxds:
            # print '%-10s  %s' % (missedMxd['ModTime'], missedMxd['Path'])
            print '"' + missedMxd['Path'] + '"'

    return missedMxds


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Compare two files on mxds and list differences into spreadsheets')
    parser.add_argument('-m','--mxd', help='MXD Folder (input)', required=True)
    parser.add_argument('-x','--xls', help='XLS Folder (input)', required=True)
    parser.add_argument('-t','--ext', help='Extra File Extension (mxd, xlsx)', required=False, default='xlsx')

    params = parser.parse_args()

    find_missed_mxds(params.mxd, params.xls, params.ext)

