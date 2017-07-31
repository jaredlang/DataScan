import sys
import os
import argparse
import tempfile
import xml.etree.ElementTree as ET
from arcpy import mapping
from openpyxl import Workbook
from openpyxl import load_workbook
import logging

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

HEADERS = []

def load_config(configFile):
    tree = ET.parse(configFile)
    root = tree.getroot()

    for child in root.iter('header'):
        HEADERS.append(child.attrib['name'])

    del root
    del tree


def read_from_workbook(wbPath, sheetName=None):
    wb = load_workbook(filename = wbPath, read_only=True)
    ws = wb["dataSource"]

    dsList = []
    # skip the first 2 rows
    r = 3
    hdrs = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    while ws["A"+str(r)].value is not None:
        dsRecord = {}
        for c in range(0, len(HEADERS)):
            dsRecord[HEADERS[c]] = ws[hdrs[c]+str(r)].value
        dsList.append(dsRecord)
        r = r + 1

    wb.close()

    del wb

    return dsList


def update_mxd_ds(mxdPath, xlsPath, newMxdPath):
    dsList = read_from_workbook(xlsPath)
    try:
        mxd = arcpy.mapping.MapDocument(mxdPath)
        lyrs = arcpy.mapping.ListLayers(mxd)
        for lyr in lyrs:
             if lyr.isFeatureLayer == True:
                 lyrType = "FeatureLayer"
             elif lyr.isRasterLayer == True:
                 lyrType = "RasterLayer"
             elif lyr.isServiceLayer == True:
                 lyrType = "ServiceLayer"
             elif lyr.isGroupLayer == True:
                 lyrType = "GroupLayer"
             else:
                 lyrType = "UnknownLayer"
             try:
                 if lyrType in ["GroupLayer", "ServiceLayer", "UnknownLayer"]:
                     print('%-60s%s' % (lyr.name,"*** skip " + lyrType))
                 else:
                     # get the layer data source
                     lyrDS = lyr.dataSource
                     for ds in dsList:
                        if ds["Data Source"] == lyrDS and bool(ds["Verified?"]) == True and ds["Loaded?"] in ["LOADED", 'EXIST']:
                            if ds["Layer Type"] == "RasterLayer":
                                wsType = "RASTER_WORKSPACE"
                            elif ds["Layer Type"] == "FeatureLayer":
                                wsType = "SDE_WORKSPACE"
                            else:
                                wsType = "NONE"
                            try:
                                # update the layer data source
                                lyr.replaceDataSource(ds["SDE Conn"], wsType, ds["SDE Name"], True)
                            except:
                                print('%-60s%s' % (lyr.name,">>> failed to update data source " + lyrType))
             except:
                 print('%-60s%s' % (lyr.name,">>> failed to retrieve info " + lyrType))
        # save the updated mxd file
        mxd.saveACopy(newMxdPath)
    except:
        print('Unable to open the mxd file [%s]' % mxdPath)
    finally:
        del mxd


def update_mxds(mxd_folder, xls_folder, new_mxd_folder):
    # walk through the mxd files
    for mxdRoot, mxdDirs, mxdFiles in os.walk(mxd_folder):
        for mxdFile in mxdFiles:
            if mxdFile.endswith(".mxd"):
                mxdPath = os.path.join(mxdRoot, mxdFile)
                print('\nThe mxd file: %s' % mxdPath)

                newMxdPath = os.path.join(new_mxd_folder, mxdPath.replace(mxd_folder, new_mxd_folder))
                newMxdFolder = os.path.dirname(newMxdPath)
                if not os.path.exists(newMxdFolder):
                        os.makedirs(newMxdFolder)

                xlsPath = os.path.join(xls_folder, mxdPath.replace(mxd_folder, xls_folder))
                xlsPath = xlsPath + ".xlsx"

                if not os.path.exists(xlsPath):
                    print('The xls file not exist: %s' % xlsPath)
                else:
                    print('Updating data sources ...')
                    update_mxd_ds(mxdPath, xlsPath, newMxdPath)
                    print("**** completed")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Update mxds with data sources specified in spreadsheets')
    parser.add_argument('-m','--mxd', help='MXD Folder (input)', required=True)
    parser.add_argument('-x','--xls', help='XLS Folder (input)', required=True)
    parser.add_argument('-n','--newMxd', help='New MXD Folder (output)', required=True)
    parser.add_argument('-a','--action', help='Action Options (batch, single)', required=False, default='batch')
    parser.add_argument('-c','--cfg', help='Config File', required=False, default=r'H:\MXD_Scan\config.xml')

    params = parser.parse_args()

    if params.cfg is not None:
        load_config(params.cfg)

    if params.action == 'batch':
        update_mxds(params.mxd, params.xls, params.newMxd)
    elif params.action == 'single':
        print "[%s] Not Implemented Yet" % params.action
    else:
        print 'Error: unknown action [%s] for scanning' % params.action

