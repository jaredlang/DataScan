import sys
import os
import re
import xml.etree.ElementTree as ET
from datetime import datetime
from arcpy import mapping
from openpyxl import Workbook
import logging

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

HEADERS = []

DATA_SOURCES = []
DATA_TARGETS = []

LDRV_LL_PATHS = []


def load_config(configFile):
    tree = ET.parse(configFile)
    root = tree.getroot()

    for child in root.iter('header'):
        HEADERS.append(child.attrib['name'])
    # print HEADERS

    for child in root.iter('source'):
        DATA_SOURCES.append({
            'name':child.attrib['name'],
            'Source':child.attrib['name'],
            'LDrv':child.attrib['path']
        })
    # print DATA_SOURCES

    for child in root.iter('target'):
        target = {}
        target['name'] = 'MOZGIS'
        for sc in child:
            target[sc.tag] = sc.text
            DATA_TARGETS.append(target)
    # print DATA_TARGETS

    for src in DATA_SOURCES:
        for tgt in DATA_TARGETS:
            if src['name'] == tgt['name']:
                LDRV_LL_PATHS.append({
                    'Name':src['name'],
                    'LDrv':src['LDrv'],
                    'LL':tgt['livelink']
                })
    # print LDRV_LL_PATHS


def find_Livelink_path(lDrvPath):
    for cvt in LDRV_LL_PATHS:
        if lDrvPath.find(cvt["LDrv"]) == 0:
            return cvt["LL"]
    return None


def verify_layer_dataSource(lyr, lyrType):
    try:
        return arcpy.Exists(lyr)
    except:
        logger.error("failed to open layer [%s]" % fc)
        return False


def write_to_workbook(wbPath, dsList, sheetName=None):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "dataSource"

    # headers
    for c in range(0, len(HEADERS)):
        ws1.cell(row=1, column=c+1, value=HEADERS[c])

    # content
    s = 1 # skip the first 1 row
    for r in range(0, len(dsList)):
        for c in dsList[r]:
            # TODO: add style to cell?
            h = HEADERS.index(c)
            if h > -1:
                ws1.cell(row=r+s+1, column=h+1, value=dsList[r][HEADERS[h]])
            else:
                print('Invalid header [%s] in xls [%s]' % (c, mxdPath))

    wb.save(filename = wbPath)
    wb.close()

    del wb


def scan_layers_in_mxd(mxdPath):
     dsList = []
     # scan the mxd file
     try:
         mxd = mapping.MapDocument(mxdPath)
         lyrs = mapping.ListLayers(mxd)
         dsList.append({"Name": "MXD File", "Data Source": mxdPath})
         lyrTitle = 'Map Layer'
         sourceTitle = 'Data Source'
         print('%-60s%s' % (lyrTitle,sourceTitle))
         print('%-60s%s' % ('-' * len(lyrTitle), '-' * len(sourceTitle)))
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
                     ds = lyr.dataSource
                     print('%-60s%s' % (lyr.name,ds))
                     # get the layer def query
                     defQry = None
                     if lyr.supports("DEFINITIONQUERY") == True:
                         defQry = lyr.definitionQuery
                     # get the layer description
                     dspt = None
                     if lyr.supports("DESCRIPTION") == True:
                         dspt = lyr.description
                     #
                     # verify the layer data source
                     verified = verify_layer_dataSource(ds, lyrType)
                     # get the Livelink path
                     llPath = find_Livelink_path(ds)
                     #
                     dsList.append({"Name": lyr.name, "Data Source": ds, "Layer Type": lyrType,
                                    "Verified?": verified, "Definition Query": defQry, "Description": dspt,
                                    "Livelink Link": llPath})
             except:
                 print('%-60s%s' % (lyr.name,">>> failed to retrieve info " + lyrType))
     except:
         print('Unable to open the mxd file [%s]' % mxdPath)
     finally:
         del mxd

     return dsList


def scan_mxd_in_folder(mxdFolder, xlsFolder, date_filters):

    # parse date filters
    lower_date = None
    upper_date = None
    if date_filters is not None:
        dfs = re.split('<', date_filters)
        if len(dfs) == 2:
            if dfs[0] is not None and len(dfs[0]) > 0:
                lower_date = datetime.datetime.strptime(dfs[0], "%Y/%m/%d")
            if dfs[1] is not None and len(dfs[1]) > 0:
                upper_date = datetime.datetime.strptime(dfs[1], "%Y/%m/%d")

    # walk through all files
    for root, dirs, files in os.walk(mxdFolder):
        for fname in files:
            if fname.endswith(".mxd"):
                mxdPath = os.path.join(root, fname)
                modified_time = datetime.datetime.fromtimestamp(os.stat(mxdPath).st_mtime)
                # check against the filter
                is_filter_met = True
                if lower_date is not None:
                    is_filter_met = modified_time >= lower_date
                if upper_date is not None:
                    is_filter_met = is_filter_met and modified_time <= upper_date
                # work on the mxd file
                if is_filter_met == True:
                    print('\nThe mxd file: %s' % mxdPath)
                    dsList = scan_layers_in_mxd(mxdPath)
                    wbFolder = root.replace(mxdFolder, xlsFolder)
                    if not os.path.exists(wbFolder):
                        os.makedirs(wbFolder)
                    wbFilePath = os.path.join(wbFolder, fname + ".xlsx")
                    print('\nThe xlsx file: %s' % wbFilePath)
                    write_to_workbook(wbFilePath, dsList)


if __name__ == "__main__":
    if len(sys.argv) < 3 and len(sys.argv) > 5:
        print("scan_mxds mxd_folder xls_folder [date_filter ex. 2016/2/19<]  [config_file]")
    else:
        date_filters = None
        if len(sys.argv) == 4:
            date_filters = sys.argv[3]
        config_file = r'H:\MXD_Scan\config.xml'
        if len(sys.argv) == 5:
            config_file = sys.argv[4]
        load_config(config_file)
        scan_mxd_in_folder(sys.argv[1], sys.argv[2], date_filters)

