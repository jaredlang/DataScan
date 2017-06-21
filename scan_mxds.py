import sys
import os
import re
from datetime import datetime
from arcpy import mapping
from openpyxl import Workbook
import logging

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

HEADERS = ["Name", "Data Source", "Layer Type", "Verified?", "Loaded?", "SDE Name", "SDE Conn",
           "Livelink Link", "Definition Query", "Description"]

LDRV_LL_PATHS = [{
    "LDrv": "\\\\anadarko.com\\world\\SharedData\\Houston\\IntlDeepW\\MOZAMBIQUE\\MOZGIS\\",
    "LL": "http://wwprojectstest.anadarko.com/projects/llisapi.dll?func=ll&objId=33642312&objAction=browse&viewType=1"
  }, {
    "LDrv": "L:\\SharedData\\Houston\\IntlDeepW\\MOZAMBIQUE\\MOZGIS\\",
    "LL": "http://wwprojectstest.anadarko.com/projects/llisapi.dll?func=ll&objId=33642312&objAction=browse&viewType=1"
  }, {
    "LDrv": "\\\\anadarko.com\\world\\SharedData\\Houston\\Mozambique LNG\\",
    "LL": "http://wwprojectstest.anadarko.com/projects/llisapi.dll?func=ll&objId=33638511&objAction=browse&viewType=1"
  }, {
    "LDrv": "L:\\SharedData\\Houston\\Mozambique LNG\\",
    "LL": "http://wwprojectstest.anadarko.com/projects/llisapi.dll?func=ll&objId=33638511&objAction=browse&viewType=1"
  }]


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
    if len(sys.argv) < 3 and len(sys.argv) > 4:
        print("scan_mxds mxd_folder xls_folder [date_filter ex. 2016/2/19<]")
    else:
        date_filters = None
        if len(sys.argv) == 4:
            date_filters = sys.argv[3]
        scan_mxd_in_folder(sys.argv[1], sys.argv[2], date_filters)

