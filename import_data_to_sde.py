import sys
import os
from arcpy import mapping
from openpyxl import load_workbook
import logging

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

HEADERS = ["Name", "Data Source", "Layer Type", "Verified?", "Loaded?", "SDE Name", "SDE Path",
           "Livelink Link", "Definition Query", "Description"]

DATA_SOURCES = [{
        "LDrv": "\\\\anadarko.com\\world\\SharedData\\Houston\\IntlDeepW\\MOZAMBIQUE\\MOZGIS\\",
        "Source": "MOZGIS"
    }, {
        "LDrv": "L:\\SharedData\\Houston\\IntlDeepW\\MOZAMBIQUE\\MOZGIS\\",
        "Source": "MOZGIS"
    }, {
        "LDrv": "\\\\anadarko.com\\world\\SharedData\\Houston\\Mozambique LNG\\",
        "Source": "LNG"
    }, {
        "LDrv": "L:\\SharedData\\Houston\\Mozambique LNG\\",
        "Source": "LNG"
    }, {
        "LDrv": "\\\\anadarko.com\\expdata\\Houston\\gngdata\\raster_depot",
        "Source": "Raster_Depot"
    }]

DATA_TARGETS = [{
        "Connection": r"C:\Users\kdb086\Documents\ArcGIS\SDE2T_MOZ_GIS.sde",
        "Target": "MOZGIS"
    }, {
        "Connection": r"C:\Users\kdb086\Documents\ArcGIS\SDE2T_MOZ_LNG.sde",
        "Target": "LNG"
    }]


def get_source_type(lDrvPath):
    for cvt in DATA_SOURCES:
        if lDrvPath.find(cvt["LDrv"]) == 0:
            return cvt["Source"]
    return None


def get_target_connection(target):
    for cvt in DATA_TARGETS:
        if cvt["Target"] == target:
            return cvt["Connection"]
    return None


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


def write_to_workbook(wbPath, dsList, sheetName=None):
    wb = Workbook(write_only = True)
    ws1 = wb.active
    ws1.title = "dataSource"

    # headers
    headers_for_update = ["Loaded?"]

    # content
    for r in range(0, len(dsList)):
        for c in dsList[r]:
            # TODO: add style to cell?
            h = headers_for_update.index(c)
            if h > -1:
                ws1.cell(row=r+2, column=h+1, value=dsList[r][headers_for_update[h]])
            else:
                print('Invalid header [%s] in xls [%s]' % (c, mxdPath))

    wb.save(filename = wbPath)
    wb.close()

    del wb


def load_layers_in_xls(wbPath):
    dsList = read_from_workbook(wbPath)
    for ds in dsList:
        if ds["Verified?"] is not None and bool(ds["Verified?"]) == True:
            if ds["SDE Name"] is not None and len(ds["SDE Name"]) > 0:
                if ds["Layer Type"] is not None and ds["Layer Type"] == "FeatureLayer":
                    tgt_conn = get_target_connection(get_source_type(ds["Data Source"]))
                    print('%-60s%s' % (ds["Name"],"load to SDE at %s as %s" % (tgt_conn, ds["SDE Name"])))
                    # TODO: upload the actual data
                    #
                    ds["Loaded?"] = 'YES'
                else:
                    print('%-60s%s' % (ds["Name"],"*** skip non-feature layer"))
                    ds["Loaded?"] = "NON-FEATURE"
            else:
                print('%-60s%s' % (ds["Name"],"*** skip layer with no target SDE"))
                ds["Loaded?"] = "NO TARGET"
        else:
            print('%-60s%s' % (ds["Name"],"*** skip invalid layer"))
            ds["Loaded?"] = "INVALID"

    return dsList


def load_layers_in_folder(xlsFolder):
    for root, dirs, files in os.walk(xlsFolder):
        # walk through all files
        for fname in files:
            if fname.endswith(".xlsx"):
                 xlsPath = os.path.join(root, fname)
                 print('\nThe xlsx file: %s' % xlsPath)
                 dsList = load_layers_in_xls(xlsPath)
                 # TODO: update the Loaded? column in xls
                 #


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("import_data_to_sde xls_folder")
    else:
        load_layers_in_folder(sys.argv[1])
