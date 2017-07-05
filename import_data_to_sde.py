import sys
import os
from arcpy import mapping
from openpyxl import Workbook
from openpyxl import load_workbook
import logging

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

HEADERS = ["Name", "Data Source", "Layer Type", "Verified?", "Loaded?", "SDE Name", "SDE Conn",
           "Livelink Link", "Definition Query", "Description"]

HEADERS_FOR_UPDATE = ["Loaded?", "SDE Name", "SDE Conn"]

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


DATA_CATEGORIES = [{
    "Key": "BIO",
    "Name": "BIOLOGICAL"
}, {
    "Key": "BND",
    "Name": "BOUND"
}, {
    "Key": "ELV",
    "Name": "BATHY"
}, {
    "Key": "ELV",
    "Name": "ELEV"
}, {
    "Key": "ELEV",
    "Name": "ELEV"
}, {
    "Key": "FAU",
    "Name": "FAULTS"
}, {
    "Key": "GEOL",
    "Name": "GEOLOGY"
}, {
    "Key": "GET",
    "Name": "GEOTECH"
}, {
    "Key": "GET",
    "Name": "GEOTECH"
}, {
    "Key": "GPH",
    "Name": "GEOPHYSICS"
}, {
    "Key": "HYD",
    "Name": "HYDRO"
}, {
    "Key": "HYD",
    "Name": "HYDROLOGY"
}, {
    "Key": "INF",
    "Name": "INFRASTRUCT"
}, {
    "Key": "LAND",
    "Name": "LAND"
}, {
    "Key": "LSE",
    "Name": "LEASE"
}, {
    "Key": "REF",
    "Name": "REFERENCE"
}, {
    "Key": "STR",
    "Name": "STRUCT"
}, {
    "Key": "TRN",
    "Name": "TRANS"
}, {
    "Key": "TRAN",
    "Name": "TRANS"
}, {
    "Key": "VEN",
    "Name": "VENDOR"
}, {
    "Key": "WELL",
    "Name": "WELL"
}]

WORD_SHORTHANDS = [{
    "Key": "Archae",
    "Name": "Archaelogical"
}, {
    "Key": "ADRO",
    "Name": "Aerodrome"
}, {
    "Key": "", # omit
    "Name": "Area"
}, {
    "Key": "Lyt",
    "Name": "Layout"
}, {
    "Key": "Bndy",
    "Name": "Boundary"
}, {
    "Key": "Coord",
    "Name": "Coordinate"
}, {
    "Key": "Coord",
    "Name": "Coordinates"
}, {
    "Key": "RefPt",
    "Name": "ReferencePt"
}, {
    "Key": "RefPt",
    "Name": "ReferencePts"
}, {
    "Key": "GeoRef",
    "Name": "Georeferenced"
}, {
    "Key": "Crdr",
    "Name": "Corridor"
}, {
    "Key": "Ctur",
    "Name": "Contour"
}, {
    "Key": "Ctur5m",
    "Name": "Contours5m"
}, {
    "Key": "Ctur1m",
    "Name": "Contours1m"
}, {
    "Key": "5mCtur",
    "Name": "5mContour"
}, {
    "Key": "5mCtur",
    "Name": "5mContours"
}, {
    "Key": "1mCtur",
    "Name": "1mContour"
}, {
    "Key": "1mCtur",
    "Name": "1mContours"
}, {
    "Key": "Cesn",
    "Name": "Concession"
}, {
    "Key": "Const",
    "Name": "Construction"
}, {
    "Key": "Conti",
    "Name": "Continental"
}, {
    "Key": "Charz",
    "Name": "Characterization"
}, {
    "Key": "Ctlg",
    "Name": "Catalog"
}, {
    "Key": "Drg",
    "Name": "Dredge"
}, {
    "Key": "DrgDspo",
    "Name": "DredgeDisposal"
}, {
    "Key": "Fac",
    "Name": "Facilities"
}, {
    "Key": "GeoPhy",
    "Name": "Geophysical"
}, {
    "Key": "GF",
    "Name": "Golfinho"
}, {
    "Key": "Ofld",
    "Name": "Offloading"
}, {
    "Key": "Opn",
    "Name": "Operation"
}, {
    "Key": "Opt",
    "Name": "Option"
}, {
    "Key": "Opt1",
    "Name": "Option1"
}, {
    "Key": "Otln",
    "Name": "Outline"
}, {
    "Key": "RainMntr",
    "Name": "RainMonitoring"
}, {
    "Key": "PB",
    "Name": "PalmaBay"
}, {
    "Key": "Pt",
    "Name": "Point"
}, {
    "Key": "Pt",
    "Name": "Points"
}, {
    "Key": "Ln",
    "Name": "Line"
}, {
    "Key": "Ln",
    "Name": "Lines"
}, {
    "Key": "PR",
    "Name": "Prosperidade"
}, {
    "Key": "GfnoPLN",
    "Name": "GolfinhoPipeline"
}, {
    "Key": "MbaPLN",
    "Name": "MambaPipeline"
}, {
    "Key": "PioCmp",
    "Name": "PioneerCamp"
}, {
    "Key": "PB",
    "Name": "PalmaBay"
}, {
    "Key": "PLN",
    "Name": "Pipeline"
}, {
    "Key": "Espm",
    "Name": "Escarpment"
}, {
    "Key": "MB",
    "Name": "Mamba"
}, {
    "Key": "Prop",
    "Name": "Proposed"
}, {
    "Key": "Onsh",
    "Name": "Onshore"
}, {
    "Key": "Ofsh",
    "Name": "Offshore"
}, {
    "Key": "Rst",
    "Name": "Resettlement"
}, {
    "Key": "Alt",
    "Name": "Alternate"
}, {
    "Key": "Bch",
    "Name": "Beach"
}, {
    "Key": "Dev",
    "Name": "Development"
}, {
    "Key": "Deply",
    "Name": "Deployment"
}, {
    "Key": "Ldg",
    "Name": "Landing"
}, {
    "Key": "Dspo",
    "Name": "Disposal"
}, {
    "Key": "DspoArea",
    "Name": "DisposalArea"
}, {
    "Key": "Pote",
    "Name": "Potential"
}, {
    "Key": "Rd",
    "Name": "Road"
}, {
    "Key": "Rd",
    "Name": "Roads"
}, {
    "Key": "Term",
    "Name": "Terminal"
}, {
    "Key": "Vlg",
    "Name": "Village"
}, {
    "Key": "Rplc",
    "Name": "Replacement"
}, {
    "Key": "RplcLand",
    "Name": "ReplacementLand"
}, {
    "Key": "RplcVlg",
    "Name": "ReplacementVillage"
}, {
    "Key": "VegiClear",
    "Name": "VegetationClearance"
}, {
    "Key": "Moz",
    "Name": "Mozambique"
}, {
    "Key": "Proj",
    "Name": "Project"
}, {
    "Key": "Indu",
    "Name": "Industrial"
}, {
    "Key": "Intl",
    "Name": "International"
}, {
    "Key": "Sec",
    "Name": "Section"
}, {
    "Key": "Stg",
    "Name": "Storage"
}, {
    "Key": "Vtx",
    "Name": "Vertices"
}, {
    "Key": "Vtx",
    "Name": "Vertex"
}, {
    "Key": "Zn",
    "Name": "Zone"
}, {
    "Key": "Zn",
    "Name": "Zones"
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


def guess_target_name(lDrvPath):
    target_source = get_source_type(lDrvPath)
    if target_source == "LNG":
        return None
    elif target_source == "MOZGIS":
        # parse the path
        for ds in DATA_SOURCES:
            if ds["Source"] == "MOZGIS":
                lDrvPath = lDrvPath.replace(ds["LDrv"], "")
        parts = lDrvPath.split("\\")
        category = parts[0]
        dataFormat = parts[1]
        dataName = None
        dataKeys = []
        if category not in ["WORKING"]:
            if dataFormat == 'gdb':
                dataName = parts[-1]
            elif dataFormat == 'shapefiles':
                dataName = os.path.splitext(parts[-1])[0]
                dataName = dataName.replace(' ', '_')
            else:
                return None
            for c in DATA_CATEGORIES:
                if c["Name"] == category:
                    dataKeys.append(c["Key"])
            minKey = None
            keyLen = 0
            for k in dataKeys:
                if dataName.find(k) == 0:
                    return dataName
                if minKey is None or keyLen > len(k):
                    minKey = k
                    keyLen = len(k)
            if minKey is None:
                return dataName
            else:
                return minKey + "_" + dataName

    return None


def shorten_target_name(name, maxLen=30):
    if name is None or len(name) <= maxLen:
        return name
    else:
        nameParts = name.split("_")
        shortNameParts = list(nameParts)
        for p in range(0, len(nameParts)):
            nmPart = nameParts[p].upper()
            for sh in WORD_SHORTHANDS:
                if sh["Name"].upper() == nmPart:
                    shortNameParts[p] = sh["Key"]
                    break
            if len("_".join(shortNameParts)) <= maxLen:
                break
        shortNameParts = [e for e in shortNameParts if len(e) > 0]
        return "_".join(shortNameParts)


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


def update_status_in_workbook(wbPath, dsList, sheetName=None):
    wb = load_workbook(filename = wbPath)
    ws1 = wb["dataSource"]

    # content
    s = 2 # skip the first 2 rows
    for r in range(0, len(dsList)):
        for c in dsList[r]:
            for u in HEADERS_FOR_UPDATE:
                h = HEADERS.index(u)
                ws1.cell(row=r+s+1, column=h+1, value=dsList[r][HEADERS[h]])

    wb.save(filename = wbPath)
    wb.close()

    del wb


def load_layers_in_xls(wbPath, test):
    dsList = read_from_workbook(wbPath)
    for ds in dsList:
        if ds["Loaded?"] not in ["LOADED", 'EXIST']:
            if ds["Layer Type"] is not None and ds["Layer Type"] == "FeatureLayer":
                if ds["Verified?"] is not None and bool(ds["Verified?"]) == True:
                    tgt_conn = get_target_connection(get_source_type(ds["Data Source"]))
                    if tgt_conn is not None:
                        ds["SDE Conn"] = tgt_conn
                        if ds["SDE Name"] is None:
                            ds["SDE Name"] = guess_target_name(ds["Data Source"])
                        if ds["SDE Name"] is not None and len(ds["SDE Name"]) > 0:
                            ds["SDE Name"] = shorten_target_name(ds["SDE Name"])
                            if len(ds["SDE Name"]) > 30:
                                print('%-60s%s' % (ds["Name"],"*** name too long [%s]" % ds["SDE Name"]))
                                ds["Loaded?"] = "NAME TOO LONG"
                            elif arcpy.Exists(tgt_conn + "\\" + ds["SDE Name"]) == True:
                                print('%-60s%s' % (ds["Name"],"*** existing layer"))
                                ds["Loaded?"] = 'EXIST'
                            else:
                                print('%-60s%s' % (ds["Name"],"loading to SDE at %s as %s" % (tgt_conn, ds["SDE Name"])))
                                # upload the actual data
                                if test == "test":
                                    print('%-60s%s' % (" ","^^^ TESTED"))
                                    ds["Loaded?"] = 'TESTED'
                                else:
                                    try:
                                        arcpy.CopyFeatures_management(ds["Data Source"], tgt_conn + "\\" + ds["SDE Name"])
                                        print('%-60s%s' % (" ","^^^ LOADED"))
                                        ds["Loaded?"] = 'LOADED'
                                    except:
                                        print('%-60s%s' % (" ",">>> FAILED"))
                                        ds["Loaded?"] = 'FAILED'
                        else:
                            print('%-60s%s' % (ds["Name"],"*** no target name"))
                            ds["Loaded?"] = "NO TARGET NAME"
                    else:
                        print('%-60s%s' % (ds["Name"],"*** no target SDE"))
                        ds["Loaded?"] = "NO TARGET SDE"
                else:
                    print('%-60s%s' % (ds["Name"],"*** invalid layer"))
                    ds["Loaded?"] = "INVALID"
            else:
                print('%-60s%s' % (ds["Name"],"*** non-feature layer"))
                ds["Loaded?"] = "NON-FEATURE"
        else:
            print('%-60s%s' % (ds["Name"],"*** existing layer"))
            # ds["Loaded?"] = 'EXIST'

    return dsList


def load_layers_in_folder(xlsFolder, test):
    for root, dirs, files in os.walk(xlsFolder):
        # walk through all files
        for fname in files:
            if fname.endswith(".xlsx"):
                # read from a xls file
                xlsPath = os.path.join(root, fname)
                print('\nThe xlsx file: %s' % xlsPath)
                dsList = load_layers_in_xls(xlsPath, test)
                # update status in the xls file
                update_status_in_workbook(xlsPath, dsList)


if __name__ == "__main__":
    if len(sys.argv) < 2 and len(sys.argv) > 3:
        print("import_data_to_sde xls_folder [test]")
    else:
        test = None
        if len(sys.argv) == 3:
            test = sys.argv[2]
        load_layers_in_folder(sys.argv[1], test)

