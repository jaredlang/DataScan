import sys
import os
import tempfile
import shutil
import xml.etree.ElementTree as ET
from arcpy import mapping
from openpyxl import Workbook
from openpyxl import load_workbook
import logging

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

AGS_HOME = arcpy.GetInstallInfo("Desktop")["InstallDir"]
METADATA_TRANSLATOR = os.path.join(AGS_HOME, r'Metadata/Translator/ARCGIS2FGDC.xml')

HEADERS = []
HEADERS_FOR_UPDATE = []

DATA_SOURCES = []
DATA_TARGETS = []

DATA_CATEGORIES = []
WORD_SHORTHANDS = []

def load_config(configFile):
    tree = ET.parse(configFile)
    root = tree.getroot()

    for child in root.iter('header'):
        HEADERS.append(child.attrib['name'])
        if 'update' in child.attrib.keys() and child.attrib['update'] == 'Y':
            HEADERS_FOR_UPDATE.append(child.attrib['name'])
    # print HEADERS
    # print HEADERS_FOR_UPDATE

    for child in root.iter('source'):
        DATA_SOURCES.append({
            'name':child.attrib['name'],
            'Source':child.attrib['name'],
            'LDrv':child.attrib['path']
        })
    # print DATA_SOURCES

    for child in root.iter('target'):
        target = {}
        target['name'] = child.attrib["name"]
        for sc in child:
            target[sc.tag] = sc.text
        DATA_TARGETS.append(target)
    # print DATA_TARGETS

    for child in root.iter('dataCategory'):
        DATA_CATEGORIES.append({
            'Key':child.attrib['key'],
            'Name':child.attrib['name']
        })
    # print DATA_CATEGORIES

    for child in root.iter('shortHand'):
        WORD_SHORTHANDS.append({
            'Key':child.attrib['key'],
            'Name':child.attrib['name'],
            'Type':child.attrib['type']
        })
    # print WORD_SHORTHANDS

    del root
    del tree

def get_source_type(lDrvPath):
    for cvt in DATA_SOURCES:
        if lDrvPath.find(cvt["LDrv"]) == 0:
            return cvt["Source"]
    return None


def get_sde_connection(target):
    for cvt in DATA_TARGETS:
        if cvt["name"] == target:
            return cvt["sdeConn"]
    return None


def get_raster_connection(target):
    for cvt in DATA_TARGETS:
        if cvt["name"] == target:
            return cvt["rasterDepot"]
    return None


def guess_target_name(lDrvPath):
    target_source = get_source_type(lDrvPath)
    if target_source == "LNG":
        # TODO: come up with a naming convention for the LNG vendor data
        return None
    elif target_source == "MOZGIS":
        # parse the path
        for src in DATA_SOURCES:
            if src["Source"] == "MOZGIS":
                lDrvPath = lDrvPath.replace(src["LDrv"], "")
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


def update_sde_metadata(sdeFC, props, srcType):

    TEMP_DIR = tempfile.gettempdir()
    metadataFile = os.path.join(TEMP_DIR, os.path.basename(sdeFC) + '-metadata.xml')
    if srcType == "MOZGIS":
        migrationText = " *** Migrated from the L Drive (%s)" % props["Data Source"]

    if os.path.exists(metadataFile):
        os.remove(metadataFile)

    # A- export the medata from SDE feature class
    # print 'exporting the metadata of %s to %s' % (sdeFC, metadataFile)
    arcpy.ExportMetadata_conversion(
    	Source_Metadata=sdeFC,
    	Translator=METADATA_TRANSLATOR,
    	Output_File=metadataFile
    )

    # B- modify metadata
    # print 'modifying the metadata file [%s]' % (metadataFile)
    tree = ET.parse(metadataFile)
    root = tree.getroot()
    idinfo = root.find('idinfo')
    dspt = idinfo.find('descript')
    # B1- add the element
    if dspt is None:
        dspt = ET.SubElement(idinfo, 'descript')
        ET.SubElement(dspt, 'abstract')
    else:
        abstract = dspt.find('abstract')
        if abstract is None:
            ET.SubElement(dspt, 'abstract')
    # B2- modify the element text
    abstract = dspt.find('abstract')
    if abstract.text is None:
        abstract.text = migrationText
    elif abstract.text.find(migrationText) == -1:
        abstract.text = abstract.text + migrationText

    tree.write(metadataFile)

    # C- import the modified metadata back to SDE feature class
    # print 'importing the metadata file [%s] to %s' % (metadataFile, sdeFC)
    arcpy.ImportMetadata_conversion(
    	Source_Metadata=metadataFile,
    	Import_Type="FROM_FGDC",
    	Target_Metadata=sdeFC,
    	Enable_automatic_updates="ENABLED"
    )

    # print 'The metadata of %s is updated' % sdeFC


def load_layers_in_xls(wbPath, test):
    dsList = read_from_workbook(wbPath)
    for ds in dsList:
        srcType = get_source_type(ds["Data Source"])
        if srcType is not None:
            if ds["Loaded?"] not in ["LOADED", 'EXIST']:
                if ds["Verified?"] is not None and bool(ds["Verified?"]) == True:
                    if ds["Layer Type"] is not None and ds["Layer Type"] == "RasterLayer":
                        # find out the target workspace
                        tgt_conn = get_raster_connection(srcType)
                        if tgt_conn is not None:
                            tgt_workspace = ds["Data Source"]
                            for src in DATA_SOURCES:
                                tgt_workspace = tgt_workspace.replace(src["LDrv"], tgt_conn)
                            ds["SDE Conn"] = os.path.dirname(tgt_workspace)
                            ds["SDE Name"] = os.path.basename(tgt_workspace)
                            if os.path.exists(tgt_workspace) or arcpy.Exists(tgt_workspace):
                                print('%-60s%s' % (ds["Name"],"*** existing layer"))
                                ds["Loaded?"] = 'EXIST'
                            else:
                                # copy data to the target workspace
                                print('%-60s%s' % (ds["Name"],"loading to Raster Depot as %s" % tgt_workspace))
                                if test == "test":
                                    print('%-60s%s' % (" ","^^^ TESTED"))
                                    ds["Loaded?"] = 'TESTED'
                                else:
                                    if ds["SDE Conn"].lower().endswith('.gdb') == True:
                                        # copy the whole GDB
                                        if os.path.exists(ds["SDE Conn"]):
                                            shutil.rmtree(ds["SDE Conn"])
                                        # arcpy.Copy_management(os.path.dirname(ds["Data Source"]), ds["SDE Conn"])
                                        shutil.copytree(os.path.dirname(ds["Data Source"]), ds["SDE Conn"])
                                    else:
                                        try:
                                            if not os.path.exists(ds["SDE Conn"]):
                                                os.makedirs(ds["SDE Conn"])
                                            # copy the raster file set
                                            arcpy.CopyRaster_management(in_raster=ds["Data Source"], out_rasterdataset=tgt_workspace)
                                            print('%-60s%s' % (" ","^^^ LOADED"))
                                            ds["Loaded?"] = 'LOADED'
                                        except:
                                            print('%-60s%s' % (" ",">>> FAILED to load data"))
                                            ds["Loaded?"] = 'FAILED'
                        else:
                            print('%-60s%s' % (ds["Name"],"*** no target store"))
                            ds["Loaded?"] = "NO TARGET STORE"

                    elif ds["Layer Type"] is not None and ds["Layer Type"] == "FeatureLayer":
                        tgt_conn = get_sde_connection(srcType)
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

                                            print('%-60s%s' % (ds["Name"],"updating SDE metadata of %s\\%s" % (tgt_conn, ds["SDE Name"])))
                                            update_sde_metadata(tgt_conn + "\\" + ds["SDE Name"], ds, srcType)
                                            print('%-60s%s' % (" ","^^^ METADATA UPDATED"))
                                        except:
                                            print('%-60s%s' % (" ",">>> FAILED to load data"))
                                            ds["Loaded?"] = 'FAILED'
                            else:
                                print('%-60s%s' % (ds["Name"],"*** no target name"))
                                ds["Loaded?"] = "NO TARGET NAME"
                        else:
                            print('%-60s%s' % (ds["Name"],"*** no target store"))
                            ds["Loaded?"] = "NO TARGET STORE"
                    else:
                        print('%-60s%s' % (ds["Name"],"*** unknown layer type"))
                        ds["Loaded?"] = "UNKNOWN LAYER"
                else:
                    print('%-60s%s' % (ds["Name"],"*** invalid layer"))
                    ds["Loaded?"] = "INVALID"
            else:
                print('%-60s%s' % (ds["Name"],"*** existing layer"))
                # ds["Loaded?"] = 'EXIST'
        else:
            print('%-60s%s' % (ds["Name"],"*** out of scope"))
            ds["Loaded?"] = "NOT IN SCOPE"

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
    if len(sys.argv) < 2 and len(sys.argv) > 4:
        print("import_data_to_sde xls_folder [test] [config_file]")
    else:
        test = None
        if len(sys.argv) == 3:
            test = sys.argv[2]
        config_file = r'H:\MXD_Scan\config.xml'
        if len(sys.argv) == 4:
            config_file = sys.argv[3]
        load_config(config_file)
        load_layers_in_folder(sys.argv[1], test)

