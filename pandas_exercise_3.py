import sys
import os
import argparse
import pandas as pd
from lxml import etree as ET

XLS_TAB_NAME = "dataSource"
XLS_HEADERS = []

NETWORK_DRIVES = []

DATA_SOURCES = {}

XREF_TAB_NAME = "xRef"
LL_OUTPUT_HEADERS = ["Source Type", "Data Source", "Livelink Node Id"]
SDE_OUTPUT_HEADERS = ["Source Type", "Data Source", "SDE Path"]


def get_alias_path(lDrvPath):
    for nd in NETWORK_DRIVES:
        if lDrvPath.find(nd["Path"]) == 0:
            return lDrvPath.replace(nd["Path"], nd["Drive"])
        if lDrvPath.find(nd["Drive"]) == 0:
            return lDrvPath.replace(nd["Drive"], nd["Path"])
    return lDrvPath


def get_unc_path(lDrvPath):
    for nd in NETWORK_DRIVES:
        if lDrvPath.find(nd["Drive"]) == 0:
            return lDrvPath.replace(nd["Drive"], nd["Path"])
    return lDrvPath


def get_mapped_path(lDrvPath):
    for nd in NETWORK_DRIVES:
        if lDrvPath.find(nd["Path"]) == 0:
            return lDrvPath.replace(nd["Path"], nd["Drive"])
    return lDrvPath


def load_config(configFile):
    tree = ET.parse(configFile)
    root = tree.getroot()

    for child in root.iter('xlsTab'):
        # only one tab
        XLS_TAB_NAME = child.attrib['name']

    for child in root.iter('xlsHeader'):
        XLS_HEADERS.append(child.attrib['name'])
    # print XLS_HEADERS

    for child in root.iter('networkDrive'):
        NETWORK_DRIVES.append({
            'Drive': child.attrib['drive'],
            'Path': child.attrib['path']
        })
    # print NETWORK_DRIVES

    for child in root.iter('source'):
        nm = child.attrib['name']
        if nm not in DATA_SOURCES.keys():
            DATA_SOURCES[nm] = []
        for sc in child:
            DATA_SOURCES[nm].append({
                'LDrv':sc.attrib['path'],
                'Mode':sc.tag
            })
            DATA_SOURCES[nm].append({
                'LDrv':get_alias_path(sc.attrib['path']),
                'Mode':sc.tag
            })
    # print DATA_SOURCES


def get_source_type(lDrvPath):
    for k in DATA_SOURCES.keys():
        excluded = False
        ds = DATA_SOURCES[k]
        for cvt in [s for s in ds if s["Mode"] == "exclude"]:
            if lDrvPath.find(cvt["LDrv"]) == 0:
                excluded = True
                break
        if not excluded:
            for cvt in [s for s in ds if s["Mode"] == "add"]:
                if lDrvPath.find(cvt["LDrv"]) == 0:
                    return k
    return None


def parse_data_folder(dataSource):
    parts = dataSource.split('\\')
    idx = -1
    for p in list(reversed(range(0, len(parts)))):
        if parts[p].find('.') > -1:
            idx = p
            break
    return '\\'.join(parts[:idx])


def compile_data_folders_For_Livelink(xlsFolder, xlsOutput):
    df_folders_all = None
    for root, dirs, files in os.walk(xlsFolder):
        # walk through all files
        for fname in files:
            if fname.endswith(".xlsx") and not fname.startswith('~$'):
                # read from a xls file
                xlsPath = os.path.join(root, fname)
                print('xlsx file: %s' % xlsPath)
                try:
                    data_xlsx = pd.ExcelFile(xlsPath)
                    df = data_xlsx.parse(XLS_TAB_NAME, header=0, index_col=None)
                    df_1 = df[1:]
                    is_verified = df_1["Verified?"] is not None and df_1["Verified?"] > 0.0
                    is_featureLayer = df_1["Layer Type"] is not None and df_1["Layer Type"] == "FeatureLayer"
                    is_loaded = df_1["Loaded?"] is not None and df_1["Loaded?"].isin(["LOADED", 'EXIST'])
                    df_folders = df_1[is_verified & is_featureLayer & is_loaded]["Data Source"].apply(parse_data_folder)
                    if df_folders_all is None:
                        df_folders_all = df_folders
                    else:
                        df_folders_all.append(df_folders)
                except:
                    print "#### Failed to process the xls file: %s" % xlsPath

    df_folders_unique = pd.DataFrame(columns=LL_OUTPUT_HEADERS)
    df_folders_unique["Data Source"] = df_folders_all.unique()
    df_folders_unique["Source Type"] = df_folders_unique["Data Source"].apply(get_source_type)

    df_folders_unique.to_excel(xlsOutput, sheet_name=XREF_TAB_NAME, index=False)


def compile_data_sources_For_SDE(xlsFolder, xlsOutput):
    df_folders_all = None
    for root, dirs, files in os.walk(xlsFolder):
        # walk through all files
        for fname in files:
            if fname.endswith(".xlsx") and not fname.startswith('~$'):
                # read from a xls file
                xlsPath = os.path.join(root, fname)
                print('xlsx file: %s' % xlsPath)
                try:
                    data_xlsx = pd.ExcelFile(xlsPath)
                    df = data_xlsx.parse(XLS_TAB_NAME, header=0, index_col=None)
                    df_1 = df[1:]
                    df_1["SDE Path"] = df_1["SDE Conn"] + df_1["SDE Name"]
                    is_verified = df_1["Verified?"] > 0.0
                    is_featureLayer = df_1["Layer Type"] == "FeatureLayer"
                    is_loaded = df_1["Loaded?"].isin(["LOADED", 'EXIST'])
                    df_paths = df_1[is_verified & is_featureLayer & is_loaded]["Data Source", "SDE Path"]
                    if df_paths_all is None:
                        df_paths_all = df_paths
                    else:
                        df_paths_all = pd.merge(df_paths_all, df_paths)
                except:
                    print "#### Failed to process the xls file: %s" % xlsPath

    df_paths_unqiue = pd.DataFrame(columns=["Source Type", "Data Source", "SDE Path"])
    df_paths_unique["Data Source"] = df_paths_all["Data Source"].unique()
    df_paths_unique["Source Type"] = df_folders_unique["Data Source"].apply(get_source_type)
    df_paths_unique["SDE Path"] = df_folders_unique["SDE Conn"] + df_folders_unqiue["SDE Name"]

    df_folders_unqiue.to_excel(xlsOutput, sheet_name=XREF_TAB_NAME, index=False)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Compile all data folders or sources into one spreadsheet')
    parser.add_argument('-x','--xls', help='XLS Folder (input)', required=True)
    parser.add_argument('-o','--dsf', help='XLS File (output)', required=True)
    parser.add_argument('-t','--tgt', help="Target Options (Livelink, SDE)", required=True)
    parser.add_argument('-c','--cfg', help='Config File', required=False, default=r'H:\MXD_Scan\config.xml')

    params = parser.parse_args()

    if os.path.exists(params.dsf):
        print('**** No Overwrite the existing output XLS file: %s' % params.dsf)
        exit(-1)

    if params.cfg is not None:
        load_config(params.cfg)

    if params.tgt == "Livelink":
        # -t Livelink -x "H:\MXD_Scan\xlsx\arcgis_files\PASS_1\mxds_load_all" -o "H:\nbProjects\LDrive2Livelink_Data_Path_2.xlsx"
        compile_data_folders_For_Livelink(params.xls, params.dsf)
        print('**** The compiled data folder list for Livelink: %s' % params.dsf)
    elif params.tgt == "SDE":
        # -t SDE -x "H:\MXD_Scan\xlsx\arcgis_files\PASS_1\mxds_load_all" -o "H:\nbProjects\LDrive2SDE_Data_Path_2.xlsx"
        compile_data_sources_For_SDE(params.xls, params.dsf)
        print('**** The compiled data source list for SDE: %s' % params.dsf)

