#-------------------------------------------------------------------------------
# Name:        List data (tables and feature classes inside a file geodatabase
# Purpose:
#
# Author:      kdb086
# Created:     03/05/2016
# Copyright:   (c) kdb086 2016
#-------------------------------------------------------------------------------

import os
import arcpy
from arcpy import env
import logging

logger = logging.getLogger()
# handler = logging.StreamHandler()
handler = logging.FileHandler(r'C:\Users\kdb086\Documents\DataScan\listDataInGDB.log')
formatter = logging.Formatter('%(asctime)s %(levelname)-8s %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


def listDataInGDB(fgdb):
    # Set the current workspace
    env.workspace = fgdb

    # Get and print a list of feature classes
    fcs = arcpy.ListFeatureClasses()
    fcs.sort()
    for fc in fcs:
        logger.info('featureclass: ' + fc)

    # Get and print a list of tables
    tables = arcpy.ListTables()
    tables.sort()
    for table in tables:
        logger.info('table: ' + table)


def listGDBsInFolder(folder_path):
    for root, dirs, files in os.walk(folder_path):
        if not os.path.exists(root):
            logging.error('Path "'+root+'" doesn\'t exist!')
            break

        if root.lower().endswith(".gdb"):
            listDataInGDB(root)


if __name__ == '__main__':
    listGDBsInFolder(r'C:\Users\kdb086\Documents\ArcGIS')
