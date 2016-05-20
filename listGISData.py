#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Chen.Liang
#
# Created:     17/05/2016
# Copyright:   (c) Chen.Liang 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os, datetime
import arcpy
import logging

logger = logging.getLogger()
# handler = logging.StreamHandler()
handler = logging.FileHandler(r'C:\Users\kdb086\Documents\DataScan\list_gis_data_4b.log')
formatter = logging.Formatter('%(asctime)s %(levelname)-8s %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


# constants
EXCLUDED_FOLDERS = [
    r'\sent',
    r'\VENDOR',
    r'\WORKING',
]

def scan_gis_data(root_folder):
    for dirpath, dirnames, filenames in os.walk(root_folder):
        # skip any unwanted folders
        is_folder_excluded = False
        for excld in EXCLUDED_FOLDERS:
            if localdata.dirpath.lower().find(excld) > -1:
                is_folder_excluded = True
                break
        if is_folder_excluded == True:
            continue

        # logger.info('dirpath = %s' % dirpath)
        arcpy.env.workspace = dirpath

        # list feature classes
        for fc in arcpy.ListFeatureClasses():
            logger.info('FeatureClass: %s' % os.path.join(dirpath, fc))


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    logger.info('start at %s' % start_time)

    scan_gis_data(r'P:\\')

    end_time = datetime.datetime.now()
    logger.info('complete at %s' % end_time)

    logger.info('time elapsed: %s' % (end_time - start_time))

