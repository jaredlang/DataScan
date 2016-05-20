#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      kdb086
#
# Created:     11/05/2016
# Copyright:   (c) kdb086 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------

import arcpy, os
from arcpy import env

def scan_mxd(scan_folder):
    total_count = 0
    bad_mxd_count = 0
    rel_path_count = 0
    abs_path_count = 0
    for root, dirs, files in os.walk(scan_folder):
        mxdList = [fn for fn in files
            if fn.lower().endswith(".mxd")]

        #get relative path setting for each MXD in list.
        for file in mxdList:
            #set map document to change
            filePath = os.path.join(root, file)
            total_count += 1
            try:
                mxd = arcpy.mapping.MapDocument(filePath)
                if mxd.relativePaths == True:
                    rel_path_count += 1
                else:
                    abs_path_count += 1
##                #set relative paths property
##                mxd.relativePaths = True
##                #save map doucment change
##                mxd.save()
                del mxd
            except:
                bad_mxd_count += 1

    print 'total mxd count: %i' % total_count
    print 'bad mxd count: %i, %.2f' % (bad_mxd_count, bad_mxd_count*1.0/total_count)
    print 'count of mxd w/ relative path : %i %.2f' % (rel_path_count, rel_path_count*1.0/total_count)
    print 'count of mxd w/ absolute path : %i %.2f' % (abs_path_count, abs_path_count*1.0/total_count)
    

if __name__ == '__main__':
    scan_mxd(r"L:\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS - Copy")
