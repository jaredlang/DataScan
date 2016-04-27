#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Chen.Liang
#
# Created:     18/04/2016
# Copyright:   (c) Chen.Liang 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os, datetime

excluded_folders = [
    r'\sent',
    r'\hh_documents',
]

foldertype_dict = {
    'fgdb': ['.gdb'],
    'kingdom': ['kingdom'],
}

datatype_dict = {
    'gis': ['.shp', '.lyr', '.mxd', '.kmz'],
    'doc': ['.pdf', '.doc', '.docx', '.msg'],
    'excel': ['.xls', '.xlsx', '.csv'],
    'cad': ['.dwg', '.dgn'],
    'video': ['.ts', '.vob', '.mts', '.asf', '.mpg', '.mpeg'],
    'image': ['.png', '.jpg', '.tif', '.img', '.rrd', '.ecw'],
    'lidar': ['.las', '.bil', '.flt', '.qtt', '.xyzi', '.txt'],
    'zip': ['.zip'],
    'geophysical': ['.sgy', '.xyz', '.xtf', '.vel', '.grd'],
}

datastats_dict = {}

byte_units = ['B', 'KB', 'MB', 'GB', 'TB']

precision = 2

def bytesInKMGT(bytesize):
    sz0 = bytesize
    if sz0 == 0:
        return str(sz0)
    for u in byte_units:
        sz1 = sz0
        sz0 /= 1024.0
        if sz0 <= 1:
            return str(round(sz1, precision)) + u
    return str(round(sz0, precision)) + byte_units[-1]


def get_foler_size(dirpath):
    total_size = 0
    for f in os.listdir(dirpath):
        fp = os.path.join(dirpath, f)
        if os.path.isfile(fp):
            try:
                total_size += os.path.getsize(fp)
            except:
                print 'can\'t open ' + fp
    return total_size


def prn_stats():
    for dt in datastats_dict:
        print dt + ' -'
        print '\t count: ' + str(datastats_dict[dt]['count'])
        print '\t size: ' + bytesInKMGT(datastats_dict[dt]['size'])


def scan_data(root_folder):
    for dirpath, dirnames, filenames in os.walk(root_folder):
        # skip any unwanted folders
        for excld in excluded_folders:
            if dirpath.lower().find(excld) > -1:
                continue

        is_data_composite = False
        print 'dir path = ' + dirpath
        fldrname, dirname = os.path.split(dirpath)
        dirname = dirname.lower()
        for ft in foldertype_dict:
            for dt in foldertype_dict[ft]:
                # dirname ending with dt
                if dirname.endswith(dt) == True:
                    if dt not in datastats_dict.keys():
                        datastats_dict[dt] = {'size': 0, 'count': 0}
                    datastats_dict[dt]['size'] += get_foler_size(dirpath)
                    datastats_dict[dt]['count'] += 1
                    is_data_composite = True
                    break

        if is_data_composite == True:
            # count as one and skip individual files
            continue

        for dirname in dirnames:
            # print 'sub folder: ' + dirname
            None
        for fname in filenames:
            # print 'file name: ' + fname
            fnm, fext = os.path.splitext(fname)
            fext = fext.lower()
            for dt in datatype_dict:
                if fext in datatype_dict[dt]:
                    if dt not in datastats_dict.keys():
                        datastats_dict[dt] = {'size': 0, 'count': 0}
                    fp = os.path.join(dirpath, fname)
                    try:
                        datastats_dict[dt]['size'] += os.path.getsize(fp)
                    except:
                        print 'can\'t open ' + fp
                    datastats_dict[dt]['count'] += 1
                    break

    prn_stats()


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    print 'start at %s' % start_time

    scan_data(r'G:\Tessellations\CAS')

    end_time = datetime.datetime.now()
    print 'complete at %s' % end_time

    print 'time elapsed: %s' % (end_time - start_time)


