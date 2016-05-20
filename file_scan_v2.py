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
import threading
import Queue
import logging

logger = logging.getLogger()
# handler = logging.StreamHandler()
handler = logging.FileHandler(r'C:\Users\kdb086\Documents\DataScan\file_scan_MOZGIS_2a.log')
formatter = logging.Formatter('%(asctime)s %(levelname)-8s %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


# constants
EXCLUDED_FOLDERS = [
    r'\sent',
    r'\VENDOR',
    r'\WORKING',
    r'\arcgis_files', # 3a
]

FOLDERTYPE_DICT = {
    'fgdb': ['.gdb'],
    'kingdom': ['kingdom'],
}

DATATYPE_DICT = {
    'gis': ['.shp', '.lyr', '.kmz', '.kml', '.mdb'],
    'doc': ['.pdf', '.doc', '.docx', '.msg'],
    'excel': ['.xls', '.xlsx', '.csv'],
    'cad': ['.dwg', '.dgn'],
    'video': ['.ts', '.vob', '.mts', '.asf', '.mpg', '.mpeg'],
    'image': ['.png', '.jpg', '.bmp', '.tif', '.img', '.rrd', '.ecw'],
    'lidar': ['.las', '.bil', '.flt', '.qtt', '.xyzi', '.txt'],
    'zip': ['.zip'],
    'geophysical': ['.sgy', '.xyz', '.xtf', '.vel', '.grd'],
}

BYTE_UNITS = ['B', 'KB', 'MB', 'GB', 'TB']

PRECISION = 2

# functional code
def bytesInKMGT(bytesize):
    sz0 = bytesize
    if sz0 == 0:
        return str(sz0)
    for u in BYTE_UNITS:
        sz1 = sz0
        sz0 /= 1024.0
        if sz0 <= 1:
            return str(round(sz1, PRECISION)) + u
    return str(round(sz0, PRECISION)) + BYTE_UNITS[-1]


def prn_stats(result_dict):
    for dt in result_dict:
        logger.info(dt + ' -')
        logger.info('\t count: ' + str(result_dict[dt]['count']))
        logger.info('\t size: ' + bytesInKMGT(result_dict[dt]['size']))


def scan_data(root_folder, result_queue):
    this_thread = threading.currentThread()
    localdata = threading.local()
    localdata.ds_dict = {}
    for localdata.dirpath, localdata.dirnames, localdata.filenames in os.walk(root_folder):
        # skip any unwanted folders
        localdata.is_folder_excluded = False
        for localdata.excld in EXCLUDED_FOLDERS:
            if localdata.dirpath.lower().find(localdata.excld.lower()) > -1:
                localdata.is_folder_excluded = True
                break
        if localdata.is_folder_excluded == True:
            continue

        logger.info('[%s] dirpath = %s' % (this_thread.getName(), localdata.dirpath))

        localdata.fldrname, localdata.dirname = os.path.split(localdata.dirpath)
        localdata.dirname = localdata.dirname.lower()
        localdata.is_data_composite = False

        for ft in FOLDERTYPE_DICT:
            for dt in FOLDERTYPE_DICT[ft]:
                # dirname ending with dt
                if localdata.dirname.endswith(dt) == True:
                    if ft not in localdata.ds_dict.keys():
                        localdata.ds_dict[ft] = {'size': 0, 'count': 0}
                    localdata.is_data_composite = True
                    localdata.ds_dict[ft]['count'] += 1
                    localdata.total_size = 0
                    for f in os.listdir(localdata.dirpath):
                        localdata.fp = os.path.join(localdata.dirpath, f)
                        if os.path.isfile(localdata.fp):
                            try:
                                localdata.total_size += os.path.getsize(localdata.fp)
                            except:
                                logger.info('can\'t open ' + localdata.fp)
                    localdata.ds_dict[ft]['size'] += localdata.total_size
                    break

        if localdata.is_data_composite == True:
            # count as one and skip individual files
            continue

        for dirname in localdata.dirnames:
            # logger.info('sub folder: ' + dirname
            None
        for localdata.fname in localdata.filenames:
            # logger.info('file name: ' + localdata.fname
            localdata.fnm, localdata.fext = os.path.splitext(localdata.fname)
            localdata.fext = localdata.fext.lower()
            for dt in DATATYPE_DICT:
                if localdata.fext in DATATYPE_DICT[dt]:
                    if dt == 'gis':
                        logger.info('gis data found in %s' % localdata.dirpath)
                    if dt not in localdata.ds_dict.keys():
                        localdata.ds_dict[dt] = {'size': 0, 'count': 0}
                    localdata.fp = os.path.join(localdata.dirpath, localdata.fname)
                    try:
                        localdata.ds_dict[dt]['size'] += os.path.getsize(localdata.fp)
                    except:
                        logger.info('can\'t open %s' % localdata.fp)
                    localdata.ds_dict[dt]['count'] += 1
                    break

    result_queue.put(localdata.ds_dict)


def scan_data_mr(root_folder):
    # map each subdir into a thread
    q = Queue.Queue()
    subfolders = os.listdir(root_folder)
    subfolders.sort()
    for sfld in subfolders:
        t = threading.Thread(name=sfld, target=scan_data, args=(os.path.join(root_folder, sfld), q, ))
        t.start()
    # wait until all threads complete
    main_thread = threading.currentThread()
    for t in threading.enumerate():
        if t is main_thread:
            continue
        t.join()
    # sum up results from each thread
    datastats_dict = {}
    while not q.empty():
        result_dict = q.get()
        for dt in result_dict.keys():
            entry = result_dict[dt]
            if dt not in datastats_dict.keys():
                datastats_dict[dt] = entry
            else:
                datastats_dict[dt]['size'] += entry['size']
                datastats_dict[dt]['count'] += entry['count']
    # logger.info(out final result
    prn_stats(datastats_dict)


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    logger.info('start at %s' % start_time)

    scan_data_mr(r'P:\\')

    end_time = datetime.datetime.now()
    logger.info('complete at %s' % end_time)

    logger.info('time elapsed: %s' % (end_time - start_time))



