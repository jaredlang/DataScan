#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      kdb086
#
# Created:     27/04/2016
# Copyright:   (c) kdb086 2016
# Licence:     <your licence>
#-------------------------------------------------------------------------------
import os

def distinct(seq):
    seen = set()
    seen_add = seen.add
    return [x for x in seq if not (x in seen or seen_add(x))]


def write_distinct(src_file, tgt_file):
    with open(tgt_file, 'w') as tgtf:
        with open(src_file, 'r') as srcf:
            lines = srcf.readlines()
            lines.sort()
            distinct_lines = distinct(lines)
            tgtf.writelines(distinct_lines)


def extract(src_file, tgt_file, key_value, limiters):
    limiters.append(key_value)
    with open(tgt_file, 'w') as tgtf:
        with open(src_file, 'r') as srcf:
            lines = srcf.readlines()
            for ln in lines:
                i = ln.find(key_value)
                while i > -1:
                    e = -1
                    for lmt in limiters:
                        x = ln.find(lmt, i+1)
                        if e == -1:
                            e = x
                        elif x > -1:
                            e = min(x, e)

                    if e > -1:
                        tgtf.write(ln[i:e].replace('\n', '').strip() + '\n')
                        i = ln.find(key_value, i+1)
                    else:
                        tgtf.write(ln[i:].replace('\n', '').strip() + '\n')
                        break

if __name__ == '__main__':
    root_folder = r'C:\Users\kdb086\Documents\DataScan'
    # extract gis data folders
    extract(os.path.join(root_folder, 'raw_result_test.txt'),
        os.path.join(root_folder, 'gis_folder_list_tmp.txt'),
        'gis data found', ['[', 'dir path', 'can\'t open'])
    write_distinct(os.path.join(root_folder, 'gis_folder_list_tmp.txt'),
        os.path.join(root_folder, 'gis_folder_list.txt'))

    # extract inaccessible data folders
    extract(os.path.join(root_folder, 'raw_result_test.txt'),
        os.path.join(root_folder, 'cant_open_file_list_tmp.txt'),
        'can\'t open', ['[', 'dir path', 'gis data found'])
    write_distinct(os.path.join(root_folder, 'cant_open_file_list_tmp.txt'),
        os.path.join(root_folder, 'cant_open_file_list.txt'))
