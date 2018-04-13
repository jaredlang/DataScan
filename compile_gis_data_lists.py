import os
import argparse
import pandas as pd
import xlrd

def combine_xls_files(xls_inputs, xls_output):
    # read them in
    #excels = [pd.ExcelFile(name) for name in xls_inputs]
    excels = []
    for fname in xls_inputs:
        try:
            excels.append(pd.ExcelFile(fname))
        except xlrd.biffh.XLRDError as e:
            print("#### Corrupted file: %s" % fname)

    # turn them into dataframes
    frames = [x.parse(x.sheet_names[0], header=None,index_col=None) for x in excels]

    # delete the first row for all frames except the first
    # i.e. remove the header row -- assumes it's the first
    frames[1:] = [df[1:] for df in frames[1:]]

    # concatenate them..
    combined = pd.concat(frames)

    # write it out
    combined.to_excel(xls_output, header=False, index=False)


def combine_xls_files_in_dir(xlsFolder, xlsOutput):
    xlsFiles = []
    # browse for xls files
    for fname in os.listdir(xlsFolder):
        fpath = os.path.join(xlsFolder, fname)
        if os.path.isfile(fpath) and os.path.splitext(fname)[1] == ".xlsx":
            xlsFiles.append(fpath)

    combine_xls_files(xlsFiles, xlsOutput)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Combine the excel files into one')

    parser.add_argument('-d','--dir', help='A direcotry for the input Excel files', required=True)
    parser.add_argument('-x','--xls', help='A path for the output Excel file', required=True)
    parser.add_argument('-a','--action', help='Action Options (scan)', required=False, default='scan')

    params = parser.parse_args()

    if params.action == "scan":
        combine_xls_files_in_dir(params.dir, params.xls)
        print('###### Completed combining xls files in [%s]'%params.xls)
    else:
        print 'Error: unknown action [%s] for combining' % params.action
