import os
import pandas as pd
import re
import argparse

def split_data_by_types(xls_file_path):
    fbase,fext = os.path.splitext(xls_file_path)

    xls_file = pd.ExcelFile(xls_file_path)
    df = xls_file.parse(xls_file.sheet_names[0], header=0)

    ESRI_VECTOR_EXT = re.compile(r"\.shp$|\.mdb\\|\.gdb\\", flags=re.IGNORECASE)
    CAD_DRAWING_EXT = re.compile(r"\.dwg\\|\.dxf\\|\.dgn\\", flags=re.IGNORECASE)

    is_not_loaded = df["Loaded In SDE?"] == "No"
    is_esri_vector = df["Data Source"].str.contains(ESRI_VECTOR_EXT)
    is_cad_drawing = df["Data Source"].str.contains(CAD_DRAWING_EXT)

    df_vector = df[is_not_loaded & (is_esri_vector | is_cad_drawing)]
    df_vector.to_excel("%s_esri-cad%s" % (fbase, fext), index=False)
    df_vector = None

    df_esri_vector = df[is_not_loaded & is_esri_vector]
    df_esri_vector.to_excel("%s_esri%s" % (fbase, fext), index=False)
    df_esri_vector = None

    df_esri_vector_w_loaded = pd.concat([df_esri_vector, df[~is_not_loaded]])
    df_esri_vector_w_loaded.to_excel("%s_esri_w_loaded%s" % (fbase, fext), index=False)
    df_esri_vector_w_loaded = None

    df_cad_drawing = df[is_not_loaded & is_cad_drawing]
    df_cad_drawing.to_excel("%s_cad%s" % (fbase, fext), index=False)
    df_cad_drawing = None

    df_other_data = df[is_not_loaded & ~(is_esri_vector | is_cad_drawing)]
    df_other_data.to_excel("%s_others%s" % (fbase, fext), index=False)
    df_other_data = None

    df = None
    xls_file = None


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Filter data based on types')
    parser.add_argument('-x','--xls', help='XLS File (input)', required=True)

    params = parser.parse_args()

    split_data_by_types(params.xls)
    print('**** The data source list was splitted: %s' % params.xls)
