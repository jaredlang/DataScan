#-------------------------------------------------------------------------------
# Name:        list content of a folder with mxd's
# Purpose:
#
# Author:      kdb086
# Created:     03/05/2016
# Copyright:   (c) kdb086 2016
#-------------------------------------------------------------------------------

import arcpy, os, sys, getpass, datetime, time
# import re
from arcpy import env
from operator import itemgetter
import logging

# configure logger
logger = logging.getLogger()
# handler = logging.StreamHandler()
handler = logging.FileHandler(r'C:\Users\kdb086\Documents\DataScan\listFolderMXD_MOZGIS.log')
# formatter = logging.Formatter('%(asctime)s %(levelname)-8s %(message)s')
formatter = logging.Formatter('%(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)


env.overwriteOutput = True

# sdeUser = os.environ['USERNAME']
# sdeClient = os.environ['COMPUTERNAME']

# Get transformation method
def getTrans(df):
    transList = df.geographicTransformations
    if len(transList) == 0:
        return '<none>'
    elif len(transList) == 1:
        return transList[0]
    else:
        return ' & '.join(transList)

# Get scale range
def getScaleRange(lyr):
    minScale = lyr.minScale
    maxScale = lyr.maxScale
    if minScale <> maxScale:
        return (str(minScale)+' to '+str(maxScale))
    else:
        return ''

# Get labels
def getLabel(lyr):
    if lyr.showLabels:
        for lblClass in lyr.labelClasses:
                if lblClass.showClassLabels:
                    return (lblClass.className+' | '+lblClass.expression+' | '+lblClass.SQLQuery)
    else:
        return ''


# Go through each mxd files
def listLayersInMXD(prosFolder, start_filename):
    if start_filename is not None:
        start_filename = start_filename.lower()

    for root, dirs, files in os.walk(prosFolder):
        if not os.path.exists(root):
            logging.error('Path "'+root+'" doesn\'t exist!')
            break

        filtered_files = [fn for fn in files
            if fn.lower().endswith(".mxd") and fn.lower() >= start_filename]

        for file in sorted(filtered_files): # Sort file alphabetically
            try:
                fp = os.path.join(root,file)
                logging.info('---- List MXD content log [%s]----' % fp)
                start_ts = datetime.datetime.now()
                logging.info('-- Starting time: %s' % start_ts)
                #######################
                # Find active dataframe
                mxd = arcpy.mapping.MapDocument(fp)
                dfList = arcpy.mapping.ListDataFrames(mxd)
                logging.info('Input: "'+mxd.filePath)
                logging.info("last modified: %s" % time.ctime(os.path.getmtime(fp)))
                for df in dfList:
                    transMethod = getTrans(df)
                    logging.info('*'*50)
                    logging.info('** Data Frame: '+df.name)
                    logging.info('('+df.spatialReference.name+' - Transformation: '+transMethod)
                    srcList = []
                    ##################
                    # List fc
                    fcList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isFeatureLayer]
                    if len(fcList) > 0:
                        logging.info('='*20)
                        logging.info('* Feature Class:')
                    grpName = tmpGrpName = ''
                    sdeList = shpList = ''
                    for fc in fcList:
                        fcName = 'Feature Class: '
                        cnt = len(fc.longName.split('\\'))
                        if cnt > 1:
                            for i in range(1,cnt):
                                grpName += fc.longName.split('\\')[i-1]+' | '
                            grpName = grpName[:-3]
                            if grpName <> tmpGrpName:
                                logging.info('-'*12)
                                logging.info('GROUP: '+grpName)
                                tmpGrpName = grpName
                        else:
                            logging.info('-'*12)
                        fcName += fc.longName.split('\\')[-1]

                        if fc.supports('serviceProperties'): #???
                            #if fc.workspacePath == '': # feature dataset classes + others?
                            #if len(fc.dataSource.split('\\')) == 3: # feature dataset classes + others?
                            try:
                                db = fc.serviceProperties['Db_Connection_Properties'] # Direct Connect
                            except:
                                db = fc.serviceProperties['Server'] # 3-tier Connect
                            if len(fc.dataSource.split('\\')) == 3: # feature dataset classes + others?
                                #print '-'*12
                                logging.info('- '+fc.name+': '+db.upper()+' | '+\
                                      fc.serviceProperties['Service'].upper()+' | '+\
                                      fc.dataSource.split('\\')[1]+' | '+fc.dataSource.split('\\')[2]+\
                                      '('+fc.serviceProperties['Version']+'/'+fc.serviceProperties['UserName']+')') # TOC: SDE | service | FD |FC (version/user)
                                srcList.append((db.upper()+' | '+fc.serviceProperties['Service'].upper()+' | '+\
                                             fc.serviceProperties['Version']+' | '+fc.serviceProperties['UserName'],grpName,fc.name,fc.dataSource.split('\\')[2]))
                            else: # feature class
                                #print '-'*12
                                logging.info('- '+fc.name+': '+db.upper()+' | '+\
                                      str(fc.serviceProperties['Service']).upper()+' | '+\
                                      fc.datasetName+'('+fc.serviceProperties['Version']+'/'+fc.serviceProperties['UserName']+')') # TOC: SDE | service | FC (version/user)
                                srcList.append((db.upper()+' | '+fc.serviceProperties['Service'].upper()+' | '+\
                                             fc.serviceProperties['Version']+' | '+fc.serviceProperties['UserName'],grpName,fc.name,fc.datasetName))
                        else: # ?
                            # if bool(re.match(fc.dataSource[-4:],'.SHP', re.I)): # shapefile
                            if fc.dataSource[-4:].upper().find('.SHP') > -1: # shapefile
                                logging.info('-'*12)
                                logging.info('- '+fc.name+' | '+fc.workspacePath+' | '+fc.datasetName+'.shp') # TOC: path | shape_name
                                srcList.append((fc.workspacePath,grpName,fc.name,fc.datasetName+'.shp'))
                            # elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.ODC', re.I)): # XY event
                            elif fc.workspacePath <> '' and fc.workspacePath[-4:].upper().find('.ODC') > -1: # XY event
                                logging.info('-'*12)
                                logging.info('- '+fc.name+': '+fc.dataSource.split('\\')[-1]+' (XY Event)') # TOC: table_name
                            # elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.GDB', re.I)): # File GDB
                            elif fc.workspacePath <> '' and fc.workspacePath[-4:].upper().find('.GDB') > -1: # File GDB
                                logging.info('-'*12)
                                logging.info('- '+fc.name+': '+fc.dataSource.replace(fc.datasetName,'')+' | '+fc.datasetName) # TOC: file_gdb_path+FD | fc
                                srcList.append((fc.dataSource.replace(fc.datasetName,''),grpName,fc.name,fc.datasetName))
                            # elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.MDB', re.I)): # Personal GDB
                            elif fc.workspacePath <> '' and fc.workspacePath[-4:].upper().find('.MDB') > -1: # Personal GDB
                                logging.info('-'*12)
                                logging.info('- '+fc.name+': '+fc.dataSource.replace(fc.datasetName,'')+' | '+fc.datasetName) # TOC: file_gdb_path+FD | fc
                                srcList.append((fc.dataSource.replace(fc.datasetName,''),grpName,fc.name,fc.datasetName))
                            # elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.SDE', re.I)): # SDE - but may have error
                            elif fc.workspacePath <> '' and fc.workspacePath[-4:].upper().find('.SDE') > -1: # SDE - but may have error
                                logging.info('- '+fc.name+': '+fc.dataSource+' | '+fc.datasetName) # TOC: path | fc
                                srcList.append((fc.workspacePath,grpName,fc.name,fc.datasetName))
                            else:
                                logging.info('-'*12)
                                logging.info('no service properties:')
                                logging.info('- '+fc.name+' | ' +fc.workspacePath+' | '+fc.dataSource+' | '+fc.datasetName)

                        # List projection
                        try:
                            prjName = arcpy.Describe(fc).spatialReference.name
                            logging.info('COORD SYSTEM - '+prjName)
                        except:
                            pass

                        # List scale depedency/range
                        try:
                            scaleText = getScaleRange(fc)
                            if scaleText <> '':
                                logging.info('SCALE RANGE - '+scaleText)
                        except:
                            pass

                        # List label
                        try:
                            lblText = getLabel(fc)
                            if lblText <> '':
                                logging.info('LABEL - '+lblText)
                        except:
                            pass

                        # List definition query
                        if fc.supports("DEFINITIONQUERY"):
                            defString = fc.definitionQuery
                            if defString <> '':
                                logging.info('DEFINITION QUERY - '+defString)

                        grpName = ''
                        fcName = 'Feature Class: '

                    ##########################
                    # List raster
                    rastList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isRasterLayer]
                    if len(rastList) > 0:
                        logging.info('='*20)
                        logging.info('* Raster:')
                    grpName = tmpGrpName = ''
                    for rast in rastList:
                        #svName = 'ArcGIS Service: '
                        cnt = len(rast.longName.split('\\'))
                        if cnt > 1:
                            for i in range(1,cnt):
                                grpName += rast.longName.split('\\')[i-1]+' | '
                            grpName = grpName[:-3]
                            if grpName <> tmpGrpName:
                                logging.info('-'*12)
                                logging.info('GROUP: '+grpName)
                                tmpGrpName = grpName
                        else:
                            logging.info('-'*12)

                        if not rast.supports('serviceProperties'): # not Image server service
                            logging.info('- '+rast.name+': '+rast.dataSource) # TOC: rast_source
                            #srcList.append((rast.dataSource,grpName,rast.name,''))
                            srcList.append((rast.workspacePath,grpName,rast.name,rast.datasetName))

                        # List projection
                        try:
                            prjName = arcpy.Describe(rast).spatialReference.name
                            logging.info('(COORD SYSTEM - '+prjName)
                        except:
                            pass

                        # List scale depedency/range
                        try:
                            scaleText = getScaleRange(rast)
                            if scaleText <> '':
                                logging.info('(SCALE RANGE - '+scaleText)
                        except:
                            pass

                        # List label
                        try:
                            lblText = getLabel(rast)
                            if lblText <> '':
                                logging.info('(LABEL - '+lblText)
                        except:
                            pass

                        # List definition query
                        if rast.supports("DEFINITIONQUERY"):
                            defString = rast.definitionQuery
                            if defString <> '':
                                logging.info('(DEFINITION QUERY - '+defString)

                        grpName = ''

                    ##########################
                    # List service
                    svList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isServiceLayer and lyr.supports('serviceProperties')]
                    if len(svList) > 0:
                        logging.info('='*20)
                        logging.info('* ArcGIS Service:')
                    grpName = tmpGrpName = ''
                    for sv in svList:
                        #svName = 'ArcGIS Service: '
                        cnt = len(sv.longName.split('\\'))
                        if cnt > 1:
                            for i in range(1,cnt):
                                grpName += sv.longName.split('\\')[i-1]+' | '
                            grpName = grpName[:-3]
                            if grpName <> tmpGrpName:
                                logging.info('-'*12)
                                logging.info('GROUP: '+grpName)
                                tmpGrpName = grpName
                        else:
                            logging.info('-'*12)

                        if not sv.isGroupLayer:
                            logging.info('- '+sv.name+': '+sv.serviceProperties['ServiceType']+' | '+sv.serviceProperties['URL']) # TOC: type | url

                        srcList.append((sv.serviceProperties['URL'],grpName,sv.name,sv.serviceProperties['ServiceType']))

                        # List projection
                        try:
                            prjName = arcpy.Describe(sv).spatialReference.name
                            logging.info(' (COORD SYSTEM - '+prjName+')')
                        except:
                            pass

                        # List scale depedency/range
                        try:
                            scaleText = getScaleRange(sv)
                            if scaleText <> '':
                                logging.info('(SCALE RANGE - '+scaleText+')')
                        except:
                            pass

                        # List label
                        try:
                            lblText = getLabel(sv)
                            if lblText <> '':
                                logging.info('(LABEL - '+lblText+')')
                        except:
                            pass

                        # List definition query
                        if sv.supports("DEFINITIONQUERY"):
                            defString = sv.definitionQuery
                            if defString <> '':
                                logging.info('(DEFINITION QUERY - '+defString+')')

                        grpName = ''

                    ##########################
                    # List ESRI Basemap
                    #url = 'http://goto.arcgisonline.com'
                    urlString = 'arcgisonline.com'
                    bmList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.description.find(urlString) > -1]
                    if len(bmList) > 0:
                        logging.info('='*20)
                        logging.info('* ESRI Basemap:')
                    grpName = tmpGrpName = ''
                    for bm in bmList:
                        cnt = len(bm.longName.split('\\'))
                        if cnt > 1:
                            for i in range(1,cnt):
                                grpName += bm.longName.split('\\')[i-1]+' | '
                            grpName = grpName[:-3]
                            if grpName <> tmpGrpName:
                                logging.info('-'*12)
                                logging.info('GROUP: '+grpName)
                                tmpGrpName = grpName
                        else:
                            logging.info('-'*12)

                        if not bm.isGroupLayer:
                            logging.info('- '+bm.name+': '+lyr.description[lyr.description.find('http:'):]) # TOC: url

                        srcList.append((lyr.description[lyr.description.find('http:'):],grpName,bm.name,''))

                        # List projection
                        try:
                            prjName = arcpy.Describe(bm).spatialReference.name
                            logging.info(' (COORD SYSTEM - '+prjName+')')
                        except:
                            pass

                        # List scale depedency/range
                        try:
                            scaleText = getScaleRange(bm)
                            if scaleText <> '':
                                logging.info('(SCALE RANGE - '+scaleText+')')
                        except:
                            pass

                        # List label
                        try:
                            lblText = getLabel(bm)
                            if lblText <> '':
                                logging.info('(LABEL - '+lblText+')')
                        except:
                            pass

                        # List definition query
                        if bm.supports("DEFINITIONQUERY"):
                            defString = bm.definitionQuery
                            if defString <> '':
                                logging.info('(DEFINITION QUERY - '+defString+')')

                        grpName = ''

                    ##########################
                    # List table
                    tblList = [tbl for tbl in arcpy.mapping.ListTableViews(mxd, data_frame=df)]
                    if len(tblList) > 0:
                        logging.info('='*20)
                        logging.info('* Table:')
                        logging.info('-'*12)

                        for tbl in tblList:
                            logging.info('- '+tbl.name+': '+tbl.workspacePath+' | '+tbl.datasetName) # TOC: path | table
                            srcList.append((tbl.workspacePath,'',tbl.name,tbl.datasetName)) # [(PATH,GROUP,TOC,name),(...)]

                    ##########################
                    # List broken layers
                    brokeList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isBroken]
                    if len(brokeList) > 0:
                        logging.info('x'*40)
                        logging.info('* BROKEN LAYERS:')
                    grpName = tmpGrpName = ''
                    for broke in brokeList:
                        #svName = 'ArcGIS Service: '
                        cnt = len(broke.longName.split('\\'))
                        if cnt > 1:
                            for i in range(1,cnt):
                                grpName += broke.longName.split('\\')[i-1]+' | '
                            grpName = grpName[:-3]
                            if grpName <> tmpGrpName:
                                logging.info('-'*12)
                                logging.info('GROUP: '+grpName)
                                tmpGrpName = grpName
                        else:
                            logging.info('-'*12)
                        logging.info('- '+broke.name)
                        #svName += sv.longName.split('\\')[-1]

                        grpName = ''

                    ##########################
                    # List data source
                    #print srcList # [(PATH,GROUP,TOC,name),(...)]
                    sortSrcList = sorted(srcList, key=itemgetter(0,1,2,3)) # sort srcList
                    if len(sortSrcList) > 0:
                        logging.info('#'*40)
                        logging.info('# DATA SOURCE: #')
                    tmpPath=''
                    for row in sortSrcList:
                        path=row[0]
                        group=row[1]
                        toc=row[2]
                        name=row[3]
                        if path<>tmpPath:
                            logging.info('-'*12)
                            logging.info('* SOURCE: "'+path+'"')
                        if group<>'':
                            logging.info('- GROUP: '+group+' | '+toc+' | '+name)
                        else:
                            logging.info('- '+toc+' | '+name)
                        tmpPath=path

                logging.info('*'*50)
                # get the ending time
                end_ts = datetime.datetime.now()
                logging.info('-- End time: %s' % end_ts)

                # calculate the processing time
                logging.info('Processing time on [%s]: %s' % (file ,end_ts - start_ts))

                del mxd

            except:
                logging.error('error out on %s' % file)


if __name__ == '__main__':
    prosFolder = arcpy.GetParameterAsText(0)
    prosFolder = r"L:\SharedData\Houston\IntlDeepW\MOZAMBIQUE\MOZGIS - Copy\arcgis_files"
    start_filename = arcpy.GetParameterAsText(1)
    start_filename = 'Potential_Resettlement_Sites_RS2_EL_prelim.mxd'
    listLayersInMXD(prosFolder, start_filename)
