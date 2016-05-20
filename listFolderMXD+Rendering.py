# listFolderMXD+Rendering.py
# - list content of a folder with mxd's
# - jyl (6-25-2015)
# Added rendering time (8-29-15)

import arcpy, os, sys, getpass, time, tempfile, jTool, re
from arcpy import env
from os.path import join
from operator import itemgetter

env.overwriteOutput = True
userID = getpass.getuser()

sdeUser = os.environ['USERNAME']
sdePass = os.environ['USERNAME']
###print os.environ['COMPUTERNAME']

#mxd = arcpy.mapping.MapDocument(r"\\anadarko.com\world\Temporary(30Days)\Houston\jyl\esri.mxd")
#mxd = arcpy.mapping.MapDocument(arcpy.GetParameterAsText(0))

prosFolder = arcpy.GetParameterAsText(0)
#prosFolder = r"\\anadarko.com\world\Temporary(30Days)\Houston\jyl\Bhavna"

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

for root, dirs, files in os.walk(prosFolder):
    if not os.path.exists(root):
        errorMessage += '"'+root+'" doesn\'t exist!\n'
        break

    for f in sorted(files): # Sort file alphabetically
        if f.endswith(".mxd") or f.endswith(".MXD"):
            # Set up mail properties
            sender = 'jia.liu@anadarko.com'
            to = [userID+'@anadarko.com', 'jia.liu@anadarko.com']
            subject = 'list: MXD Content and Rendering Time'
            startTime = time.strftime('%x %X')
            try: 
                startTuple = time.strptime(startTime, "%m/%d/%Y %I:%M:%S %p")
            except:
                startTuple = time.strptime(startTime, "%m/%d/%y %H:%M:%S")
            message = '---- List MXD content and rendering time log ----' + '\n\n'
            message += '-- Starting time: ' + time.strftime('%x %X') + '\n\n'
            errorMessage = ''
            totalTime = 'Total processing time: '
            #######################
            # Find active dataframe
            mxdFullName = join(root,f)
            mxd = arcpy.mapping.MapDocument(mxdFullName)
            dfList = arcpy.mapping.ListDataFrames(mxd)
            print arcpy.AddMessage('\n'+'*'*50+'\nInput: "'+mxd.filePath+'"\n')
            message += 'Input: "'+mxd.filePath+'".\n'
            for df in dfList:
                print arcpy.AddMessage('\n'+'*'*50)
                print arcpy.AddMessage('** Data Frame: '+df.name)
                transMethod = getTrans(df)
                print arcpy.AddMessage('('+df.spatialReference.name+' - Transformation: '+transMethod+').')
                message += '\n'+'*'*50+'\n'
                message += '** Data Frame: '+df.name+'...\n'
                message += '('+df.spatialReference.name+' - Transformation: '+transMethod+').\n'
                srcList = []
                ##################
                # List fc
                fcList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isFeatureLayer]
                if len(fcList) > 0:
                    print arcpy.AddMessage('='*20)
                    print arcpy.AddMessage('* Feature Class:')
                    message += '='*20+'\n'
                    message += '* Feature Class:\n'
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
                            print arcpy.AddMessage('-'*12+'\nGROUP: '+grpName)
                            message += '-'*12+'\nGROUP: '+grpName+'\n'
                            tmpGrpName = grpName
                    else:
                        print arcpy.AddMessage('-'*12)
                        message += '-'*12+'\n'
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
                            print arcpy.AddMessage('- '+fc.name+': '+db.upper()+' | '+\
                                  fc.serviceProperties['Service'].upper()+' | '+\
                                  fc.dataSource.split('\\')[1]+' | '+fc.dataSource.split('\\')[2]+\
                                  '('+fc.serviceProperties['Version']+'/'+fc.serviceProperties['UserName']+')') # TOC: SDE | service | FD |FC (version/user)
                            message += '- '+fc.name+': '+db.upper()+' | '+\
                                  fc.serviceProperties['Service'].upper()+' | '+\
                                  fc.dataSource.split('\\')[1]+' | '+fc.dataSource.split('\\')[2]+\
                                  '('+fc.serviceProperties['Version']+'/'+fc.serviceProperties['UserName']+')\n' # TOC: SDE | service | FD |FC (version/user)
                            srcList.append((db.upper()+' | '+fc.serviceProperties['Service'].upper()+' | '+\
                                         fc.serviceProperties['Version']+' | '+fc.serviceProperties['UserName'],grpName,fc.name,fc.dataSource.split('\\')[2]))
                        else: # feature class
                            #print '-'*12
                            print arcpy.AddMessage('- '+fc.name+': '+db.upper()+' | '+\
                                  str(fc.serviceProperties['Service']).upper()+' | '+\
                                  fc.datasetName+'('+fc.serviceProperties['Version']+'/'+fc.serviceProperties['UserName']+')') # TOC: SDE | service | FC (version/user)
                            message += '- '+fc.name+': '+db.upper()+' | '+\
                                  str(fc.serviceProperties['Service']).upper()+' | '+\
                                  fc.datasetName+'('+fc.serviceProperties['Version']+'/'+fc.serviceProperties['UserName']+')\n' # TOC: SDE | service | FC (version/user)
                            srcList.append((db.upper()+' | '+fc.serviceProperties['Service'].upper()+' | '+\
                                         fc.serviceProperties['Version']+' | '+fc.serviceProperties['UserName'],grpName,fc.name,fc.datasetName))
                    else: # ?
                        if bool(re.match(fc.dataSource[-4:],'.SHP', re.I)): # shapefile
                            #print '-'*12
                            print arcpy.AddMessage('- '+fc.name+' | '+fc.workspacePath+' | '+fc.datasetName+'.shp') # TOC: path | shape_name
                            message += '- '+fc.name+' | '+fc.workspacePath+' | '+fc.datasetName+'.shp\n' # TOC: path | shape_name
                            srcList.append((fc.workspacePath,grpName,fc.name,fc.datasetName+'.shp'))
                        elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.ODC', re.I)): # XY event
                            #print '-'*12
                            print arcpy.AddMessage('- '+fc.name+': '+fc.dataSource.split('\\')[-1]+' (XY Event)') # TOC: table_name
                            message += '- '+fc.name+': '+fc.dataSource.split('\\')[-1]+' (XY Event)\n' # TOC: table_name
                        elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.GDB', re.I)): # File GDB
                            #print '-'*12
                            print arcpy.AddMessage('- '+fc.name+': '+fc.dataSource.replace(fc.datasetName,'')+' | '+fc.datasetName) # TOC: file_gdb_path+FD | fc
                            message += '- '+fc.name+': '+fc.dataSource.replace(fc.datasetName,'')+' | '+fc.datasetName+'\n' # TOC: file_gdb_path+FD | fc
                            srcList.append((fc.dataSource.replace(fc.datasetName,''),grpName,fc.name,fc.datasetName))
                        elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.MDB', re.I)): # Personal GDB
                            #print '-'*12
                            print arcpy.AddMessage('- '+fc.name+': '+fc.dataSource.replace(fc.datasetName,'')+' | '+fc.datasetName) # TOC: personal_gdb_path+FD | fc
                            message += '- '+fc.name+': '+fc.dataSource.replace(fc.datasetName,'')+' | '+fc.datasetName+'\n' # TOC: file_gdb_path+FD | fc
                            srcList.append((fc.dataSource.replace(fc.datasetName,''),grpName,fc.name,fc.datasetName))
                        elif fc.workspacePath <> '' and bool(re.match(fc.workspacePath[-4:],'.SDE', re.I)): # SDE - but may have error
                            print arcpy.AddMessage('- '+fc.name+': '+fc.dataSource+' | '+fc.datasetName) # TOC: path | fc
                            message += '- '+fc.name+': '+fc.dataSource+' | '+fc.datasetName+'\n' # TOC: path | fc
                            srcList.append((fc.workspacePath,grpName,fc.name,fc.datasetName))
                        else:
                            print arcpy.AddMessage('-'*12)
                            print arcpy.AddMessage('other no service properties:')
                            print arcpy.AddMessage(fc.name)
                            print arcpy.AddMessage(fc.workspacePath)
                            print arcpy.AddMessage(fc.dataSource)
                            print arcpy.AddMessage(fc.datasetName)
                            message += '-'*12+'\n'
                            message += 'no service properties:'+'\n'
                            message += fc.name+'\n'
                            message += fc.workspacePath+'\n'
                            message += fc.dataSource+'\n'
                            message += fc.datasetName+'\n'

                    # List projection
                    try:
                        prjName = arcpy.Describe(fc).spatialReference.name
                        print arcpy.AddMessage('('+prjName+')')
                        message += '('+prjName+').\n'
                    except:
                        pass
        
                    # List scale depedency/range
                    try:
                        scaleText = getScaleRange(fc)
                        if scaleText <> '':
                            print arcpy.AddMessage('(SCALE RANGE - '+scaleText+')')
                            message += '(SCALE RANGE - '+scaleText+')\n'
                    except:
                        pass

                    # List label
                    try:
                        lblText = getLabel(fc)
                        if lblText <> '':
                            print arcpy.AddMessage('(LABEL - '+lblText+')')
                            message += '(LABEL - '+lblText+').\n'
                    except:
                        pass

                    # List definition query
                    if fc.supports("DEFINITIONQUERY"):
                        defString = fc.definitionQuery
                        if defString <> '':
                            print arcpy.AddMessage('(DEFINITION QUERY - '+defString+')')
                            message += '(DEFINITION QUERY - '+defString+').\n'

                    grpName = ''
                    fcName = 'Feature Class: '

                ##########################
                # List raster
                rastList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isRasterLayer]
                if len(rastList) > 0:
                    print arcpy.AddMessage('='*20)
                    print arcpy.AddMessage('* Raster:')
                    message += '='*20+'\n'
                    message += '* Raster:\n'
                grpName = tmpGrpName = ''
                for rast in rastList:
                    #svName = 'ArcGIS Service: '
                    cnt = len(rast.longName.split('\\'))
                    if cnt > 1:
                        for i in range(1,cnt):
                            grpName += rast.longName.split('\\')[i-1]+' | '
                        grpName = grpName[:-3]
                        if grpName <> tmpGrpName:
                            print arcpy.AddMessage('-'*12+'\nGROUP: '+grpName)
                            message += '-'*12+'\nGROUP: '+grpName+'\n'
                            tmpGrpName = grpName
                    else:
                        print arcpy.AddMessage('-'*12)
                        message += '-'*12+'\n'

                    if not rast.supports('serviceProperties'): # not Image server service
                        print arcpy.AddMessage('- '+rast.name+': '+rast.dataSource) # TOC: rast_source
                        message += '- '+rast.name+': '+rast.dataSource+'\n' # TOC: rast_source
                        #srcList.append((rast.dataSource,grpName,rast.name,''))
                        srcList.append((rast.workspacePath,grpName,rast.name,rast.datasetName))

                    # List projection
                    try:
                        prjName = arcpy.Describe(rast).spatialReference.name
                        print arcpy.AddMessage('('+prjName+')')
                        message += '('+prjName+').\n'
                    except:
                        pass
        
                    # List scale depedency/range
                    try:
                        scaleText = getScaleRange(rast)
                        if scaleText <> '':
                            print arcpy.AddMessage('(SCALE RANGE - '+scaleText+')')
                            message += '(SCALE RANGE - '+scaleText+')\n'
                    except:
                        pass

                    # List label
                    try:
                        lblText = getLabel(rast)
                        if lblText <> '':
                            print arcpy.AddMessage('(LABEL - '+lblText+')')
                            message += '(LABEL - '+lblText+').\n'
                    except:
                        pass

                    # List definition query
                    if rast.supports("DEFINITIONQUERY"):
                        defString = rast.definitionQuery
                        if defString <> '':
                            print arcpy.AddMessage('(DEFINITION QUERY - '+defString+')')
                            message += '(DEFINITION QUERY - '+defString+').\n'

                    grpName = ''

                ##########################
                # List service
                svList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isServiceLayer and lyr.supports('serviceProperties')]
                if len(svList) > 0:
                    print arcpy.AddMessage('='*20)
                    print arcpy.AddMessage('* ArcGIS Service:')
                    message += '='*20+'\n'
                    message += '* ArcGIS Service:\n'
                grpName = tmpGrpName = ''
                for sv in svList:
                    #svName = 'ArcGIS Service: '
                    cnt = len(sv.longName.split('\\'))
                    if cnt > 1:
                        for i in range(1,cnt):
                            grpName += sv.longName.split('\\')[i-1]+' | '
                        grpName = grpName[:-3]
                        if grpName <> tmpGrpName:
                            print arcpy.AddMessage('-'*12+'\nGROUP: '+grpName)
                            message += '-'*12+'\nGROUP: '+grpName+'\n'
                            tmpGrpName = grpName
                    else:
                        print arcpy.AddMessage('-'*12)
                        message += '-'*12+'\n'
                        
                    if not sv.isGroupLayer:
                        print arcpy.AddMessage('- '+sv.name+': '+sv.serviceProperties['ServiceType']+' | '+sv.serviceProperties['URL']) # TOC: type | url
                        message += '- '+sv.name+': '+sv.serviceProperties['ServiceType']+' | '+sv.serviceProperties['URL']+'\n' # TOC: type | url

                    srcList.append((sv.serviceProperties['URL'],grpName,sv.name,sv.serviceProperties['ServiceType']))
                    
                    # List projection
                    try:
                        prjName = arcpy.Describe(sv).spatialReference.name
                        print arcpy.AddMessage('('+prjName+')')
                        message += '('+prjName+').\n'
                    except:
                        pass
        
                    # List scale depedency/range
                    try:
                        scaleText = getScaleRange(sv)
                        if scaleText <> '':
                            print arcpy.AddMessage('(SCALE RANGE - '+scaleText+')')
                            message += '(SCALE RANGE - '+scaleText+')\n'
                    except:
                        pass

                    # List label
                    try:
                        lblText = getLabel(sv)
                        if lblText <> '':
                            print arcpy.AddMessage('(LABEL - '+lblText+')')
                            message += '(LABEL - '+lblText+').\n'
                    except:
                        pass

                    # List definition query
                    if sv.supports("DEFINITIONQUERY"):
                        defString = sv.definitionQuery
                        if defString <> '':
                            print arcpy.AddMessage('(DEFINITION QUERY - '+defString+')')
                            message += '(DEFINITION QUERY - '+defString+').\n'

                    grpName = ''

                ##########################
                # List ESRI Basemap
                #url = 'http://goto.arcgisonline.com'
                urlString = 'arcgisonline.com'
                bmList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.description.find(urlString) > -1]
                if len(bmList) > 0:
                    print arcpy.AddMessage('='*20)
                    print arcpy.AddMessage('* ESRI Basemap:')
                    message += '='*20+'\n'
                    message += '* ESRI Basemap:\n'
                grpName = tmpGrpName = ''
                for bm in bmList:
                    cnt = len(bm.longName.split('\\'))
                    if cnt > 1:
                        for i in range(1,cnt):
                            grpName += bm.longName.split('\\')[i-1]+' | '
                        grpName = grpName[:-3]
                        if grpName <> tmpGrpName:
                            print arcpy.AddMessage('-'*12+'\nGROUP: '+grpName)
                            message += '-'*12+'\nGROUP: '+grpName+'\n'
                            tmpGrpName = grpName
                    else:
                        print arcpy.AddMessage('-'*12)
                        message += '-'*12+'\n'
                        
                    if not bm.isGroupLayer:
                        print arcpy.AddMessage('- '+bm.name+': '+lyr.description[lyr.description.find('http:'):]) # TOC: url
                        message += '- '+bm.name+': '+lyr.description[lyr.description.find('http:'):]+'\n' # TOC: url

                    srcList.append((lyr.description[lyr.description.find('http:'):],grpName,bm.name,''))
                    
                    # List projection
                    try:
                        prjName = arcpy.Describe(bm).spatialReference.name
                        print arcpy.AddMessage('('+prjName+')')
                        message += '('+prjName+').\n'
                    except:
                        pass
        
                    # List scale depedency/range
                    try:
                        scaleText = getScaleRange(bm)
                        if scaleText <> '':
                            print arcpy.AddMessage('(SCALE RANGE - '+scaleText+')')
                            message += '(SCALE RANGE - '+scaleText+')\n'
                    except:
                        pass

                    # List label
                    try:
                        lblText = getLabel(bm)
                        if lblText <> '':
                            print arcpy.AddMessage('(LABEL - '+lblText+')')
                            message += '(LABEL - '+lblText+').\n'
                    except:
                        pass

                    # List definition query
                    if bm.supports("DEFINITIONQUERY"):
                        defString = bm.definitionQuery
                        if defString <> '':
                            print arcpy.AddMessage('(DEFINITION QUERY - '+defString+')')
                            message += '(DEFINITION QUERY - '+defString+').\n'

                    grpName = ''
                    
                ##########################
                # List table
                tblList = [tbl for tbl in arcpy.mapping.ListTableViews(mxd, data_frame=df)]
                if len(tblList) > 0:
                    print arcpy.AddMessage('='*20)
                    print arcpy.AddMessage('* Table:')
                    message += '='*20+'\n'
                    message += '* Table:\n'
                    print arcpy.AddMessage('-'*12)
                    message += '-'*12+'\n'

                    for tbl in tblList:
                        print arcpy.AddMessage('- '+tbl.name+': '+tbl.workspacePath+' | '+tbl.datasetName) # TOC: path | table
                        message += '- '+tbl.name+': '+tbl.workspacePath+' | '+tbl.datasetName+'\n' # TOC: path | table
                        srcList.append((tbl.workspacePath,'',tbl.name,tbl.datasetName)) # [(PATH,GROUP,TOC,name),(...)]

                ##########################
                # List broken layers
                brokeList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isBroken]
                if len(brokeList) > 0:
                    print arcpy.AddMessage('x'*40)
                    print arcpy.AddMessage('* BROKEN LAYERS:')
                    message += '\n'+'x'*40+'\n'
                    message += '* BROKEN LAYERS:\n'
                grpName = tmpGrpName = ''
                for broke in brokeList:
                    #svName = 'ArcGIS Service: '
                    cnt = len(broke.longName.split('\\'))
                    if cnt > 1:
                        for i in range(1,cnt):
                            grpName += broke.longName.split('\\')[i-1]+' | '
                        grpName = grpName[:-3]
                        if grpName <> tmpGrpName:
                            print arcpy.AddMessage('-'*12+'\nGROUP: '+grpName)
                            message += '-'*12+'\nGROUP: '+grpName+'\n'
                            tmpGrpName = grpName
                    else:
                        print arcpy.AddMessage('-'*12)
                        message += '-'*12+'\n'
                    print arcpy.AddMessage('- '+broke.name)
                    message += '- '+broke.name+'\n'
                    #svName += sv.longName.split('\\')[-1]

                    grpName = ''
                
                ##########################
                # List data source
                #print srcList # [(PATH,GROUP,TOC,name),(...)]
                sortSrcList = sorted(srcList, key=itemgetter(0,1,2,3)) # sort srcList
                if len(sortSrcList) > 0:
                    print arcpy.AddMessage('#'*40)
                    print arcpy.AddMessage('# DATA SOURCE: #')
                    message += '\n'+'#'*40+'\n'
                    message += '# DATA SOURCE: #\n'
                tmpPath=''
                for row in sortSrcList:
                    path=row[0]
                    group=row[1]
                    toc=row[2]
                    name=row[3]
                    if path<>tmpPath:
                        print arcpy.AddMessage('-'*12)
                        print arcpy.AddMessage('* SOURCE: "'+path+'"')
                        message += '-'*12+'\n'
                        message += '* SOURCE: "'+path+'"\n'
                    if group<>'':
                        print arcpy.AddMessage('- GROUP: '+group+' | '+toc+' | '+name)
                        message += '- GROUP: '+group+' | '+toc+' | '+name+'\n'
                    else:
                        print arcpy.AddMessage('- '+toc+' | '+name)
                        message += '- '+toc+' | '+name+'\n'
                    tmpPath=path

                ##########################
                # List rendering time
                # create script file
                renderSCR = open(join(tempfile.gettempdir(),'render.scr'), 'w')
                renderSCR.writelines('StartRendering\nStopRendering')
                #print renderSCR.name
                renderSCR.close()
                # create log file
                logFile = join(tempfile.gettempdir(),'benchmarkLOG.txt')
                os.system(r"C:\PerfQAnalyzer\bin\perfqanalyzer.exe"+' "'+mxdFullName+'" /scr:"'+renderSCR.name+'" /log:"'+logFile+'" /seconds')
                ##C:\PerfQAnalyzer\bin\perfqanalyzer.exe test.mxd" /scr:"..rendering.txt" /log:"\\anadarko.com\world\Temporary(30Days)\Houston\jyl\log.txt" /seconds /reps:5
                #print arcpy.AddMessage(logFile)
                renderingTime = file(logFile, "r").readlines()[-2].strip('Script Execution Time: ') # get time in seconds
                print arcpy.AddMessage('#'*40)
                print arcpy.AddMessage('# RENDERING TIME: (hr:min:sec) #')
                print arcpy.AddMessage(renderingTime)
                message += '\n'+'#'*40+'\n'
                message += '# RENDERING TIME: (hr:min:sec) #\n'
                message += renderingTime+'\n' 
                                               
            # Send out mail

            endTime = time.strftime('%x %X')
            #endTuple = time.strptime(endTime, "%m/%d/%Y %I:%M:%S %p")
            try: 
                endTuple = time.strptime(endTime, "%m/%d/%Y %I:%M:%S %p")
            except:
                endTuple = time.strptime(endTime, "%m/%d/%y %H:%M:%S")
            diffTime = time.mktime(endTuple) - time.mktime(startTuple)
            totalHr = int(diffTime/3600.0)
            if totalHr > 0:
                totalTime += str(totalHr) + ' Hr '
            totalMin = int(diffTime%3600.0)/60
            if totalMin > 0:
                totalTime += str(totalMin) + ' Min '
            totalSec = int(diffTime%3600.0)%60
            if totalSec > 0:
                totalTime += str(totalSec) + ' Sec'
            message += '\n-- End time: ' + endTime + '\n\n'
            message += totalTime+'.'
            if errorMessage <> '':
                message += '='*12+'\n'
                message += '* Problem mxd:\n'
                message += errorMessage
            jTool.mail(sender, to, subject, message.encode('utf-8'))
            ##
            print arcpy.AddMessage('-'*20)
            print arcpy.AddMessage(' Sent E-mail log!')
            print arcpy.AddMessage('-'*20)

            del mxd

