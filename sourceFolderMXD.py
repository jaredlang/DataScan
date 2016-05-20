# sourceFolderMXD.py
# - replace data source from SDEP to DENSDEP
# - jyl (7-1-2015)

import arcpy, os, sys, getpass, time, tempfile, jTool, re
from arcpy import env
from os.path import join
from operator import itemgetter

env.overwriteOutput = True
userID = getpass.getuser()

sdeUser = os.environ['USERNAME']
sdePass = sdeUser
tmpDR = tempfile.gettempdir()

#mxd = arcpy.mapping.MapDocument(r"\\anadarko.com\world\Temporary(30Days)\Houston\jyl\MRB_3-Houston.mxd")
prosFolder = arcpy.GetParameterAsText(0)

#loc = 'Houston'
loc = arcpy.GetParameterAsText(1)

# This part is taken care of in the ArcToolbox tool
houSDE = jTool.checkCreateDBConn('ORACLE', 'HOUSDEP03', sdeUser, sdePass, 'SDEP', 'SDE', 0)
houWORK = jTool.checkCreateDBConn('ORACLE', 'HOUSDEP03', sdeUser, sdePass, 'SDEP', 'WORK', 0)

for root, dirs, files in os.walk(prosFolder):
    if not os.path.exists(root):
        errorMessage = '"'+root+'" doesn\'t exist!\n'
        print arcpy.AddMessage('"'+root+'" doesn\'t exist!')
        break

    for file in sorted(files): # Sort file alphabetically
        if file.endswith(".mxd") or file.endswith(".MXD"):
            mxd = arcpy.mapping.MapDocument(join(root,file))
            mxdPrefix = mxd.filePath[:-4] # get the full path minus .mxd or .Mxd or .MXD
            #tmpMxd = mxdPrefix+'-'+loc+'tmp.mxd'
            tmpMxd = join(tmpDR,'ttt.mxd')
            newMxd = mxdPrefix+'-'+loc+'.mxd'

            # Set up mail properties
            sender = 'jia.liu@anadarko.com'
            to = [userID+'@anadarko.com', 'jia.liu@anadarko.com']
            subject = 'redirect: Resource MXD Database Connection'
            startTime = time.strftime('%x %X')
            try: 
                startTuple = time.strptime(startTime, "%m/%d/%Y %I:%M:%S %p")
            except:
                startTuple = time.strptime(startTime, "%m/%d/%y %H:%M:%S")
            message = '---- Redirected mxd database connection log ----' + '\n\n'
            message += '-- Starting time: ' + time.strftime('%x %X') + '\n\n'
            message = errorMessage = ''
            totalTime = 'Total processing time: '
            tmpDR = tempfile.gettempdir()

##            ####################
##            # This section needs to be inserted in ArcToolbox | def updateParameters(self):
##            # Get list of locations
##            locList = sorted(list(set([row[0] for row in arcpy.da.SearchCursor(join(houWORK,'APC.APC2GRP_MASTER'), 'LOC', """STAGE_DIR is not null and
##                                                                        DES_DIR not in ('ArcSDE', 'SAME')""")])))
##            locList.insert(0, 'Houston') # insert Houston to 1st of the list
##            ###################

            if loc <> 'Houston': # Find remote File GDB name and replicated list
                locGDB = [join(row[0], row[1]) for row in arcpy.da.SearchCursor(join(houWORK,'APC.APC2GRP_MASTER'),
                                                                                       ["DES_DIR", "GDB_NAME"], """LOC = '"""+loc+"""'""")][0] # local gdb full name
                        
                locSrcList = list(set([row[0] for row in arcpy.da.SearchCursor(join(houWORK,'APC.APC2GRP_LAYER'),
                                                                                       'SRC_LAYER', """LOC = '"""+loc+"""'""")])) # locally replicated list

            def switchPath(loc, path, lyr, version, sid):
                try: # Switch to Houston standards - no matter the destination
                    #if 'HOUSDEP03' in path.upper() or 'SDEP' in path.upper(): # only SDEP has versioning
                    if 'HOUSDEP03' in path.upper() or 'SDEP' in path.upper() or 'SDEP' in sid.upper(): # only SDEP has versioning
                        #if version <> '':
                        if version <> '' and version not in ('SDE.DEFAULT', 'WORK.DEFAULT'):
                            versionSDE = jTool.checkCreateDBversion('SDEP', sdeUser, sdePass, 'SDE', version)
                            lyr.replaceDataSource(versionSDE, 'SDE_WORKSPACE', lyr.datasetName, 'FALSE')
                            returnSDE = versionSDE
                        elif 'WORK' not in path.upper(): # SDE.DEFAULT
                            lyr.replaceDataSource(houSDE, 'SDE_WORKSPACE', lyr.datasetName, 'FALSE') # SDE_WORKSPACE, FILEGDB_WORKSPACE, etc.
                            returnSDE = houSDE
                        else: # WORK.DEFAULT
                            lyr.replaceDataSource(houWORK, 'SDE_WORKSPACE', lyr.datasetName, 'FALSE')
                            returnSDE = houWORK
                    if 'QLSP' in path.upper() or 'QLSP' in sid.upper():
                        qlsSDE = jTool.checkCreateDBConn('ORACLE', 'HOUSDEP03', sdeUser, sdePass, 'QLSP', 'SDE', 0)
                        lyr.replaceDataSource(qlsSDE, 'SDE_WORKSPACE', lyr.datasetName, 'FALSE') # SDE_WORKSPACE, FILEGDB_WORKSPACE, etc.
                        returnSDE = qlsSDE
                    if 'HOUSDEP02' in path.upper() or 'SDE2P' in path.upper() or 'SDE2P' in sid.upper():
                        houSDE2 = jTool.checkCreateDBConn('ORACLE', 'HOUSDEP02', sdeUser, sdePass, '5155', 'SDE', 0)
                        lyr.replaceDataSource(houSDE2, 'SDE_WORKSPACE', lyr.datasetName, 'FALSE') # SDE_WORKSPACE, FILEGDB_WORKSPACE, etc.            
                        returnSDE = houSDE2
                except:
                    returnSDE = '0'

                ##############
                if lyr.datasetName.startswith('GRP_'): # check if it's GRP_
                    srcLyr = lyr.datasetName.replace('_','.',2).replace('.','_',1)
                else:
                    srcLyr = lyr.datasetName.replace('_','.',1)
                if lyr.datasetName.startswith('HARD_ROCK_LSE_'): # check if it's HARD_ROCK_LSE
                    srcLyr = lyr.datasetName.replace('HARD_ROCK_LSE_','HARD_ROCK_LSE.')
                if lyr.datasetName.startswith('GRANTED_TO_OTHER_'): # check if it's GRANTED_TO_OTHER
                    srcLyr = lyr.datasetName.replace('GRANTED_TO_OTHER_','GRANTED_TO_OTHER.')
                ###############

                if loc == 'Houston': # replicated file gdb source and switch to SDE
                    if 'HOUSTON2' in path.upper() and '.GDB' in path.upper(): # replicated file gdb source
                        sdeRow = [row for row in arcpy.da.SearchCursor(join(houWORK,'APC.APC2GRP_LAYER'), \
                                                                       ["SRC_SRVR", "SRC_SID", "SRC_DATABASE"], """SRC_LAYER = '"""+srcLyr+"""'""")][0] 
                        returnSDE = jTool.checkCreateDBConn('ORACLE', sdeRow[0], sdeUser, sdePass, sdeRow[1], sdeRow[2], 0)
                        lyr.replaceDataSource(returnSDE, 'SDE_WORKSPACE', srcLyr, 'FALSE') # SDE_WORKSPACE, FILEGDB_WORKSPACE, etc. 
                    
                if loc <> 'Houston': # Non-HOU File GDB
                    if lyr.datasetName in locSrcList or srcLyr in locSrcList: # source is SDE or file gdb
                        lyr.replaceDataSource(locGDB, 'FILEGDB_WORKSPACE', lyr.datasetName.replace('.','_'), 'FALSE')
                        returnSDE = locGDB
                        
                if returnSDE <> '0' and returnSDE <> path:
                    return returnSDE
                else:
                    return '0'
                    

            cntReplace = 0
            # Find active dataframe
            dfList = arcpy.mapping.ListDataFrames(mxd)
            message += 'Input: "'+mxd.filePath+'"\n'
            message += 'Redirect to: "'+loc+'"...\n'
            for df in dfList:
                print arcpy.AddMessage('\n'+'*'*50)
                print arcpy.AddMessage('** Data Frame: '+df.name)
                message += '\n'+'*'*50+'\n'
                message += '** Data Frame: '+df.name+'...\n'
                srcList = []
                version = ''

                # List fc
                fcList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isFeatureLayer]
                grpName = ''
                for fc in fcList:
                    cnt = len(fc.longName.split('\\'))
                    if cnt > 1:
                        for i in range(1,cnt):
                            grpName += fc.longName.split('\\')[i-1]+' | '
                        grpName = grpName[:-3]
                    if fc.supports('serviceProperties'): #???
                        try:
                            db = fc.serviceProperties['Db_Connection_Properties'] # Direct Connect and 3-tier Connect
                        except:
                            db = fc.serviceProperties['Server'] # 3-tier Connect?
                        version = fc.serviceProperties['Version']
                        # feature class
                        if fc.workspacePath <> '':
                            srcList.append((fc.workspacePath,grpName,fc.name,fc.datasetName,db))
                        else:
                            srcList.append((db.upper()+' | '+version,grpName,fc.name,fc.datasetName,db))
                    else: # ?
                        if fc.workspacePath <> '': 
                            srcList.append((fc.workspacePath,grpName,fc.name,fc.datasetName,''))
                    grpName = ''
                
                ##########################
                # List table
                tblList = [tbl for tbl in arcpy.mapping.ListTableViews(mxd, data_frame=df)]
                for tbl in tblList:
                    #if tbl.workspacePath <> '' and bool(re.match(tbl.workspacePath[-4:],'.SDE', re.I)):
                    if tbl.workspacePath <> '':
                        srcList.append((tbl.workspacePath,'',tbl.name,tbl.datasetName,'')) # [(PATH,GROUP,TOC,name),(...)]

                ##########################
                # List data source
                newSrcList = []
                badSrcList = []
                #print srcList # [(PATH,GROUP,TOC,name),(...)]
                sortSrcList = sorted(srcList, key=itemgetter(0,1,2,3,4)) # sort srcList
                if len(sortSrcList) > 0:
                    print arcpy.AddMessage('#'*40)
                    print arcpy.AddMessage('# DATA SOURCE: #')
                    message += '\n'+'#'*40+'\n'
                    message += '# ORIGINAL DATA SOURCE: #\n'
                #tmpPath=tmpGroup=''
                tmpPath=''

                for row in sortSrcList:
                    newPath = '0'
                    path=row[0]
                    group=row[1]
                    toc=row[2]
                    name=row[3]
                    sid=row[4]

                    #sys.exit()
                    try: # switch to new path
                        lyr = arcpy.mapping.ListLayers(mxd, toc, df)[0] # get lyr object
                        newPath = switchPath(loc, path, lyr, version, sid)
                        #if newPath > '0':
                        if newPath > '0' and toc  <> '': # toc <> '' to deal with null in toc
                            cnt += 1
                            newSrcList.append((newPath,group,toc,name,sid))
                        else:
                            badSrcList.append((path,group,toc,name,sid)) #  couldn't switch <- like credential access problem
                    except:
                        newPath = '0' # couldn't switch <- like table
                        badSrcList.append((path,group,toc,name,sid))

                    if path<>tmpPath:
                        print arcpy.AddMessage('-'*12)
                        print arcpy.AddMessage('* ORIGINAL SOURCE: "'+path+'"')
                        message += '-'*12+'\n'
                        message += '* ORIGINAL SOURCE: "'+path+'"\n'
                    if group<>'':
                        print arcpy.AddMessage('- GROUP: '+group+' | '+toc+' | '+name)
                        message += '- GROUP: '+group+' | '+toc+' | '+name+'\n'
                    else:
                        print arcpy.AddMessage('- '+toc+' | '+name)
                        message += '- '+toc+' | '+name+'\n'
                    tmpPath=path

                ##########################
                # List switched data source
                cntReplace += len(newSrcList)
                if cntReplace > 0:
                    sortNewSrcList = sorted(newSrcList, key=itemgetter(0,1,2,3,4)) # sort newSrcList
                    print arcpy.AddMessage('#'*40)
                    print arcpy.AddMessage('# NEW DATA SOURCE: #')
                    message += '\n'+'#'*40+'\n'
                    message += '# NEW DATA SOURCE: #\n'
                    
                    tmpPath=''
                    for row in sortNewSrcList:
                        path=row[0]
                        group=row[1]
                        toc=row[2]
                        if loc == 'Houston':
                            name=row[3]
                        else:
                            name=row[3].replace('.','_')
                        
                        if path<>tmpPath:
                            print arcpy.AddMessage('-'*12)
                            print arcpy.AddMessage('* NEW SOURCE: "'+path+'"')
                            message += '-'*12+'\n'
                            message += '* NEW SOURCE: "'+path+'"\n'

                        #if group<>tmpGroup and group<>'':
                        if group<>'':
                            print arcpy.AddMessage('- GROUP: '+group+' | '+toc+' | '+name)
                            message += '- GROUP: '+group+' | '+toc+' | '+name+'\n'
                        else:
                            print arcpy.AddMessage('- '+toc+' | '+name)
                            message += '- '+toc+' | '+name+'\n'
                        tmpPath=path

                ##########################
                # List bad data source
                if len(badSrcList) > 0:
                    sortBadSrcList = sorted(badSrcList, key=itemgetter(0,1,2,3,4)) # sort badSrcList
                    print arcpy.AddMessage('#'*40)
                    print arcpy.AddMessage('# DATA SOURCE - did not switch: #')
                    message += '\n'+'#'*40+'\n'
                    message += '# DATA SOURCE - did not switch: #\n'
                    tmpPath=''

                    for row in sortBadSrcList:
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
                
            if cntReplace > 0:
                if os.path.exists(newMxd):
                    try:
                        os.remove(newMxd)
                    except:
                        print arcpy.AddMessage('"'+newMxd+'" is open. Please close then rerun the script.')
                        errorMessage += '"'+newMxd+'" is open. Please close then rerun the script.\n'
                        break
##                if os.path.exists(tmpMxd):
##                    try:
##                        os.remove(tmpMxd)
##                    except:
##                        print arcpy.AddMessage('"'+tmpMxd+'" is open. Please close then rerun the script.')
##                        errorMessage += '"'+tmpMxd+'" is open. Please close then rerun the script.\n'
##                        break
##                        
##                mxd.saveACopy(tmpMxd)
##                del mxd
##                mxd = arcpy.mapping.MapDocument(tmpMxd)
                mxd.saveACopy(newMxd)
                #mxd.saveACopy(newMxd, '10.0')
                del mxd
                #os.remove(tmpMxd) # clean up tmpMxd
                mxd = arcpy.mapping.MapDocument(newMxd)

            ##########################
            # List broken layers
            # Find active dataframe
            dfList = arcpy.mapping.ListDataFrames(mxd)
            #message += 'Input: "'+mxd.filePath+'"\n'
            for df in dfList:
                brokeList = [lyr for lyr in arcpy.mapping.ListLayers(mxd, data_frame=df) if lyr.isBroken]
                #brknList = arcpy.mapping.ListBrokenDataSources(mxd)
                if len(brokeList) > 0:
                    print arcpy.AddMessage('x'*40)
                    print arcpy.AddMessage('* NEW BROKEN LAYERS:')
                    message += '\n'+'x'*40+'\n'
                    message += '* NEW BROKEN LAYERS:\n'
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

                    grpName = ''

            del mxd

            # Send out mail
            if cntReplace > 0:
                message += '\nOutput new mxd is: "'+newMxd+'"\n\n'
            else:
                message += '\n** No layer was redirected. **\n\n'
            endTime = time.strftime('%x %X')
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
            message += '-- End time: ' + endTime + '\n\n'
            message += totalTime
            if errorMessage <> '':
                message += '='*12+'\n'
                #message += '* Problem layer(s):\n'
                message += errorMessage+'\n'
            jTool.mail(sender, to, subject, message.encode('utf-8'))

            print arcpy.AddMessage('- Sent E-mail log!\n')

