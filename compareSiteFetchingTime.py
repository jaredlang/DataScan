# -*- coding: utf-8 -*-
# compareSite.py
# - compare fetching time of local site, ArcSDE in Houston, and ArcSDE in Denver (temp from DR)
# - 8-24-15 (jyl)

import arcpy, os, sys, getpass, time, tempfile, jTool, re, locale
from arcpy import env
from os.path import join
from operator import itemgetter
from collections import defaultdict

env.overwriteOutput = True
userID = getpass.getuser()

# Set up mail properties
sender = 'jia.liu@anadarko.com'
to = [userID+'@anadarko.com', 'jia.liu@anadarko.com']
subject = 'List: Compare Site Performance'
startTime = time.strftime('%x %X')
try: 
    startTuple = time.strptime(startTime, "%m/%d/%Y %I:%M:%S %p")
except:
    startTuple = time.strptime(startTime, "%m/%d/%y %H:%M:%S")
message = '---- Compare site performance log (in msec) ----' + '\n\n'
message += '-- Starting time: ' + time.strftime('%x %X') + '\n\n'
totalTime = 'Total processing time: '

sdeUser = os.environ['USERNAME']
sdePass = sdeUser
###print os.environ['COMPUTERNAME']

#loc = 'Liberal'
#loc = 'DJBasin84'
loc = arcpy.GetParameterAsText(0)

sdepSDE = jTool.checkCreateDBConn('ORACLE', 'HOUSDEP03', sdeUser, sdePass, 'SDEP', 'SDE', 0)
sdepWORK = jTool.checkCreateDBConn('ORACLE', 'HOUSDEP03', sdeUser, sdePass, 'SDEP', 'WORK', 0)
sde2pSDE = jTool.checkCreateSDEConn('HOUSDEP02', 'SDE2P', '5155', sdeUser, sdePass, 'SDE')
qlspSDE = jTool.checkCreateDBConn('ORACLE', 'HOUSDEP03', sdeUser, sdePass, 'QLSP', 'SDE', 0)
# temp Denver SDEP
dnvSDE = jTool.checkCreateDBConn('ORACLE', 'DRSHOUSDEP03', sdeUser, sdePass, 'DENSDEP', 'SDE', 0)
dnvWORK = jTool.checkCreateDBConn('ORACLE', 'DRSHOUSDEP03', sdeUser, sdePass, 'DENSDEP', 'WORK', 0)

# Get list of locations
##locList = sorted(list(set([row[0] for row in arcpy.da.SearchCursor(join(sdepWORK,'APC.APC2GRP_MASTER'), 'LOC', """STAGE_DIR is not null and
##                                                            DES_DIR not in ('ArcSDE', 'SAME')""")])))

locGDB = [join(row[0], row[1]) for row in arcpy.da.SearchCursor(join(sdepWORK,'APC.APC2GRP_MASTER'),
                                                                       ["DES_DIR", "GDB_NAME"], """LOC = '"""+loc+"""'""")][0] # local gdb full name

# Get extent
aoiExtent = ''
for row in arcpy.da.SearchCursor(join(sdepWORK,'APC.APC2GRP_MASTER'), ["CLIP_BASE", "CLIP_STATE", "CLIP_SOURCE"], """LOC = '"""+loc+"""'"""):
    x1 = x2 = y1 = y2 = xQ = yQ = 0.0
    if loc <> 'Denver': # not Denver
        if not row[1]: # clip from AOI
            fc = join(row[2],row[0])
            aoi = arcpy.Describe(fc).extent
        else: # clip from COUNTY
            fc = row[0]
            #MakeFeatureLayer_management (in_features, out_layer, {where_clause}, {workspace}, {field_info})
            #print row[1]
            arcpy.MakeFeatureLayer_management(join(sdepSDE,fc), "in_memory\\tmpAOI", row[1]) # from county
            arcpy.Dissolve_management("in_memory\\tmpAOI", "in_memory\\tmpAOIDissolve") # need to dissolve to get right extent
            aoi = arcpy.Describe("in_memory\\tmpAOIDissolve").extent
        x1 = aoi.XMin
        x2 = aoi.XMax
        y1 = aoi.YMin
        y2 = aoi.YMax
        #aoiExtent = (str(x1+xQ)+', '+str(y1+yQ)+', '+str(x2-xQ)+', '+str(y2-yQ)) # make extent 1/3 smaller
        #aoiExtent = (str(aoi.XMin)+', '+str(aoi.YMin)+', '+str(aoi.XMax)+', '+str(aoi.YMax))
    #print fc
    else:
        #aoiExtent = '-124, 24, -70, 50' # a random appromimate box for USA <- still too big
        #aoiExtent = '-112, 39, -102, 42.2' # around landgrant <- still too big
        x1 = -112
        x2 = -102
        y1 = 39
        y2 = 42.2
    xQ = (x2-x1)/3.0
    yQ = (y2-y1)/3.0
    aoiExtent = (str(x1+xQ)+', '+str(y1+yQ)+', '+str(x2-xQ)+', '+str(y2-yQ)) # make extent original 1/3

###############
# Loop through various sites
dictData = defaultdict(list) # create empty dictionary

for x in (loc, 'HOU-SDE', 'DNV-SDE'):
#for x in ('HOU-SDE'):
    ##########################
    # List rendering time
    # create script file
    renderSCR = open(join(tempfile.gettempdir(),'render.scr'), 'w')

    # Replicated list
    locSrc = [[row[0], row[1], row[2], row[3], row[4]] for row in arcpy.da.SearchCursor(join(sdepWORK,'APC.APC2GRP_LAYER'),
                                                                   ["SRC_SID", "SRC_PORT", "SRC_DATABASE", "SRC_DATASET", "SRC_LAYER"], """LOC = '"""+loc+"""'""")]
    sortedLocSrc = sorted(locSrc, key=itemgetter(0,1,2,3,4))
    tmpSrv = ''
    cnt = 0
    for row in sortedLocSrc:
        #print row
        if x == loc: # File GDB
            if cnt == 0:
                renderSCR.writelines('Workspace DATABASE="'+locGDB+'"\n')
                print arcpy.AddMessage('\n* Workspace DATABASE="'+locGDB+'"')
                print arcpy.AddMessage('Fetch '+aoiExtent)
        else:
            if row[0]+row[2] <> tmpSrv:
                if x == 'HOU-SDE':
                    renderSCR.writelines('Workspace SERVER='+row[0]+'; DATABASE="'+row[2]+'"; INSTANCE="'+row[1]+':'+row[2]+'"; VERSION='+row[2]+'.DEFAULT; USER='+sdeUser+'; PASSWORD='+sdePass+';\n')
                    print arcpy.AddMessage('\n* Workspace SERVER='+row[0]+'; DATABASE="'+row[2]+'"; INSTANCE="'+row[1]+':'+row[2]+'"; VERSION='+row[2]+'.DEFAULT; USER='+sdeUser+'; PASSWORD='+sdePass+';')
                    print arcpy.AddMessage('Fetch '+aoiExtent)
                else: # Denver ArcSDE
                    renderSCR.writelines('Workspace SERVER=DRSHOUSDEP03; DATABASE="DENSDEP"; INSTANCE="sde:oracle11g:densdep:'+row[2]+'"; VERSION='+row[2]+'.DEFAULT; USER='+sdeUser+'; PASSWORD='+sdePass+';\n')
                    print arcpy.AddMessage('\n* Workspace SERVER=DRSHOUSDEP03; DATABASE="DENSDEP"; INSTANCE="sde:oracle11g:densdep:'+row[2]+'"; VERSION='+row[2]+'.DEFAULT; USER='+sdeUser+'; PASSWORD='+sdePass+';')
                    print arcpy.AddMessage('Fetch '+aoiExtent)
            tmpSrv = row[0]+row[2]
        if x == loc: # File GDB
            renderSCR.writelines('FeatureClass '+row[4].replace('.','_')+'\n')
            #print arcpy.AddMessage('FeatureClass '+row[4].replace('.','_'))
        else:
            renderSCR.writelines('FeatureClass '+row[4]+'\n')
            #print arcpy.AddMessage('FeatureClass '+row[4])
        renderSCR.writelines('Fetch '+aoiExtent+'\n')
        cnt += 1
    renderSCR.close()

    # create log file
    logFile = join(tempfile.gettempdir(),'benchmarkLOG.txt')
    os.system(r"C:\PerfQAnalyzer\bin\perfqanalyzer.exe"+' /scr:"'+renderSCR.name+'" /log:"'+logFile+'" /seconds')
    print arcpy.AddMessage('\n* Running PerfQAnalyzer on '+x+'. (this may take a while...)')
    ##C:\PerfQAnalyzer\bin>perfqanalyzer.exe /scr:"\\anadarko.com\world\Temporary(30Days)\Houston\jyl\render.scr" /log:"\\anadarko.com\world\Temporary(30Days)\Houston\jyl\log.txt" /seconds

    # parse log file
    with open(logFile, "r") as infile:
        for line in infile:
            if ('Fetch found' in line): # and (fc in line)
                FC = line[line.find('"',line.find('"')+1)+3:line.find('feature(s)')-2]
                mTime = line[line.find(': "')+3:-2] # parse between ': "' and '"'
                #print FC+': '+mTime
                #dictData[FC].append(mTime)
                dictData[x].append((FC, mTime))
    print arcpy.AddMessage(dictData[x])

message += '%-40s%20s%20s%20s' % ('', loc, 'HOU-SDE', 'DNV-SDE') + '\n'
message += '-'*105+'\n'

fcList = [row[0] for row in arcpy.da.SearchCursor(join(sdepWORK,'APC.APC2GRP_LAYER'),["SRC_LAYER"], """LOC = '"""+loc+"""'""")] # fc
sortedfcList = sorted(fcList, key=itemgetter(0))

totalTime1 = totalTime2 = totalTime3 = 0.0
for f in sortedfcList:
    time1 = time2 = time3 = '<n/a>'
    if f.replace('.','_') in str(dictData[loc]):
        for i in range(len(dictData[loc])):
            if dictData[loc][i][0] == f.replace('.','_'):
                floatTime1 = float(dictData[loc][i][1])
                totalTime1 += floatTime1
                time1 = locale.format("%.2f", floatTime1, grouping=True) # Convert float to comma-separated string
    if f in str(dictData['HOU-SDE']):
        for i in range(len(dictData['HOU-SDE'])):
            if dictData['HOU-SDE'][i][0] == f:
                floatTime2 = float(dictData['HOU-SDE'][i][1])
                totalTime2 += floatTime2
                time2 = locale.format("%.2f", floatTime2, grouping=True)
    if f in str(dictData['DNV-SDE']):
        for i in range(len(dictData['DNV-SDE'])):
            if dictData['DNV-SDE'][i][0] == f:
                floatTime3 = float(dictData['DNV-SDE'][i][1])
                totalTime3 += floatTime3
                time3 = locale.format("%.2f", floatTime3, grouping=True)
        
    message += '%-40s%20s%20s%20s' % (f, time1, time2, time3)+'\n'
message += '-'*105+'\n'
message += '%-40s%20s%20s%20s' % ('Total Time', \
                                  locale.format("%.2f", totalTime1, grouping=True), \
                                  locale.format("%.2f", totalTime2, grouping=True), \
                                  locale.format("%.2f", totalTime3, grouping=True))+'\n'
#print (message+'\n')

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

jTool.mail(sender, to, subject, message.encode('utf-8'))
print arcpy.AddMessage('- Sent E-mail log!\n')

