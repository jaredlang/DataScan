import os
import tempfile
import arcpy
import xml.etree.ElementTree as ET

AGS_HOME = arcpy.GetInstallInfo("Desktop")["InstallDir"]
METADATA_TRANSLATOR = os.path.join(AGS_HOME, r'Metadata/Translator/ARCGIS2FGDC.xml')

def update_sde_metadata(sdeFC, srcFC):

    TEMP_DIR = tempfile.gettempdir()
    metadataFile = os.path.join(TEMP_DIR, os.path.basename(sdeFC) + '-metadata.xml')
    migrationText = "*** Migrated from the L Drive (%s)" % srcFC

    if os.path.exists(metadataFile):
        os.remove(metadataFile)

    # A- export the medata from SDE feature class
    print 'exporting the metadata of %s to %s' % (sdeFC, metadataFile)
    arcpy.ExportMetadata_conversion(
    	Source_Metadata=sdeFC,
    	Translator=METADATA_TRANSLATOR,
    	Output_File=metadataFile
    )

    # B- modify metadata
    print 'modifying the metadata file [%s]' % (metadataFile)
    tree = ET.parse(metadataFile)
    root = tree.getroot()
    idinfo = root.find('idinfo')
    dspt = idinfo.find('descript')
    # B1- add the element
    if dspt is None:
        dspt = ET.SubElement(idinfo, 'descript')
        ET.SubElement(dspt, 'abstract')
    else:
        abstract = dspt.find('abstract')
        if abstract is None:
            ET.SubElement(dspt, 'abstract')
    # B2- modify the element text
    abstract = dspt.find('abstract')
    if abstract.text is None:
        abstract.text = migrationText
    elif abstract.text.find(migrationText) == -1:
        abstract.text = abstract.text + migrationText

    tree.write(metadataFile)

    # C- import the modified metadata back to SDE feature class
    print 'importing the metadata file [%s] to %s' % (metadataFile, sdeFC)
    arcpy.ImportMetadata_conversion(
    	Source_Metadata=metadataFile,
    	Import_Type="FROM_FGDC",
    	Target_Metadata=sdeFC,
    	Enable_automatic_updates="ENABLED"
    )

    print 'The metadata of %s is updated' % sdeFC


if __name__ == "__main__":
    update_sde_metadata(r'C:\Users\kdb086\Documents\ArcGIS\SDE2T_MOZ_GIS.sde\MOZGIS.BND_DU_Build_Area_F', 'some shapefile')
    update_sde_metadata(r'C:\Users\kdb086\Documents\ArcGIS\SDE2T_MOZ_GIS.sde\MOZGIS.BND_AirStrip_Coord_20160818', 'some feature class in gdb')

