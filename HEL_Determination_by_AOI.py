#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      Adolfo.Diaz
#
# Created:     07/06/2017
# Copyright:   (c) Adolfo.Diaz 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------

## ===================================================================================
def AddMsgAndPrint(msg, severity=0):
    # prints message to screen if run as a python script
    # Adds tool message to the geoprocessor
    #
    #Split the message on \n first, so that if it's multiple lines, a GPMessage will be added for each line
    try:
        print msg

        #for string in msg.split('\n'):
            #Add a geoprocessing message (in case this is run as a tool)
        if severity == 0:
            arcpy.AddMessage(msg)

        elif severity == 1:
            arcpy.AddWarning(msg)

        elif severity == 2:
            arcpy.AddError(msg)

    except:
        pass

## ===================================================================================
def errorMsg():

    try:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        theMsg = "\t" + traceback.format_exception(exc_type, exc_value, exc_traceback)[1] + "\n\t" + traceback.format_exception(exc_type, exc_value, exc_traceback)[-1]
        AddMsgAndPrint(theMsg,2)

    except:
        AddMsgAndPrint("Unhandled error in errorMsg method", 2)
        pass

### ===================================================================================
def setScratchWorkspace():
    """ This function will set the scratchWorkspace for the interim of the execution
        of this tool.  The scratchWorkspace is used to set the scratchGDB which is
        where all of the temporary files will be written to.  The path of the user-defined
        scratchWorkspace will be compared to existing paths from the user's system
        variables.  If there is any overlap in directories the scratchWorkspace will
        be set to C:\TEMP, assuming C:\ is the system drive.  If all else fails then
        the packageWorkspace Environment will be set as the scratchWorkspace. This
        function returns the scratchGDB environment which is set upon setting the scratchWorkspace"""

    try:
        AddMsgAndPrint("\nSetting Scratch Workspace")
        scratchWK = arcpy.env.scratchWorkspace

        # -----------------------------------------------
        # Scratch Workspace is defined by user or default is set
        if scratchWK is not None:

            # dictionary of system environmental variables
            envVariables = os.environ

            # get the root system drive
            if envVariables.has_key('SYSTEMDRIVE'):
                sysDrive = envVariables['SYSTEMDRIVE']
            else:
                sysDrive = None

            varsToSearch = ['ESRI_OS_DATADIR_LOCAL_DONOTUSE','ESRI_OS_DIR_DONOTUSE','ESRI_OS_DATADIR_MYDOCUMENTS_DONOTUSE',
                            'ESRI_OS_DATADIR_ROAMING_DONOTUSE','TEMP','LOCALAPPDATA','PROGRAMW6432','COMMONPROGRAMFILES','APPDATA',
                            'USERPROFILE','PUBLIC','SYSTEMROOT','PROGRAMFILES','COMMONPROGRAMFILES(X86)','ALLUSERSPROFILE']

            """ This is a printout of my system environmmental variables - Windows 7
            -----------------------------------------------------------------------------------------
            ESRI_OS_DATADIR_LOCAL_DONOTUSE C:\Users\adolfo.diaz\AppData\Local\
            ESRI_OS_DIR_DONOTUSE C:\Users\ADOLFO~1.DIA\AppData\Local\Temp\6\arc3765\
            ESRI_OS_DATADIR_MYDOCUMENTS_DONOTUSE C:\Users\adolfo.diaz\Documents\
            ESRI_OS_DATADIR_COMMON_DONOTUSE C:\ProgramData\
            ESRI_OS_DATADIR_ROAMING_DONOTUSE C:\Users\adolfo.diaz\AppData\Roaming\
            TEMP C:\Users\ADOLFO~1.DIA\AppData\Local\Temp\6\arc3765\
            LOCALAPPDATA C:\Users\adolfo.diaz\AppData\Local
            PROGRAMW6432 C:\Program Files
            COMMONPROGRAMFILES :  C:\Program Files (x86)\Common Files
            APPDATA C:\Users\adolfo.diaz\AppData\Roaming
            USERPROFILE C:\Users\adolfo.diaz
            PUBLIC C:\Users\Public
            SYSTEMROOT :  C:\Windows
            PROGRAMFILES :  C:\Program Files (x86)
            COMMONPROGRAMFILES(X86) :  C:\Program Files (x86)\Common Files
            ALLUSERSPROFILE :  C:\ProgramData
            ------------------------------------------------------------------------------------------"""

            bSetTempWorkSpace = False

            """ Iterate through each Environmental variable; If the variable is within the 'varsToSearch' list
                list above then check their value against the user-set scratch workspace.  If they have anything
                in common then switch the workspace to something local  """
            for var in envVariables:

                if not var in varsToSearch:
                    continue

                # make a list from the scratch and environmental paths
                varValueList = (envVariables[var].lower()).split(os.sep)          # ['C:', 'Users', 'adolfo.diaz', 'AppData', 'Local']
                scratchWSList = (scratchWK.lower()).split(os.sep)                 # [u'C:', u'Users', u'adolfo.diaz', u'Documents', u'ArcGIS', u'Default.gdb', u'']

                # remove any blanks items from lists
                if '' in varValueList: varValueList.remove('')
                if '' in scratchWSList: scratchWSList.remove('')

                # First element is the drive letter; remove it if they are
                # the same otherwise review the next variable.
                if varValueList[0] == scratchWSList[0]:
                    scratchWSList.remove(scratchWSList[0])
                    varValueList.remove(varValueList[0])

                # obtain a similarity ratio between the 2 lists above
                #sM = SequenceMatcher(None,varValueList,scratchWSList)

                # Compare the values of 2 lists; order is significant
                common = [i for i, j in zip(varValueList, scratchWSList) if i == j]

                if len(common) > 0:
                    bSetTempWorkSpace = True
                    break

            # The current scratch workspace shares 1 or more directory paths with the
            # system env variables.  Create a temp folder at root
            if bSetTempWorkSpace:
                AddMsgAndPrint("\tCurrent Workspace: " + scratchWK,0)

                if sysDrive:
                    tempFolder = sysDrive + os.sep + "TEMP"

                    if not os.path.exists(tempFolder):
                        os.makedirs(tempFolder,mode=777)

                    arcpy.env.scratchWorkspace = tempFolder
                    AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                else:
                    packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
                    if arcpy.env[packageWS[0]]:
                        arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                        AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)
                    else:
                        AddMsgAndPrint("\tCould not set any scratch workspace",2)
                        return False

            # user-set workspace does not violate system paths; Check for read/write
            # permissions; if write permissions are denied then set workspace to TEMP folder
            else:
                arcpy.env.scratchWorkspace = scratchWK

                if arcpy.env.scratchGDB == None:
                    AddMsgAndPrint("\tCurrent scratch workspace: " + scratchWK + " is READ only!",0)

                    if sysDrive:
                        tempFolder = sysDrive + os.sep + "TEMP"

                        if not os.path.exists(tempFolder):
                            os.makedirs(tempFolder,mode=777)

                        arcpy.env.scratchWorkspace = tempFolder
                        AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                    else:
                        packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
                        if arcpy.env[packageWS[0]]:
                            arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                            AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                        else:
                            AddMsgAndPrint("\tCould not set any scratch workspace",2)
                            return False

                else:
                    AddMsgAndPrint("\tUser-defined scratch workspace is set to: "  + arcpy.env.scratchGDB,0)

        # No workspace set (Very odd that it would go in here unless running directly from python)
        else:
            AddMsgAndPrint("\tNo user-defined scratch workspace ",0)
            sysDrive = os.environ['SYSTEMDRIVE']

            if sysDrive:
                tempFolder = sysDrive + os.sep + "TEMP"

                if not os.path.exists(tempFolder):
                    os.makedirs(tempFolder,mode=777)

                arcpy.env.scratchWorkspace = tempFolder
                AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

            else:
                packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
                if arcpy.env[packageWS[0]]:
                    arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                    AddMsgAndPrint("\tTemporarily setting scratch workspace to: " + arcpy.env.scratchGDB,1)

                else:
                    AddMsgAndPrint("\tCould not set scratchWorkspace. Not even to default!",2)
                    return False

        arcpy.Compact_management(arcpy.env.scratchGDB)
        return arcpy.env.scratchGDB

    except:

        # All Failed; set workspace to packageWorkspace environment
        try:
            packageWS = [f for f in arcpy.ListEnvironments() if f=='packageWorkspace']
            if arcpy.env[packageWS[0]]:
                arcpy.env.scratchWorkspace = arcpy.env[packageWS[0]]
                arcpy.Compact_management(arcpy.env.scratchGDB)
                return arcpy.env.scratchGDB
            else:
                AddMsgAndPrint("\tCould not set scratchWorkspace. Not even to default!",2)
                return False
        except:
            errorMsg()
            return False

# ===============================================================================================================
def splitThousands(someNumber):
# will determine where to put a thousands seperator if one is needed.
# Input is an integer.  Integer with or without thousands seperator is returned.

    try:
        return re.sub(r'(\d{3})(?=\d)', r'\1,', str(someNumber)[::-1])[::-1]

    except:
        errorMsg()
        return someNumber

## =============================================== Main Body ====================================================

import sys, string, os, locale, traceback, urllib, re, arcpy, operator, getpass
from arcpy import env
from arcpy.sa import *

if __name__ == '__main__':

    AOI = arcpy.GetParameter(0)
    helLayer = arcpy.GetParameter(1)
    cluLayer = arcpy.GetParameter(2)
    inputDEM = arcpy.GetParameter(3)
    zUnits = arcpy.GetParameterAsText(4)
    kFactorFld = arcpy.GetParameterAsText(5)
    tFactorFld = arcpy.GetParameterAsText(6)
    rFactorFld = arcpy.GetParameterAsText(7)

    helSummary = os.path.dirname(sys.argv[0]) + os.sep + r'HEL.mdb\HEL_Determinations\HELSummaryLayer'
    helYes = os.path.dirname(sys.argv[0]) + os.sep + r'HEL.mdb\HEL_Determinations\HELSummaryLayer'

    """ ---------------------------------------------------------------------------------------------- Routine Shit"""
    # Check Availability of Spatial Analyst Extension
    try:
        if arcpy.CheckExtension("Spatial") == "Available":
            arcpy.CheckOutExtension("Spatial")
        else:
            raise ExitError,"\n\nSpatial Analyst license is unavailable.  May need to turn it on!"

    except LicenseError:
        AddMsgAndPrint("\n\nSpatial Analyst license is unavailable.  May need to turn it on!",2)
        exit()
    except arcpy.ExecuteError:
        AddMsgAndPrint(arcpy.GetMessages(2),2)
        exit()

    # Set overwrite option
    arcpy.env.overwriteOutput = True

    # define and set the scratch workspace
    scratchWS = setScratchWorkspace()
    arcpy.env.scratchWorkspace = scratchWS

    if not scratchWS:
        exit()

    """ ---------------------------------------------------------------------------------------------- Prepare CLU Layer"""
    descAOI = arcpy.Describe(AOI)
    aoiPath = descAOI.catalogPath
    cluLayerPath = arcpy.Describe(cluLayer).catalogPath
    cluAOI = arcpy.CreateScratchName("cluAOI",data_type="FeatureClass",workspace=scratchWS)

    # AOI is digitized and resides in memory; Clip CLU based on the user-digitized polygon.
    if descAOI.dataType.upper() == "FEATURERECORDSETLAYER":
        AddMsgAndPrint("\nClipping the CLU layer to the manually digitized AOI")
        arcpy.Clip_analysis(cluLayer,AOI,cluAOI)

    # AOI is some existing feature.  Not digitized.
    else:
        # AOI came from the CLU
        if aoiPath == cluLayerPath:
            AddMsgAndPrint("\nUsing " + str(int(arcpy.GetCount_management(AOI).getOutput(0))) + " features from the CLU Layer as AOI")
            cluAOI = arcpy.CopyFeatures_management(AOI,cluAOI)
        else:
            AddMsgAndPrint("\nClipping the CLU layer to the manually digitized AOI")
            arcpy.Clip_analysis(cluLayer,AOI,cluAOI)

    # Update Acres - Add Calcacre field if it doesn't exist.
    calcAcreFld = "CALCACRES"
    if not len(arcpy.ListFields(cluAOI,calcAcreFld)) > 0:
        arcpy.AddField_management(cluAOI,calcAcreFld,"DOUBLE")

    arcpy.CalculateField_management(cluAOI,calcAcreFld,"!shape.area@acres!","PYTHON_9.3")
    totalAcres = float("%.1f" % (sum([row[0] for row in arcpy.da.SearchCursor(cluAOI, (calcAcreFld))])))
    AddMsgAndPrint("\tTotal Acres: " + splitThousands(totalAcres))

    """ ---------------------------------------------------------------------------------------------- Check DEM Coordinate System and Linear Units"""

    desc = arcpy.Describe(inputDEM)
    inputDEMPath = desc.catalogPath
    sr = desc.SpatialReference
    units = sr.LinearUnitName
    cellSize = desc.MeanCellWidth

    AddMsgAndPrint("\nGathering information about DEM: \"" + os.path.basename(inputDEMPath) + "\"")

    if units == "Meter":
        units = "Meters"
    elif units == "Foot":
        units = "Feet"
    elif units == "Foot_US":
        units = "Feet"
    else:
        AddMsgAndPrint("\tCould not determine linear units of DEM....Exiting!",2)
        exit()

    # if zUnits were left blank than assume Z-values are the same as XY units.
    if not len(zUnits) > 0:
        zUnits = units

    # Coordinate System must be a Projected type in order to continue.
    # zUnits will determine Zfactor
    # if XY units differ from Z units then a Zfactor must be calculated to adjust
    # the z units by multiplying by the Zfactor

    if sr.Type == "Projected":
        if zUnits == "Meters":
            Zfactor = 3.280839896       # 3.28 feet in a meter

        elif zUnits == "Centimeters":   # 0.033 feet in a centimeter
            Zfactor = 0.0328084

        elif zUnits == "Inches":        # 0.083 feet in an inch
            Zfactor = 0.0833333

        # z units and XY units are the same thus no conversion is required (Feet is assumed here)
        else:
            Zfactor = 1.0

        AddMsgAndPrint("\tProjection Name: " + sr.Name,0)
        AddMsgAndPrint("\tXY Linear Units: " + units,0)
        AddMsgAndPrint("\tElevation Values (Z): " + zUnits,0)
        AddMsgAndPrint("\tCell Size: " + str(desc.MeanCellWidth) + " x " + str(desc.MeanCellHeight) + " " + units,0)

    else:
        AddMsgAndPrint("\n\n\t" + os.path.basename(inputDEM) + " is NOT in a projected Coordinate System....EXITING",2)
        exit()

    """ ---------------------------------------------------------------------------------------------- Buffer CLU (AOI) Layer by 300 Meters"""
    AddMsgAndPrint("\nBuffering AOI by 300 Meters")
    cluBuffer = arcpy.CreateScratchName("cluBuffer",data_type="FeatureClass",workspace=scratchWS)
    arcpy.Buffer_analysis(cluAOI,cluBuffer,"300 Meters","FULL","ROUND")

    """ ---------------------------------------------------------------------------------------------- Intersect Soils (HEL) with CLU (AOI) and configure"""
    AddMsgAndPrint("\nIntersecting Soils and AOI")
    aoiCluIntersect = arcpy.CreateScratchName("aoiCLUIntersect",data_type="FeatureClass",workspace=scratchWS)
    arcpy.Intersect_analysis([cluAOI,helLayer],aoiCluIntersect,"ALL")

    HELvalueFld = "HELValue"
    if not len(arcpy.ListFields(aoiCluIntersect,HELvalueFld)) > 0:
        AddMsgAndPrint("\tAdding 'HELValue' field")
        arcpy.AddField_management(aoiCluIntersect,HELvalueFld,"SHORT")

    """ ---------------------------------------------------------------------------------------------- Extract DEM using CLU layer"""
    AddMsgAndPrint("\nExtracting DEM subset using buffered CLU layer")
    demExtract = arcpy.CreateScratchName("demExtract",data_type="RasterDataset",workspace=scratchWS)
    outExtractMask = ExtractByMask(inputDEM,cluAOI)
    outExtractMask.save(demExtract)

    """----------------------------------------------------------------------------------------------  Create Slope Layer"""
    AddMsgAndPrint("\nCreating Slope Derivative")
    slope = arcpy.CreateScratchName("outSlope",data_type="RasterDataset",workspace=scratchWS)
    outSlope = Slope(demExtract,"PERCENT_RISE",Zfactor)
    outSlope.save(slope)

    """---------------------------------------------------------------------------------------------- Create Flow Direction and Flow Length"""
    AddMsgAndPrint("\nCalculating Flow Direction")
    flowDirection = arcpy.CreateScratchName("flowDirection",data_type="RasterDataset",workspace=scratchWS)
    outFlowDirection = FlowDirection(demExtract, "FORCE")
    outFlowDirection.save(flowDirection)

    AddMsgAndPrint("\nCalculating Flow Length")
    flowLength = arcpy.CreateScratchName("flowLength",data_type="RasterDataset",workspace=scratchWS)
    outFlowLength = FlowLength(flowDirection,"UPSTREAM", "")
    outFlowLength.save(flowLength)

    # convert Flow Length distance units to feet if original DEM is not in feet.
    if not units == ("Feet"):
        AddMsgAndPrint("\tConverting Flow Length Distance units to Feet")
        flowLengthFT = arcpy.CreateScratchName("flowLength_FT",data_type="RasterDataset",workspace=scratchWS)
        outflowLengthFT = Raster(flowLength) * Zfactor
        outflowLengthFT.save(flowLengthFT)

    else:
        flowLengthFT = flowLength

    """---------------------------------------------------------------------------------------------- Calculate LS Factor"""
    # Calculate S Factor
    # ((0.065 +( 0.0456 * ("%slope%"))) +( 0.006541 * (Power("%slope%",2))))
    AddMsgAndPrint("\nCalculating S Factor")
    sFactor = arcpy.CreateScratchName("sFactor",data_type="RasterDataset",workspace=scratchWS)
    outsFactor = (Power(slope,2) * 0.006541) + ((Raster(slope) * 0.0456) + 0.065)
    outsFactor.save(sFactor)

    # Calculate L Factor
    # Con("%slope%" < 1,Power("%FlowLenft%" / 72.5,0.2) ,Con(("%slope%" >=  1) &("%slope%" < 3) ,Power("%FlowLenft%" / 72.5,0.3), Con(("%slope%" >= 3) &("%slope%" < 5 ),Power("%FlowLenft%" / 72.5,0.4) ,Power("%FlowLenft%" / 72.5,0.5))))
    # 1) slope < 1      --  Power 0.2
    # 2) 1 < slope < 3  --  Power 0.3
    # 3) 3 < slope < 5  --  Power 0.4
    # 4) slope > 5      --  Power 0.5
    AddMsgAndPrint("\nCalculating L Factor")
    lFactor = arcpy.CreateScratchName("lFactor",data_type="RasterDataset",workspace=scratchWS)
    outlFactor = Con(slope < 1,Power(Raster(flowLengthFT) / 72.5,0.2),Con((slope >=  1) & (slope < 3), Power(Raster(flowLengthFT) / 72.5,0.3), Con((slope >= 3) & (slope < 5 ), Power(Raster(flowLengthFT) / 72.5,0.4), Power(Raster(flowLengthFT) / 72.5,0.5))))
    outlFactor.save(lFactor)

    # Calculate LS Factor
    # "%l_factor%" * "%s_factor%"
    AddMsgAndPrint("Calculating LS Factor")
    lsFactor = arcpy.CreateScratchName("lsFactor",data_type="RasterDataset",workspace=scratchWS)
    outlsFactor = lFactor * sFactor
    outlsFactor.save(lsFactor)

    """---------------------------------------------------------------------------------------------- Convert K,T & R Factor to Rasters """
    kFactor = arcpy.CreateScratchName("kFactor",data_type="RasterDataset",workspace=scratchWS)
    tFactor = arcpy.CreateScratchName("tFactor",data_type="RasterDataset",workspace=scratchWS)
    rFactor = arcpy.CreateScratchName("rFactor",data_type="RasterDataset",workspace=scratchWS)

    AddMsgAndPrint("\nConverting K Factor field to a raster")
    arcpy.FeatureToRaster_conversion(aoiCluIntersect,kFactorFld,kFactor,cellSize)

    AddMsgAndPrint("\nConverting T Factor field to a raster")
    arcpy.FeatureToRaster_conversion(aoiCluIntersect,tFactorFld,tFactor,cellSize)

    AddMsgAndPrint("\nConverting R Factor field to a raster")
    arcpy.FeatureToRaster_conversion(aoiCluIntersect,rFactorFld,rFactor,cellSize)

    """---------------------------------------------------------------------------------------------- Calculate EI Factor"""
    AddMsgAndPrint("\nCalculating EI Factor")
    eiFactor = arcpy.CreateScratchName("eiFactor",data_type="RasterDataset",workspace=scratchWS)
    outEIfactor = Divide((lsFactor * kFactor * rFactor),tFactor)

