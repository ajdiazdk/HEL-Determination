# ==========================================================================================
# Name:   HEL Determination by AOI
#
# Author: Adolfo.Diaz
#         Region 10 GIS Specialist
# e-mail: adolfo.diaz@usda.gov
# phone: 608.662.4422 ext. 216

# Contributor: Kevin Godsey
#              Soil Scientist
# e-mail: kevin.godsey@usda.gov
# phone: 608.662.4422 ext. 190

# Contributor: Christiane Roy
#              MN Area GIS Specialist
# e-mail: christian.roy@usda.gov
# phone: 507.405.3580

# ==========================================================================================
# Modified 10/24/2016
# Line 705; outflowLengthFT was added to the scratchLayers to be deleted.  It should've
# been flowLengthFT
#
# The acreConversion variable was updated to reference a dictionary instead.  The original
# acreConversion variable was being determined from the DEM.  If input DEM was in FT then
# the wrong acre conversion was applied to layers that were in Meters.

# ==========================================================================================
# Modified 11/19/2018
# Tim Prescott was having acre disrepancies.  He had a wide variety of coordinate systems
# in his arcmap project.  I couldn't isolate the problem but I wound up changine
# the way the z-factor is assigned.  I was assuming a meter and/or feet combination but
# there could be other combinations like centimeters. Created a matrix of XY and Z
# unit combinations to be used as a look up table.

# ==========================================================================================
# Modified 11/19/2018
# Tim Prescott was having acre disrepancies.  He had a wide variety of coordinate systems
# in his arcmap project.  I couldn't isolate the problem but I wound up changine
# the way the z-factor is assigned.  I was assuming a meter and/or feet combination but
# there could be other combinations like centimeters. Created a matrix of XY and Z
# unit combinations to be used as a look up table.

# ==========================================================================================
# Modified 11/15/2018
# Christiane reported duplicate labeling in the HEL Summary feature class.  The duplicate
# labels go away when the corresponding .lyr file is added instead of the feature class.
# Modified the code to add the .lyr to Arcmap only for the HEL Summary layer.

# ==========================================================================================
# Modified 3/1/2019
# - Error Encountered:
#      Related to optional parameter of zUnits
#	   File "C:\python_scripts\GitHub\HEL\HEL_Determination_by_AOI.py", line 517, in <module>
#      zFactor = zFactorList[unitLookUpDict.get(zUnits)][unitLookUpDict.get(linearUnits)]
#	   TypeError: list indices must be integers, not NoneType
#
# - Add Focal Statistics (MAXIMUM) to Flow Length output to smooth it out
# - switched method to exit python interpreter from exit() to sys.exit() and SystemExit Exception
#   to catch errors.
# - Switched ALL output layers to in-memory layers to optimize performance; Performance increased
#   by at least 25%
# - Updated the Determination Date to include the date the script was executed: mm/dd/yyyy
# - In the 'Compute Summary of NEW HEL values' section the total CLU Acres were re-computed from the
#   outTabulate table vs. looking them up from the 'helYesNo' fc b/c they wouldn't add up to 100%
#   The outTabulate represents sums from raster cells vs. polygons.
# - Wrapped main body into 'main' function and executed from if __name__ == '__main__':
# - sys.exit() was throwing an exception and forcing a traceback to be printed.  Ignore traceback
#   message if thrown by sys.exit()

# ==========================================================================================
# Modified 3/14/2019
# - The tool was updated so that it no longer exits when there is no PHEL present or final results
#   are all HEL or NHEL.  The 'outTabulate' table had to be updated to add either 'VALUE_1' or
#   'VALUE_2' field and calc'ed to 0 so that the script could finish.
#
# - Created 'ogCLUinfoDict' to capture cluNumber and cluAcres during the original summary process.
#   This dictionary is only used if final results are ALL HEL or ALL NHEL since acres don't need
#   to be recalculated, especially from the tabulate areas results where the acreages may slightly
#   differ b/c of pixel tabulation.  This way the acreages will always match the original figures.
#
# - 2 booleans (bOnlyHEL & bOnlyNHEL) were added to determine if final results are ALL HEL or NHEL

# ==========================================================================================
# Modified 2/4/2019
# - Seperated the scratch layers by individual task vs adding several layers at a time.
#   This will ensure all scratch layers are deleted when a failure occurs.
# - Run Focal Statistics (MEAN) to Slope output to smooth it out
# - Changed Focal Statistics on Flow Length from MAXIMUM to MEAN

# ==========================================================================================
# Modified 4/12/2019
# - Report summary of CLU HEL values regardless of no PHEL Values found.  If no PHEL values
#   are found then geoprocessing will be entirely skipped however the 1026 form will still
#   be populated and brought up.  The populateForm function was added to be able to skip
#   geoprocessing and still open the 1026 form.  If no PHEL values are found then HEL will
#   be determined by the 33.3% or 50 acre rule.  If a CLU's acreage is greater or equal to 50
#   or a CLU's acre percentage is greater or equal to 33.3% then it is "HEL" otherwise the
#   rating is determined by its dominant acre HEL value.
# - Set the 33.3% or 50 acre HEL rule when computing new HEL values after geoprocessing.
# - Autopopulated 11th parameter of tool, DC Signature, to computer user name.  Assumption
#   is made that the user is the DC. User could be the technician.
# - Autopopulate 10th Parameter, State, by isolating state namne from user computer name.

# =================
# Questions:
# 1)

#-------------------------------------------------------------------------------

## ===================================================================================
def AddMsgAndPrint(msg, severity=0):
    # prints message to screen if run as a python script
    # Adds tool message to the geoprocessor
    #
    #Split the message on \n first, so that if it's multiple lines, a GPMessage will be added for each line
    try:
        #print msg

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
# Print traceback exceptions.  If sys.exit was trapped by default exception then
# ignore traceback message.

    try:
        exc_type, exc_value, exc_traceback = sys.exc_info()
        theMsg = "\t" + traceback.format_exception(exc_type, exc_value, exc_traceback)[1] + "\n\t" + traceback.format_exception(exc_type, exc_value, exc_traceback)[-1]

        if theMsg.find("sys.exit") > -1:
            AddMsgAndPrint("\n\n")
            pass
        else:
            AddMsgAndPrint(theMsg,2)

    except:
        AddMsgAndPrint("Unhandled error in errorMsg method", 2)
        pass

## ===================================================================================
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

# ===================================================================================
def FindField(layer,chkField):
    # Check table or featureclass to see if specified field exists
    # If fully qualified name is found, return that name; otherwise return ""
    # Set workspace before calling FindField

    try:

        if arcpy.Exists(layer):

            theDesc = arcpy.Describe(layer)
            theFields = theDesc.fields
            theField = theFields[0]

            for theField in theFields:

                # Parses a fully qualified field name into its components (database, owner name, table name, and field name)
                parseList = arcpy.ParseFieldName(theField.name) # (null), (null), (null), MUKEY

                # choose the last component which would be the field name
                theFieldname = parseList.split(",")[len(parseList.split(','))-1].strip()  # MUKEY

                if theFieldname.upper() == chkField.upper():
                    return theField.name

            return False

        else:
            AddMsgAndPrint("\tInput layer not found", 0)
            return False

    except:
        errorMsg()
        return False

# ===================================================================================
def removeScratchLayers():
    # This function is the last task that is executed or gets invoked in
    # an except clause.  Simply removes all temporary scratch layers.

    import itertools

    try:
        for lyr in list(itertools.chain(*scratchLayers)):
            if arcpy.Exists(lyr):
                try:
                    arcpy.Delete_management(lyr)
                except:
                    continue
    except:
        pass

# ===================================================================================
def populateForm():
    # This function will prepare the 1026 form by adding 16 fields to the
    # helYesNo feauture class.  This function will still be invoked even if
    # there was no geoprocessing executed due to no PHEL values to be computed.
    # if bNoPHELvalues is true, 3 fields will be added that otherwise would've
    # been added after HEL geoprocessing.  However, values to these fields will
    # be populated from the ogCLUinfoDict and represent original HEL values
    # primarily determined by the 33.3% or 50 acre rule

    try:

        AddMsgAndPrint("\nPreparing and Populating 1026 Form", 0)

        # Add 3 fields and populate them from ogCLUinfoDict
        if bNoPHELvalues:

            # Add the following fields to helYesNo
            fieldList = ["HEL_Acres","HEL_Pct","HEL_YES"]
            for field in fieldList:
                if not FindField(helYesNo,field):
                    if field == "HEL_YES":
                        arcpy.AddField_management(helYesNo,field,"TEXT","","",5)
                    else:
                        arcpy.AddField_management(helYesNo,field,"FLOAT")
            fieldList.append(cluNumberFld)

            # ['HEL_Acres', 'HEL_Pct', 'HEL_YES', u'CLUNBR']
            with arcpy.da.UpdateCursor(helYesNo,fieldList) as cursor:
                for row in cursor:
                    og_cluHELacres = ogCLUinfoDict.get(row[3])[0]
                    og_cluHELdetermination = ogCLUinfoDict.get(row[3])[1]
                    og_cluHELpct = ogCLUinfoDict.get(row[3])[2]

                    if og_cluHELdetermination == "HEL":
                        row[2] = "Yes"
                    elif og_cluHELdetermination == "NHEL":
                        row[2] = "No"
                    else:
                         row[2] = "NA"
                         AddMsgAndPrint("\n\tError in populating 'HEL_YES' field",2)

                    row[0] = og_cluHELacres
                    row[1] = og_cluHELpct
                    cursor.updateRow(row)
            del cursor

        # Collect Time information
        today = datetime.date.today()
        today = today.strftime('%b %d, %Y')

        # Add 18 Fields to the helYesNo feature class
        fieldDict = {"Signature":("TEXT",dcSignature,50),"SoilAvailable":("TEXT","Yes",5),"Completion":("TEXT","Office",10),
                        "SodbustField":("TEXT","No",5),"Delivery":("TEXT","Mail",10),"Remarks":("TEXT",
                        "This preliminary determination was conducted off-site with LiDAR data only if PHEL mapunits are present.",110),
                        "RequestDate":("DATE",""),"LastName":("TEXT","",50),"FirstName":("TEXT","",25),"Address":("TEXT","",50),
                        "City":("TEXT","",25),"ZipCode":("TEXT","",10),"Request_from":("TEXT","Landowner",15),"HELFarm":("TEXT","Yes",5),
                        "Determination_Date":("DATE",today),"state":("TEXT",state,2),"SodbustTract":("TEXT","No",5),"Lidar":("TEXT","Yes",5)}

        arcpy.SetProgressor("step", "Preparing and Populating 026 Form", 0, len(fieldDict), 1)

        for field,params in fieldDict.iteritems():
            arcpy.SetProgressorLabel("Adding Field: " + field + r' to "HEL YES NO" layer')
            try:
                fldLength = params[2]
            except:
                fldLength = 0
                pass

            arcpy.AddField_management(helYesNo,field,params[0],"#","#",fldLength)

            if len(params[1]) > 0:
                expression = "\"" + params[1] + "\""
                arcpy.CalculateField_management(helYesNo,field,expression,"VB")
            arcpy.SetProgressorPosition()

        AddMsgAndPrint("\tOpening 026 Form",0)
        subprocess.Popen([msAccessPath,helDatabase])

        return True

    except:
        return False
        errorMsg()

# ===================================================================================
def AddLayersToArcMap():
    # This Function will add necessary layers to ArcMap and prepare their
    # symbology.  The 1026 form will also be prepared.

    try:
        #AddMsgAndPrint("\n")  # Strictly Formatting

        # List of layers to add to Arcmap (layer path, arcmap layer name)
        if bNoPHELvalues:
            addToArcMap = [(helSummary,"HEL Summary Layer"),(helYesNo,"HEL YES NO")]
        else:
            addToArcMap = [(finalHELmap,"Final HEL Map"),(helSummary,"HEL Summary Layer"),(helYesNo,"HEL YES NO")]

        # Put this section in a try-except. It will fail if run from ArcCatalog
        mxd = arcpy.mapping.MapDocument("CURRENT")
        df = arcpy.mapping.ListDataFrames(mxd)[0]

        # redundant workaround.  ListLayers returns a list of layer objects
        # had to create a list of layer name Strings in order to see if a
        # specific layer currently exists in Arcmap.
        currentLayersObj = arcpy.mapping.ListLayers(mxd)
        currentLayersStr = [str(x) for x in arcpy.mapping.ListLayers(mxd)]

        for layer in addToArcMap:

            # remove layer from ArcMap if it exists
            if layer[1] in currentLayersStr:
                arcpy.mapping.RemoveLayer(df,currentLayersObj[currentLayersStr.index(layer[1])])

            if layer[1] == "Final HEL Map":
                tempLayer = arcpy.MakeRasterLayer_management(layer[0],layer[1])
            else:
                tempLayer = arcpy.MakeFeatureLayer_management(layer[0],layer[1])

            result = tempLayer.getOutput(0)
            symbology = os.path.join(os.path.dirname(sys.argv[0]),layer[1].lower().replace(" ","") + ".lyr")

            arcpy.ApplySymbologyFromLayer_management(result,symbology)

            # The HEL Summary Layer that was being added to ArcMap had duplicate
            # labels in spite of being multi-part.  By adding the .lyr file
            # to ArcMap instead of the feature class (don't understand why)
            if layer[1] == "HEL Summary Layer":
                 helSummaryLYR = arcpy.mapping.Layer(symbology)
                 arcpy.mapping.AddLayer(df, helSummaryLYR, "TOP")

                 # turn layer off
                 for lyr in arcpy.mapping.ListLayers(mxd, layer[1]):
                     lyr.visible = False

            else:
                 arcpy.mapping.AddLayer(df, result, "TOP")

            """ The following code will update the layer symbology for HEL Summary Layer"""
##                # to NOT include AOI acres and percentage.
##                if layer[1] == "HEL Summary Layer":
##                    lyr = arcpy.mapping.ListLayers(mxd, layer[1])[0]
##                    # lyr.symbology.classLabels = ogHELsymbologyLabels
##                    # lyr.visible = False
##                    # arcpy.RefreshActiveView()
##                    # arcpy.RefreshTOC()
##                    # del ly
##                    expression = """[HEL] & vbNewLine &  round([HEL_Acres] ,1) & "ac." & " (" & round([HEL_AcrePct] ,1) & "%)" & vbNewLine """
##                    if lyr.supports("LABELCLASSES"):
##                        for lblClass in lyr.labelClasses:
##                            if lblClass.showClassLabels:
##                                lblClass.expression = expression
##                        lyr.showLabels = True
##                        arcpy.RefreshActiveView()
##                        arcpy.RefreshTOC()
##                    del lyr

            """ The following code will update the layer symbology for the Final HEL Map. """
            # cannot add acres and percent "by AOI". acres and pecent can only be calculated by CLU field.
            if layer[1] == "Final HEL Map":
                lyr = arcpy.mapping.ListLayers(mxd, layer[1])[0]
                newHELsymbologyLabels = []
                acreConversion = acreConversionDict.get(arcpy.Describe(helYesNo).SpatialReference.LinearUnitName)

                # This assumes that both "VALUE_1 and VALUE_2 are present
                HEL = sum([rows[0] for rows in arcpy.da.SearchCursor(outTabulate, ("VALUE_2"))])/acreConversion
                NHEL = sum([rows[0] for rows in arcpy.da.SearchCursor(outTabulate, ("VALUE_1"))])/acreConversion

                if HEL > 0:
                   newHELsymbologyLabels.append("HEL")

                if NHEL > 0:
                   newHELsymbologyLabels.append("NHEL")

                lyr.symbology.classBreakLabels = newHELsymbologyLabels
                arcpy.RefreshActiveView()
                arcpy.RefreshTOC()
                del lyr,newHELsymbologyLabels

            if layer[1] == "HEL YES NO":
                lyr = arcpy.mapping.ListLayers(mxd, layer[1])[0]
                #expression = """def FindLabel ( [CLUNBR], [HEL_Acres], [HEL_AcrePct], [HEL_YES] ):  return "CLU #: " + [CLUNBR] + "\nHEL Acres: " + str(round(float([HEL_Acres]),1)) + " (" + str(round(float( [HEL_Pct] ),1)) + "%)\nHEL: " + [HEL_YES]"""
                expression = """"Tract: " & [TRACTNBR] & vbNewLine & "CLU #: " & [CLUNBR] & vbNewLine & "HEL: " & round([HEL_Acres] ,1) & " ac." & " (" & round([HEL_Pct] ,1) & "%)" & vbNewLine & "HEL: " & [HEL_YES]"""

                if lyr.supports("LABELCLASSES"):
                    for lblClass in lyr.labelClasses:
                        if lblClass.showClassLabels:
                            lblClass.expression = expression
                    lyr.showLabels = True
                    arcpy.RefreshActiveView()
                    arcpy.RefreshTOC()

            AddMsgAndPrint("Added " + layer[1] + " to your ArcMap Session",0)

    except:
        errorMsg()
        pass



## =============================================== Main Body ====================================================

import sys, string, os, locale, traceback, urllib, re, arcpy, datetime
import subprocess, operator, getpass, time
from arcpy import env
from arcpy.sa import *

if __name__ == '__main__':

    try:

        AOI = arcpy.GetParameter(0)
        cluLayer = arcpy.GetParameter(1)
        helLayer = arcpy.GetParameter(2)
        kFactorFld = arcpy.GetParameterAsText(3)
        tFactorFld = arcpy.GetParameterAsText(4)
        rFactorFld = arcpy.GetParameterAsText(5)
        helFld = arcpy.GetParameterAsText(6)
        inputDEM = arcpy.GetParameter(7)
        zUnits = arcpy.GetParameterAsText(8)
        state = arcpy.GetParameterAsText(9)
        dcSignature = arcpy.GetParameterAsText(10)

        ##    AOI = r'C:\python_scripts\HEL_MN\Sample\CLU_subset2.shp'
        ##    cluLayer = r'C:\python_scripts\HEL_MN\Sample\clu_sample.shp'
        ##    helLayer = r'C:\python_scripts\HEL_MN\Sample\HEL_sample.shp'
        ##    kFactorFld = "K"
        ##    tFactorFld = "T"
        ##    rFactorFld = "R"
        ##    helFld = "HEL"
        ##    inputDEM = r'C:\python_scripts\HEL_MN\Sample\dem_03'
        ##    zUnits = ""

        """ ---------------------------------------------------------------------------------------------- Check HEL Access Database"""
        # ------------------------------------------------------------ Make sure the HEL access database is present, Exit otherwise
        helDatabase = os.path.dirname(sys.argv[0]) + os.sep + r'HEL.mdb'
        if not arcpy.Exists(helDatabase):
            AddMsgAndPrint("\nHEL Access Database does not exist in the same path as HEL Tools",2)
            sys.exit()

        # ----------------------------------------- Set the path to the final HEL_YES_NO layer.  Essentially derived from user AOI
        helYesNo = os.path.join(helDatabase, r'HEL_YES_NO')
        helSummary = os.path.join(helDatabase, r'HELSummaryLayer')
        finalHELmap = os.path.join(helDatabase, r'finalhelmap')

        if arcpy.Exists(helYesNo):
            try:
                arcpy.Delete_management(helYesNo)
            except:
                AddMsgAndPrint("\nCould not delete the 'HEL_YES_NO' feature class in the HEL access database. Creating an additional layer",2)
                helYesNo = arcpy.CreateScratchName("HEL_YES_NO",data_type="FeatureClass",workspace=helDatabase)

        if arcpy.Exists(helSummary):
            try:
                arcpy.Delete_management(helSummary)
            except:
                AddMsgAndPrint("\nCould not delete the 'HELSummaryLayer' feature class in the HEL access database. Creating an additional layer",2)
                helSummary = arcpy.CreateScratchName("HELSummaryLayer",data_type="FeatureClass",workspace=helDatabase)

        if arcpy.Exists(finalHELmap):
            try:
                arcpy.Delete_management(finalHELmap)
            except:
                AddMsgAndPrint("\nCould not delete the 'Final_HEL_Map' raster layer in the HEL access database. Creating an additional layer",2)
                finalHELmap = arcpy.CreateScratchName("finalhelmap", data_type="RasterDataset", workspace=helDatabase)

        # --------------------------------------------------------- determine Microsoft Access path from windows version
        bAccess = True
        winVersion = sys.getwindowsversion()

        # Windows 10
        if winVersion.build == 9200:
            msAccessPath = r'C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE'
        # Windows 7
        elif winVersion.build == 7601:
            msAccessPath = r'C:\Program Files (x86)\Microsoft Office\Office15\MSACCESS.EXE'
        else:
            AddMsgAndPrint("\nCould not determine Windows version, will not populate 026 Form",2)
            bAccess = False

        if bAccess and not os.path.isfile(msAccessPath):
            bAccess = False
            AddMsgAndPrint("\nCould not locate Microsoft Access Software, will not populate 026 Form",2)

        arcpy.env.overwriteOutput = True

        """ ------------------------------------------------------------------------------------------------------------------------ Routine Stuff"""
        # Check Availability of Spatial Analyst Extension
        try:
            if arcpy.CheckExtension("Spatial") == "Available":
                arcpy.CheckOutExtension("Spatial")
            else:
                raise LicenseError

        except LicenseError:
            AddMsgAndPrint("\n\nSpatial Analyst license is unavailable.  Go to Customize -> Extensions to activate it",2)
            AddMsgAndPrint("\n\nExiting!")
            sys.exit()
        except arcpy.ExecuteError:
            AddMsgAndPrint(arcpy.GetMessages(2),2)
            sys.exit()

        # Set overwrite option
        arcpy.env.overwriteOutput = True

        # define and set the scratch workspace
        scratchWS = setScratchWorkspace()
        arcpy.env.scratchWorkspace = scratchWS
        scratchLayers = list()

        if not scratchWS:
            AddMsgAndPrint("\tExiting!")
            sys.exit()

        """ ------------------------------------------------------------------------------ Determine where AOI is coming from - Stamp CLU in helYesNO.
                                                                                           layer represents the user AOI and the CLUs together"""
        descAOI = arcpy.Describe(AOI)
        aoiPath = descAOI.catalogPath
        cluLayerPath = arcpy.Describe(cluLayer).catalogPath

        # AOI is digitized and resides in memory; Clip CLU based on the user-digitized polygon.
        if descAOI.dataType.upper() == "FEATURERECORDSETLAYER":
            AddMsgAndPrint("\nClipping the CLU layer to the manually digitized AOI")
            arcpy.Clip_analysis(cluLayer,AOI,helYesNo)

        # AOI is some existing feature.  Not manually digitized.
        else:
            # AOI came from the CLU - Most Popular option
            if aoiPath == cluLayerPath:
                AddMsgAndPrint("\nUsing " + str(int(arcpy.GetCount_management(AOI).getOutput(0))) + " features from the CLU Layer as AOI")
                helYesNo = arcpy.CopyFeatures_management(AOI,helYesNo)

            # AOI is not the CLU but an existing layer
            else:
                AddMsgAndPrint("\nClipping the CLU layer to " + descAOI.name + " (AOI)")
                arcpy.Clip_analysis(cluLayer,AOI,helYesNo)

        # Update Acres - Add Calcacre field if it doesn't exist.
        calcAcreFld = "CALCACRES"
        if not len(arcpy.ListFields(helYesNo,calcAcreFld)) > 0:
            arcpy.AddField_management(helYesNo,calcAcreFld,"DOUBLE")

        arcpy.CalculateField_management(helYesNo,calcAcreFld,"!shape.area@acres!","PYTHON_9.3")
        totalAcres = float("%.1f" % (sum([row[0] for row in arcpy.da.SearchCursor(helYesNo, (calcAcreFld))])))
        AddMsgAndPrint("\tTotal Acres: " + splitThousands(totalAcres))
        del totalAcres

        """ ---------------------------------------------------------------------------------------------- Check DEM Coordinate System, Linear Units, Z-factor"""
        # lookup dictionary to convert XY units to area.  Key = XY unit of DEM; Value = conversion factor to sq.meters
        acreConversionDict = {'Meter':4046.85642,'Foot':43560,'Foot_US':43560,'Centimeter':40470000,'Inch':6273000}

        desc = arcpy.Describe(inputDEM)
        inputDEMPath = desc.catalogPath
        sr = desc.SpatialReference
        linearUnits = sr.LinearUnitName
        cellSize = desc.MeanCellWidth

        # Coordinate System must be a Projected type
        if not sr.Type == "Projected":
            AddMsgAndPrint("\n\n\t" + os.path.basename(inputDEMPath) + " is NOT in a projected Coordinate System....EXITING",2)
            sys.exit()
        else:
            AddMsgAndPrint("\nGathering information about DEM Layer: " + os.path.basename(inputDEMPath))

        # xy Linear units must be defined
        if not linearUnits:
            AddMsgAndPrint("\tCould not determine linear units of DEM....Exiting!",2)
            sys.exit()

        # Assign Z-factor based on XY and Z units of DEM
        # the following represents a matrix of possible z-Factors
        # using different combination of xy and z units
        # ----------------------------------------------------
        #                      Z - Units
        #                       Meter    Foot     Centimeter     Inch
        #          Meter         1	    0.3048	    0.01	    0.0254
        #  XY      Foot        3.28084	  1	      0.0328084	    0.083333
        # Units    Centimeter   100	    30.48	     1	         2.54
        #          Inch        39.3701	  12       0.393701	      1
        # ---------------------------------------------------

        unitLookUpDict = {'Meter':0,'Meters':0,'Foot':1,'Foot_US':1,'Feet':1,'Centimeter':2,'Centimeters':2,'Inch':3,'Inches':3}
        zFactorList = [[1,0.3048,0.01,0.0254],
                       [3.28084,1,0.0328084,0.083333],
                       [100,30.48,1,2.54],
                       [39.3701,12,0.393701,1]]

        # if zUnits not populated assume it is the same as linearUnits
        bZunits = True
        if not zUnits: zUnits = linearUnits; bZunits = False

        # look up zFactor based on units (z,xy -- I should've reversed the table)
        zFactor = zFactorList[unitLookUpDict.get(zUnits)][unitLookUpDict.get(linearUnits)]

        AddMsgAndPrint("\tProjection Name: " + sr.Name,0)
        AddMsgAndPrint("\tXY Linear Units: " + linearUnits,0)
        AddMsgAndPrint("\tElevation Values (Z): " + zUnits,0)
        if not bZunits: AddMsgAndPrint("\t\tZ-units were auto set to: " + linearUnits,0)
        AddMsgAndPrint("\tCell Size: " + str(desc.MeanCellWidth) + " x " + str(desc.MeanCellHeight) + " " + linearUnits,0)
        AddMsgAndPrint("\tZ-Factor: " + str(zFactor))

        arcpy.SetProgressor("step", "Calculating HEL Determination", 0, 17, 1)

        # Snap every raster layer to the DEM
        arcpy.env.snapRaster = inputDEMPath

        """ ------------------------------------------------------------------------------------------------------------- Compute Summary of original HEL values"""

        # -------------------------------------------------------------------------- Intersect helYesNo (CLU & AOI) with soils (helLayer) -> cluHELintersect
        AddMsgAndPrint("\nComputing summary of original HEL Values")
        arcpy.SetProgressorLabel("Computing summary of original HEL Values")
        cluHELintersect = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("aoiCLUIntersect",data_type="FeatureClass",workspace=scratchWS))
        #cluHELintersect = arcpy.CreateScratchName("aoiCLUIntersect",data_type="FeatureClass",workspace=scratchWS)
        arcpy.Intersect_analysis([helYesNo,helLayer],cluHELintersect,"ALL")
        scratchLayers.append(cluHELintersect)

        # Test intersection --- Should we check the percentage of intersection here? what if only 50% overlap
        totalIntAcres = sum([row[0] for row in arcpy.da.SearchCursor(cluHELintersect, ("SHAPE@AREA"))]) / acreConversionDict.get(arcpy.Describe(cluHELintersect).SpatialReference.LinearUnitName)
        if not totalIntAcres:
            AddMsgAndPrint("\tThere is no overlap between AOI and CLU Layer. EXITTING!",2)
            removeScratchLayers()
            sys.exit()

        # -------------------------------------------------------------------------- Dissolve intersection output by the following fields -> helSummary
        # check for critical fields in CLU layer; exit if missing
        cluNumberFld = "CLUNBR"
        dissovleFlds = [cluNumberFld,"TRACTNBR","FARMNBR","COUNTYCD","CALCACRES"]
        for fld in dissovleFlds:
            if not FindField(helYesNo,fld):
                AddMsgAndPrint("\n\tMissing CLU Layer field: " + fld + " ---- Exiting!",2)
                removeScratchLayers()
                sys.exit()

        dissovleFlds.append(helFld)
        arcpy.Dissolve_management(cluHELintersect, helSummary, dissovleFlds, "","MULTI_PART", "DISSOLVE_LINES")

        # -------------------------------------------------------------------------- Make sure TRACTNBR and FARMNBR  are uniqe; exit otherwise
        uniqueTracts = set([row[0] for row in arcpy.da.SearchCursor(helSummary,("TRACTNBR"))])
        uniqueFarm   = set([row[0] for row in arcpy.da.SearchCursor(helSummary,("FARMNBR"))])

        if len(uniqueTracts) != 1:
           AddMsgAndPrint("\n\tThere are " + str(len(uniqueTracts)) + " different Tract Numbers. Exiting!",2)
           for tract in uniqueTracts:
               AddMsgAndPrint("\t\tTract #: " + str(tract),2)
           removeScratchLayers
           sys.exit()

        if len(uniqueFarm) != 1:
           AddMsgAndPrint("\n\tThere are " + str(len(uniqueFarm)) + " different Farm Numbers. Exiting!",2)
           for farm in uniqueFarm:
               AddMsgAndPrint("\t\tFarm #: " + str(farm),2)
           removeScratchLayers
           sys.exit()
        del uniqueTracts,uniqueFarm

        # --------------------------------------------------------------------------- Add and Update fields in the HEL Summary Layer (HEL Value, HEL Acres, HEL_Pct)
        HELvalueFld = 'HELValue'    # Used for raster purposes
        HELacres = 'HEL_Acres'
        HELacrePct = 'HEL_AcrePct'

        if not len(arcpy.ListFields(helSummary,HELvalueFld)) > 0:
            arcpy.AddField_management(helSummary,HELvalueFld,"SHORT")

        if not len(arcpy.ListFields(helSummary,HELacres)) > 0:
            arcpy.AddField_management(helSummary,HELacres,"DOUBLE")

        if not len(arcpy.ListFields(helSummary,HELacrePct)) > 0:
            arcpy.AddField_management(helSummary,HELacrePct,"DOUBLE")

        # Calculate HELValue Field
        helDict = dict()          ## tallies acres by HEL value i.e. PHEL:100
        nullHEL = 0               ## # of polygons with no HEL values
        wrongHELvalues = list()   ## Stores incorrect HEL Values
        maxAcreLength = list()    ## Stores the acres for formatting purposes
        bNoPHELvalues = False     ## Boolean flag for indicate PHEL values are missing

        with arcpy.da.UpdateCursor(helSummary,[helFld,HELvalueFld,HELacres,"SHAPE@AREA",HELacrePct,calcAcreFld]) as cursor:
            for row in cursor:

                # Acre information
                acres = float("%.1f" % (row[3] / acreConversionDict.get(arcpy.Describe(helSummary).SpatialReference.LinearUnitName)))
                row[2] = acres
                maxAcreLength.append(acres)
                row[4] = round((row[2] / row[5]) * 100,1) # HEL acre percentage

                # Collect NULL polygons
                if row[0] is None or row[0] == '':
                    nullHEL+=1
                    continue

                # Used for rasterization purposes
                elif row[0] == "PHEL":
                    row[1] = 0
                elif row[0] == "HEL":
                    row[1] = 1
                elif row[0] == "NHEL":
                    row[1] = 2
                else:
                    if not str(row[0]) in wrongHELvalues:
                        wrongHELvalues.append(str(row[0]))

                # Add hel value to dictionary to summarize by total project
                if not helDict.has_key(row[0]):
                    helDict[row[0]] = acres
                else:
                    helDict[row[0]] += acres

                cursor.updateRow(row)
                del acres

        # No PHEL values were found; Bypass geoprocessing and populate form
        if not helDict.has_key('PHEL'):
            AddMsgAndPrint("\n\tWARNING: There are no PHEL values in HEL layer",1)
            bNoPHELvalues = True

        # Inform user about NULL values
        if nullHEL:
            AddMsgAndPrint("\n\tWARNING: There are " + str(nullHEL) + " polygon(s) with no HEL values",1)

        # Inform user about HEL values that are not PHEL,HEL, NHEL
        if wrongHELvalues:
            AddMsgAndPrint("\n\tWARNING: There is " + str(len(set(wrongHELvalues))) + " incorrect HEL values in HEL Layer:",1)
            for wrongVal in set(wrongHELvalues):
                AddMsgAndPrint("\t\t" + wrongVal)

        del dissovleFlds,nullHEL,wrongHELvalues

        ## --------------------------------------------------------------------------------------------------------- Report HEl Layer Summary by CLU
        # Create 2 temporary tables to capture summary statistics
        ogHelSummaryStats = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("ogHELSummaryStats",data_type="ArcInfoTable",workspace=scratchWS))
        ogHelSummaryStatsPivot = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("ogHELSummaryStatsPivot",data_type="ArcInfoTable",workspace=scratchWS))
        #ogHelSummaryStats = arcpy.CreateScratchName("ogHELSummaryStats",data_type="ArcInfoTable",workspace=scratchWS)
        #ogHelSummaryStatsPivot = arcpy.CreateScratchName("ogHELSummaryStatsPivot",data_type="ArcInfoTable",workspace=scratchWS)

        stats = [[HELacres,"SUM"]]
        caseField = [cluNumberFld,helFld]
        arcpy.Statistics_analysis(helSummary, ogHelSummaryStats, stats, caseField)
        sumHELacreFld = [fld.name for fld in arcpy.ListFields(ogHelSummaryStats,"*" + HELacres)][0]
        scratchLayers.append(ogHelSummaryStats)

        # Pivot table will have CLUNBR & any HEL values present (HEL,NHEL,PHEL)
        arcpy.PivotTable_management(ogHelSummaryStats,cluNumberFld,helFld,sumHELacreFld,ogHelSummaryStatsPivot)
        scratchLayers.append(ogHelSummaryStatsPivot)

        pivotFields = [fld.name for fld in arcpy.ListFields(ogHelSummaryStatsPivot)][1:]  # ['CLUNBR','HEL','NHEL','PHEL']
        numOfhelValues = len(pivotFields)                                                 # Number of Pivot table fields; Min 2 fields
        maxAcreLength.sort(reverse=True)

        # This dictionary will only be used if FINAl results are all HEL or all NHEL to reference original
        # acres and not use tabulate area acres.  It will also be used when there are no PHEL Values.
        # {cluNumber:(cluAcres,original HEL value} -- HEL value is determined by the 33.3% or 50 acre rule
        ogCLUinfoDict = dict()

        AddMsgAndPrint("\n\tSummary by CLU:")

        # Iterate through the pivot table and report HEL values by CLU
        with arcpy.da.SearchCursor(ogHelSummaryStatsPivot,pivotFields) as cursor:
            for row in cursor:

                og_cluHELrating = None         # original CLU HEL Rating
                og_cluHELacresList = list()    # temp list of acres by HEL value
                og_cluHELpctList = list()      # temp list of pct by HEL value
                msgList = list()               # temp list of messages to print
                cluAcres = sum([row[i] for i in range(1,numOfhelValues,1)])

                # iterate through the pivot table fields by record
                for i in range(1,numOfhelValues,1):
                    acres =  float("%.1f" % (row[i]))
                    pct = float("%.1f" % ((row[i] / cluAcres) * 100))

                    # Determine original HEL rating of CLU and populate acres and pct
                    # into ogCLUinfoDict.  Primarily for bNoPHELvalues.
                    if og_cluHELrating == None:

                        if pivotFields[i] == "HEL" and (pct >= 33.3 or acres >= 50):
                            og_cluHELrating = "HEL"
                            if not row[0] in ogCLUinfoDict:
                               ogCLUinfoDict[row[0]] = (cluAcres,og_cluHELrating,pct)

                        # This is the last field in the pivot table
                        elif i == (numOfhelValues - 1):
                            og_cluHELrating = pivotFields[i]
                            if not row[0] in ogCLUinfoDict:
                               ogCLUinfoDict[row[0]] = (cluAcres,og_cluHELrating,pct)

                        # First field did not meet HEL criteria; add it to a temp list
                        else:
                            og_cluHELacresList.append(row[i])
                            og_cluHELpctList.append(pct)

                    # Formulate messages
                    firstSpace = " " * (4-len(pivotFields[i])) # PHEL has 4 characters
                    secondSpace = " " * (len(str(round(maxAcreLength[0],1))) - len(str(acres)))
                    msgList.append(str("\t\t\t" + pivotFields[i] + firstSpace + " -- " + str(acres) + secondSpace + " .ac -- " + str(pct) + " %"))
                    #AddMsgAndPrint("\t\t\t" + pivotFields[i] + firstSpace + " -- " + str(acres) + secondSpace + " .ac -- " + str(pct) + " %")
                    del acres,pct,firstSpace,secondSpace

                # If og_cluHELrating was not populated above then assign the HEL value of the maximum acre.
                # og_cluHELrating should be populated by this point.
                if og_cluHELrating == "":
                    og_cluHELdetermination = str(pivotFields[og_cluHELacresList.index(max(og_cluHELacresList)) + 1])
                    pct = og_cluHELpctList[og_cluHELacresList.index(max(og_cluHELacres))]
                    ogCLUinfoDict[row[0]] = (cluAcres,og_cluHELrating,pct)

                # Report messages to user; og CLU HEL rating will be reported if bNoPHELvalues is true.
                if bNoPHELvalues:
                    AddMsgAndPrint("\n\t\tCLU #: " + str(row[0]) + " - Rating: " + og_cluHELrating)
                else:
                     AddMsgAndPrint("\n\t\tCLU #: " + str(row[0]))
                for msg in msgList:
                    AddMsgAndPrint(msg)

                del og_cluHELrating,og_cluHELacresList,og_cluHELpctList,msgList,cluAcres

        del stats,caseField,sumHELacreFld,pivotFields,numOfhelValues,maxAcreLength

        ## ---------------------------------------------------------------------------------------------------------  Report HEl Layer Summary by Project AOI
        ogHELsymbologyLabels = []
        validHELsymbologyValues = ['HEL','NHEL','PHEL']
        maxAcreLength = len(str(sorted([acres for val,acres in helDict.iteritems()],reverse=True)[0]))
        AddMsgAndPrint("\n\tSummary by AOI:")

        for val in validHELsymbologyValues:
            if val in helDict:
                acres = helDict[val]
                pct = round((acres/totalIntAcres)*100,1)
                firstSpace = " " * (4-len(val)) # PHEL has 4 characters
                secondSpace = " " * (maxAcreLength - len(str(acres)))
                AddMsgAndPrint("\t\t" + val + firstSpace + " -- " + str(acres) + " .ac -- " + str(pct) + " %")
                ogHELsymbologyLabels.append(val + " -- " + str(acres) +  secondSpace + " .ac -- " + str(pct) + " %")
                del acres,pct,firstSpace,secondSpace

        del helDict,validHELsymbologyValues,maxAcreLength

        # If there are no PHEL Values add helSummary and helYesNo layers to ArcMap
        # and prepare 1026 form.  Skip geoProcessing.
        if bNoPHELvalues:

            AddLayersToArcMap()

            if bAccess:
                populateForm()
            sys.exit()

        arcpy.SetProgressorPosition()

        """ ------------------------------------------------------------------------------------------------------------- Buffer CLU (AOI) Layer by 300 Meters"""
        # cluBuffer will be used as a mask to extract DEM
        arcpy.SetProgressorLabel("Buffering AOI by 300 Meters")
        cluBuffer = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("cluBuffer",data_type="FeatureClass",workspace=scratchWS))
        arcpy.Buffer_analysis(helYesNo,cluBuffer,"300 Meters","FULL","ROUND")
        scratchLayers.append(cluBuffer)
        arcpy.SetProgressorPosition()

        """ -------------------------------------------------------------------------------------------------------------Extract DEM using CLU layer"""
        ## All Raster layers produced will be in-memory only; Hopefully memory is not exceeded
        ## To reverse uncomment line above and below and rename i.e. demExtract -> outExtractMask

        arcpy.SetProgressorLabel("Extracting DEM subset using buffered AOI")
        AddMsgAndPrint("\nExtracting DEM subset using buffered AOI")
        #demExtract = arcpy.CreateScratchName("demExtract",data_type="RasterDataset",workspace=scratchWS)
        demExtract = ExtractByMask(inputDEM,cluBuffer)
        #outExtractMask.save(demExtract)
        scratchLayers.append(demExtract)
        arcpy.SetProgressorPosition()

        """-------------------------------------------------------------------------------------------------------------  Create Slope Layer"""
        arcpy.SetProgressorLabel("Creating Slope Derivative")
        AddMsgAndPrint("\tCreating Slope Derivative")
        #preslope = arcpy.CreateScratchName("preslope",data_type="RasterDataset",workspace=scratchWS)
        preslope = Slope(demExtract,"PERCENT_RISE",zFactor)
        #outSlope.save(preslope)
        scratchLayers.append(preslope)
        arcpy.SetProgressorPosition()

        # Run a FocalMean statistics on slope output
        arcpy.SetProgressorLabel("Running Focal Statistics on Slope Percent")
        AddMsgAndPrint("\tRunning Focal Statistics on Slope Percent")
        #slope = arcpy.CreateScratchName("focStatsMean_Slope",data_type="RasterDataset",workspace=scratchWS)
        slope = FocalStatistics(preslope, NbrRectangle(3,3,"CELL"),"MEAN","DATA")
        #outFocalStatistics.save(slope)

        """------------------------------------------------------------------------------------------------------------- Create Flow Direction and Flow Length"""
        arcpy.SetProgressorLabel("Calculating Flow Direction")
        AddMsgAndPrint("\tCalculating Flow Direction")
        #flowDirection = arcpy.CreateScratchName("flowDirection",data_type="RasterDataset",workspace=scratchWS)
        flowDirection = FlowDirection(demExtract, "FORCE")
        #outFlowDirection.save(flowDirection)
        scratchLayers.append(flowDirection)
        arcpy.SetProgressorPosition()

        arcpy.SetProgressorLabel("Calculating Flow Length")
        AddMsgAndPrint("\tCalculating Flow Length")
        #preflowLength = arcpy.CreateScratchName("flowLength",data_type="RasterDataset",workspace=scratchWS)
        preflowLength = FlowLength(flowDirection,"UPSTREAM", "")
        scratchLayers.append(preflowLength)
        #outpreFlowLength.save(preflowLength)

        # Run a focal statistics on flow length
        arcpy.SetProgressorLabel("Running Focal Statistics on Flow Length")
        AddMsgAndPrint("\tRunning Focal Statistics on Flow Length")
        #flowLength = arcpy.CreateScratchName("focStatsMax_FlowLength",data_type="RasterDataset",workspace=scratchWS)
        flowLength = FocalStatistics(preflowLength, NbrRectangle(3,3,"CELL"),"MEAN","DATA")
        #outFocalStatistics.save(flowLength)
        scratchLayers.append(flowLength)

        # convert Flow Length distance units to feet if original DEM is not in feet.
        if not zUnits in ('Feet','Foot','Foot_US'):
            AddMsgAndPrint("\tConverting Flow Length Distance units to Feet")
            #flowLengthFT = arcpy.CreateScratchName("flowLength_FT",data_type="RasterDataset",workspace=scratchWS)
            #outflowLengthFT = Raster(flowLength) * 3.280839896
            flowLengthFT = flowLength * 3.280839896
            #outflowLengthFT.save(flowLengthFT)
            scratchLayers.append(flowLengthFT)
            arcpy.SetProgressorPosition()

        else:
            flowLengthFT = flowLength
            scratchLayers.append(flowLengthFT)
            arcpy.SetProgressorPosition()

        """------------------------------------------------------------------------------------------------------------- Calculate LS Factor"""
        # ------------------------------------------------------------------------------- Calculate S Factor
        # ((0.065 +( 0.0456 * ("%slope%"))) +( 0.006541 * (Power("%slope%",2))))
        arcpy.SetProgressorLabel("Calculating S Factor")
        AddMsgAndPrint("\n\tCalculating S Factor")
        #sFactor = arcpy.CreateScratchName("sFactor",data_type="RasterDataset",workspace=scratchWS)
        #outsFactor = (Power(Raster(slope),2) * 0.006541) + ((Raster(slope) * 0.0456) + 0.065)       ## Original Line
        sFactor = (Power(slope,2) * 0.006541) + ((slope * 0.0456) + 0.065)
        #outsFactor.save(sFactor)
        scratchLayers.append(sFactor)
        arcpy.SetProgressorPosition()

        # ------------------------------------------------------------------------------ Calculate L Factor
        # Con("%slope%" < 1,Power("%FlowLenft%" / 72.5,0.2) ,Con(("%slope%" >=  1) &("%slope%" < 3) ,Power("%FlowLenft%" / 72.5,0.3), Con(("%slope%" >= 3) &("%slope%" < 5 ),Power("%FlowLenft%" / 72.5,0.4) ,Power("%FlowLenft%" / 72.5,0.5))))
        # 1) slope < 1      --  Power 0.2
        # 2) 1 < slope < 3  --  Power 0.3
        # 3) 3 < slope < 5  --  Power 0.4
        # 4) slope > 5      --  Power 0.5

        arcpy.SetProgressorLabel("Calculating L Factor")
        AddMsgAndPrint("\tCalculating L Factor")
        #lFactor = arcpy.CreateScratchName("lFactor",data_type="RasterDataset",workspace=scratchWS)

        # Original outlFactor lines
    ##        outlFactor = Con(Raster(slope),Power(Raster(flowLengthFT) / 72.5,0.2),
    ##                        Con(Raster(slope),Power(Raster(flowLengthFT) / 72.5,0.3),
    ##                        Con(Raster(slope),Power(Raster(flowLengthFT) / 72.5,0.4),
    ##                        Power(Raster(flowLengthFT) / 72.5,0.5),"VALUE >= 3 AND VALUE < 5"),"VALUE >= 1 AND VALUE < 3"),"VALUE<1")

        # Remove 'Raster' function from above
        lFactor = Con(slope,Power(flowLengthFT / 72.5,0.2),
                        Con(slope,Power(flowLengthFT / 72.5,0.3),
                        Con(slope,Power(flowLengthFT / 72.5,0.4),
                        Power(flowLengthFT / 72.5,0.5),"VALUE >= 3 AND VALUE < 5"),"VALUE >= 1 AND VALUE < 3"),"VALUE<1")

        #outlFactor.save(lFactor)
        scratchLayers.append(lFactor)
        arcpy.SetProgressorPosition()

        # ----------------------------------------------------------------------------- Calculate LS Factor
        # "%l_factor%" * "%s_factor%"
        arcpy.SetProgressorLabel("Calculating LS Factor")
        AddMsgAndPrint("\tCalculating LS Factor")
        #lsFactor = arcpy.CreateScratchName("lsFactor",data_type="RasterDataset",workspace=scratchWS)
        #outlsFactor = Raster(lFactor) * Raster(sFactor)  ## Original Line
        lsFactor = lFactor * sFactor
        #outlsFactor.save(lsFactor)
        scratchLayers.append(lsFactor)
        arcpy.SetProgressorPosition()

        """------------------------------------------------------------------------------------------------------------- Convert K,T & R Factor and HEL Value to Rasters """
        AddMsgAndPrint("\nConverting Vector to Raster for Spatial Analysis Purpose")

        # All raster datasets will be created in memory
        kFactor = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("kFactor",data_type="RasterDataset",workspace=scratchWS))
        tFactor = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("tFactor",data_type="RasterDataset",workspace=scratchWS))
        rFactor = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("rFactor",data_type="RasterDataset",workspace=scratchWS))
        helValue = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("helValue",data_type="RasterDataset",workspace=scratchWS))

        arcpy.SetProgressorLabel("Converting K Factor field to a raster")
        AddMsgAndPrint("\tConverting K Factor field to a raster")
        arcpy.FeatureToRaster_conversion(cluHELintersect,kFactorFld,kFactor,cellSize)
        arcpy.SetProgressorPosition()

        arcpy.SetProgressorLabel("Converting T Factor field to a raster")
        AddMsgAndPrint("\tConverting T Factor field to a raster")
        arcpy.FeatureToRaster_conversion(cluHELintersect,tFactorFld,tFactor,cellSize)
        arcpy.SetProgressorPosition()

        arcpy.SetProgressorLabel("Converting R Factor field to a raster")
        AddMsgAndPrint("\tConverting R Factor field to a raster")
        arcpy.FeatureToRaster_conversion(cluHELintersect,rFactorFld,rFactor,cellSize)
        arcpy.SetProgressorPosition()

        arcpy.SetProgressorLabel("Converting HEL Value field to a raster")
        AddMsgAndPrint("\tConverting HEL Value field to a raster")
        arcpy.FeatureToRaster_conversion(helSummary,HELvalueFld,helValue,cellSize)
        arcpy.SetProgressorPosition()
        scratchLayers.append((kFactor,tFactor,rFactor,helValue))

        """------------------------------------------------------------------------------------------------------------- Calculate EI Factor"""
        arcpy.SetProgressorLabel("Calculating EI Factor")
        AddMsgAndPrint("\nCalculating EI Factor")
        #eiFactor = arcpy.CreateScratchName("eiFactor", data_type="RasterDataset", workspace=scratchWS)
        #outEIfactor = Divide((Raster(lsFactor) * Raster(kFactor) * Raster(rFactor)),Raster(tFactor))  # Original Lines
        eiFactor = Divide((lsFactor * kFactor * rFactor),tFactor)
        #outEIfactor.save(eiFactor)
        scratchLayers.append(eiFactor)
        arcpy.SetProgressorPosition()

        """------------------------------------------------------------------------------------------------------------- Calculate Final HEL Factor"""
        # Con("%hel_factor%"==0,"%EI_grid%",Con("%hel_factor%"==1,9,Con("%hel_factor%"==2,2)))
        # Create Conditional statement to reflect the following:
        # 1) HEL Value = 0 -- Take EI factor -- Depends
        # 2) HEL Value = 1 -- Assign 9
        # 3) HEL Value = 2 -- Assign 2 (No action needed)
        # Anything above 8 is HEL

        arcpy.SetProgressorLabel("Calculating HEL Factor")
        AddMsgAndPrint("Calculating HEL Factor")
        #helFactor = arcpy.CreateScratchName("helFactor",data_type="RasterDataset",workspace=scratchWS)
        #outHELfactor = Con(Raster(helValue),Raster(eiFactor),Con(Raster(helValue),9,Raster(helValue),"VALUE=1"),"VALUE=0")  ## Original Line
        helFactor = Con(helValue,eiFactor,Con(helValue,9,helValue,"VALUE=1"),"VALUE=0")
        scratchLayers.append(helFactor)
        #outHELfactor.save(helFactor)

        #finalHEL = arcpy.CreateScratchName("finalHEL",data_type="RasterDataset",workspace=scratchWS)
        # Reclassify values:
        #       < 8 = Value_1 = NHEL
        #       > 8 = Value_2 = HEL
        remapString = "0 8 1;8 100000000 2"
        arcpy.Reclassify_3d(helFactor, "VALUE", remapString, finalHELmap,'NODATA')
        arcpy.SetProgressorPosition()

        """------------------------------------------------------------------------------------------------------------- Compute Summary of NEW HEL values"""
        arcpy.SetProgressorLabel("Computing summary of new HEL Values:")
        AddMsgAndPrint("\nComputing summary of new HEL Values:\n")

        # Summarize the total area by CLU number
        outTabulate = "in_memory" + os.sep + os.path.basename(arcpy.CreateScratchName("HEL_Tabulate",data_type="ArcInfoTable",workspace=scratchWS))
        TabulateArea(helYesNo,cluNumberFld,finalHELmap,"VALUE",outTabulate,cellSize)
        tabulateFields = [fld.name for fld in arcpy.ListFields(outTabulate)][2:]
        scratchLayers.append(outTabulate)

        # Booleans to indicate if only HEL or only NHEL is present
        bOnlyHEL = False; bOnlyNHEL = False

        if len(tabulateFields):
            if not "VALUE_1" in tabulateFields:
                AddMsgAndPrint("\tWARNING: Entire Area is HEL",1)
                arcpy.AddField_management(outTabulate,"VALUE_1","DOUBLE")
                arcpy.CalculateField_management(outTabulate,"VALUE_1",0)
                bOnlyHEL = True

            if not "VALUE_2" in tabulateFields:
                AddMsgAndPrint("\tWARNING: Entire Area is NHEL",1)
                arcpy.AddField_management(outTabulate,"VALUE_2","DOUBLE")
                arcpy.CalculateField_management(outTabulate,"VALUE_2",0)
                bOnlyNHEL = True
        else:
            AddMsgAndPrint("\n\tReclassifying helFactor Failed",2)
            sys.exit()

        fieldList = ["HEL_Acres","HEL_Pct","HEL_YES"]
        for field in fieldList:
            if not FindField(helYesNo,field):
                if field == "HEL_YES":
                    arcpy.AddField_management(helYesNo,field,"TEXT","","",5)
                else:
                    arcpy.AddField_management(helYesNo,field,"FLOAT")

        fieldList.append(cluNumberFld)
        fieldList.append(calcAcreFld)
        cluDict = dict()  # ClU: (len of clu, helAcres, helPct, len of Acres, len of pct,is it HEL?)

        # ['HEL_Acres', 'HEL_Pct', 'HEL_YES', u'CLUNBR', 'CALCACRES']
        with arcpy.da.UpdateCursor(helYesNo,fieldList) as cursor:
            for row in cursor:

                expression = arcpy.AddFieldDelimiters(outTabulate,cluNumberFld) + " = " + str(row[3])
                outTabulateValues = ([(rows[0],rows[1]) for rows in arcpy.da.SearchCursor(outTabulate, ("VALUE_1","VALUE_2"), where_clause = expression)])[0]
                acreConversion = acreConversionDict.get(arcpy.Describe(helYesNo).SpatialReference.LinearUnitName)

                # if results are completely HEL or NHEL then total clu acres from ogCLUinfoDict b/c
                # Sometimes the results will slightly vary b/c of the raster pixels.
                # Otherwise compute them from the tabulateArea results.
                if bOnlyHEL or bOnlyNHEL:
                    if bOnlyHEL:
                        helAcres = ogCLUinfoDict.get(row[3])[0]
                        nhelAcres = 0.0
                        helPct = 100.0
                        nhelPct = 0.0
                    else:
                        nhelAcres = ogCLUinfoDict.get(row[3])[0]
                        helAcres = 0.0
                        helPct = 0.0
                        nhelPct = 100.0
                else:
                    nhelAcres = outTabulateValues[0] / acreConversion
                    helAcres = float(outTabulateValues[1]) / acreConversion

                    totalAcres =  (outTabulateValues[0] + outTabulateValues[1]) / acreConversion
                    helPct = (helAcres / totalAcres) * 100
                    nhelPct = (nhelAcres / totalAcres) * 100

                # WARNING - New acre percentages for HEL and NHEL areas did not add up exactly to
                # 100% b/c they were being computed from raster cells and divided by polygon totals.
                # This is the original way of computing new HEL and NHEL acres
    ##                helPct = (helAcres / row[4]) * 100
    ##                nhelPct = (nhelAcres / row[4]) * 100

                # set default values
                row[0] = helAcres
                row[1] = helPct
                clu = row[3]

                if helPct > 33.3333 or helAcres > 50.0:
                    row[2] = "Yes"
                else:
                    row[2] = "No"

                #cluDict[clu] = (len(str(clu)),round(helAcres,1),round(helPct,1),len(str(round(helAcres,1))),len(str(round(helPct,1))),"HEL --> " + row[2])
                cluDict[clu] = (round(helAcres,1),len(str(round(helAcres,1))),round(helPct,1),round(nhelAcres,1),len(str(round(nhelAcres,1))),round(nhelPct,1),row[2]) # {13: (2, 2.8, 37.3, 3, 4, ' HEL --> Yes')}
                del expression,outTabulateValues,helAcres,helPct,nhelAcres,nhelPct,clu
                cursor.updateRow(row)
        del cursor

        # Strictly for formatting
        maxHelAcreLength = sorted([cluinfo[1] for clu,cluinfo in cluDict.iteritems()],reverse=True)[0]
        maxNHelAcreLength = sorted([cluinfo[4] for clu,cluinfo in cluDict.iteritems()],reverse=True)[0]
        #maxPercentLength = sorted([cluinfo[4] for clu,cluinfo in cluDict.iteritems()],reverse=True)[0]

        for clu in cluDict:
            firstSpace = " "  * (maxHelAcreLength - cluDict[clu][1])
            secondSpace = " " * (maxNHelAcreLength - cluDict[clu][4])
            helAcres = cluDict[clu][0]
            helPct = cluDict[clu][2]
            nHelAcres = cluDict[clu][3]
            nHelPct = cluDict[clu][5]
            yesOrNo = cluDict[clu][6]

            AddMsgAndPrint("\tCLU #: " + str(clu))
            AddMsgAndPrint("\t\tHEL Acres:  " + str(helAcres) + firstSpace + " .ac -- " + str(helPct) + " %")
            AddMsgAndPrint("\t\tNHEL Acres: " + str(nHelAcres) + secondSpace + " .ac -- " + str(nHelPct) + " %")
            AddMsgAndPrint("\t\tHEL Determination: " + yesOrNo + "\n")
            del firstSpace,secondSpace,helAcres,helPct,nHelAcres,nHelPct,yesOrNo

        del tabulateFields,fieldList,cluDict,maxHelAcreLength,maxNHelAcreLength
        arcpy.SetProgressorPosition()

        """----------------------------------------------------------------------------------------------------- Prepare Symboloby for ArcMap and 1026 form"""
        AddLayersToArcMap()

        if bAccess:
            if not populateForm():
               AddMsgAndPrint("\nFailure to correclty populate 1026 form",2)

        arcpy.SetProgressorLabel("Removing Temp Layers\n\n")
        AddMsgAndPrint("\nRemoving Temp Layers")
        removeScratchLayers()

    except:
        removeScratchLayers()
        errorMsg()