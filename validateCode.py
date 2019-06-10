import arcpy, getpass, os

class ToolValidator(object):
    """Class for validating a tool's parameter values and controlling
    the behavior of the tool's dialog."""

    def __init__(self):
        """Setup arcpy and the list of tool parameters."""
        self.params = arcpy.GetParameterInfo()

    def initializeParameters(self):
        """Refine the properties of a tool's parameters.  This method is
        called when the tool is opened."""
        return

    def updateParameters(self):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""

        # ------------------------------------------------------------------------------------------------------------------
        # 1st Parameter - CLU Layer
        # Auto-populate CLU Layer if named in the conventional way i.e. clu_a_wi001
        # Must be found in ArcMap Table of Contents
        try:
	        mxd = arcpy.mapping.MapDocument("CURRENT")
	        df = arcpy.mapping.ListDataFrames(mxd)[0]

	        cluLayer = [str(x) for x in arcpy.mapping.ListLayers(mxd) if str(x).find('clu_a') > -1][0]

        except:
            cluLayer = ''

        if not self.params[0].hasBeenValidated and not self.params[0].altered:
            if cluLayer:
	            self.params[0].value = cluLayer

        # ------------------------------------------------------------------------------------------------------------------
        # 2nd Parameter - HEL Layer
        # Auto-populate HEL Layer if named in the conventional way i.e. HEL_a_SSA
        # Must be found in ArcMap Table of Contents

        try:
            helLayer = [str(x) for x in arcpy.mapping.ListLayers(mxd) if str(x).find('HEL_a') > -1][0]
        except:
            helLayer = ''

        if not self.params[1].hasBeenValidated and not self.params[1].altered:
	        if helLayer:
	            self.params[1].value = helLayer

    # ------------------------------------------------------------------------------------------------------------------
    # OBSOLETE: 3rd Parameter and subsequent 4 parameters - HEL Layer.
    # Auto-populate the K,T,R and HEL Field if the HEL Layer parameter is filled.
    # These parameters were removed from the tool.  They will be added as checks
    # in the updateMessages function.

##    if not self.params[1].hasBeenValidated or (self.params[1].altered and not self.params[1].hasBeenValidated):
##
##        helLayer = self.params[1].value
##
##	# Kw Factor Field - Some states may have it labeled 'K'
##	try:
##	    self.params[2].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 'k'][0])
##	except:
##         self.params[2].value = ''
##
##	# T Factor Field
##	try:
##	    self.params[3].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 't'][0])
##	except:
##	    self.params[3].value = ''
##
##	# R Factor Field
##	try:
##	    self.params[4].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 'r'][0])
##	except:
##	    self.params[4].value = ''
##
##	# HEL Field
##	try:
##	    self.params[5].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 'muhelcl'][0])
##	except:
##	    self.params[5].value = ''

        # ------------------------------------------------------------------------------------------------------------------
        # 3rd and 4th Parameter - DEM Layer.
        # Auto-populate DEM layer if possible
        # Must be found in ArcMap Table of Contents

        try:
	        rasterLyrs = [x for x in arcpy.mapping.ListLayers(mxd) if arcpy.Describe(x).dataType == 'RasterLayer']
        except:
	        rasterLyrs = ''

        if not self.params[2].hasBeenValidated and not self.params[2].altered:
            dem = ''

            # Only 1 raster found; set it
            if len(rasterLyrs) == 1:
                dem = str(rasterLyrs[0])

            # More than 1 raster in Table of contents
            elif len(rasterLyrs) > 1:

                # Find the first raster with 'dem' in the name
                for raster in rasterLyrs:
	                if str(raster).find('dem') > -1 or str(raster).find('DEM') > -1:
		                dem = raster
		                break

	            # No 'dem' raster found; choose first floating point raster
                if not dem:
                    for raster in rasterLyrs:
                        if arcpy.Describe(raster).pixelType in ('F32','F64'):
                            dem = raster
                            break

            # No Rasters found; set to blank
            else:
                self.params[2].value = ''
                self.params[3].value = ''

            if dem:
                self.params[2].value = str(dem)

                units = arcpy.Describe(dem).spatialReference.LinearUnitName

                if units == "Meter":
                    self.params[3].value = 'Meters'
                elif units == "Foot" or units == "Foot_US":
                    self.params[3].value = 'Feet'
                else:
                    self.params[3].value = ''

    # ------------------------------------------------------------------------------------------------------------------
    # 9th Parameter - DC Signature

#    if not self.params[9].hasBeenValidated and not self.params[9].altered:
#        try:
#            self.params[9].value = getpass.getuser().replace('.',' ').replace('\'','')
#        except:
#	    self.params[9].value = 'DC Name'

        return

    def updateMessages(self):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""

        # Check the CLU layer for the following fields.
        if self.params[0].value:
            cluRequired = ["CLUNBR","TRACTNBR","FARMNBR","COUNTYCD","STATECD"]
            cluFields = [f.name for f in arcpy.ListFields(self.params[0].value)]

            for fld in cluRequired:
                if not len(arcpy.ListFields(self.params[0].value,fld)) > 0:
                    self.params[0].setErrorMessage("\"" + fld + "\" field is missing from CLU Layer")
                    break

        # Check the HEL layer for the following fields.
        if self.params[1].value:
            helRequired = ["K","T","R","MUHELCL"]
            helFields = [f.name for f in arcpy.ListFields(self.params[1].value)]

            for fld in helRequired:
                if not fld in helFields:
                    self.params[1].setErrorMessage("\"" + fld + "\" field is missing from HEL Layer")
                    break

        return
