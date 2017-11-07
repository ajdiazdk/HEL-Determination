import arcpy
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

    # Auto-populate CLU Layer if named in the conventional way i.e. clu_a_wi001
    # Must be found in ArcMap Table of Contents
    try:
        mxd = arcpy.mapping.MapDocument("CURRENT")
	df = arcpy.mapping.ListDataFrames(mxd)[0]

	self.params[1].value = [str(x) for x in arcpy.mapping.ListLayers(mxd) if str(x).find('clu_a') > -1][0]
	
    except:
	self.params[1].value = ''

    # Auto-populate the K,T,R and HEL Field if the HEL Layer parameter is filled.
    if not self.params[2].hasBeenValidated or (self.params[2].altered and not self.params[2].hasBeenValidated):

        helLayer = self.params[2].value

	# K Factor Field
	try:
		self.params[3].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 'k'][0])
	except:
		self.params[3].value = ''

	# T Factor Field
	try:
		self.params[4].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 't'][0])
	except:
		self.params[4].value = ''

	# R Factor Field
	try:
		self.params[5].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 'r'][0])
	except:
		self.params[5].value = ''

	# HEL Field
	try:
		self.params[6].value = str([f.name for f in arcpy.ListFields(helLayer) if (f.name).lower() == 'hel'][0])
	except:
		self.params[6].value = ''

    # Auto-populate DEM layer if possible
    # Must be found in ArcMap Table of Contents
    try:
        mxd = arcpy.mapping.MapDocument("CURRENT")
	df = arcpy.mapping.ListDataFrames(mxd)[0]

	rasterLyrs = [x for x in arcpy.mapping.ListLayers(mxd) if arcpy.Describe(x).dataType == 'RasterLayer']

	if not len(rasterLyrs):
		self.params[7].value = ''

	if len(rasterLyrs) == 1:
		self.params[7].value = str(rasterLyrs[0])
	
	# More than 1 raster in TCO
	if len(rasterLyrs) > 1:
		dem = ''

		# Find the first raster with 'dem' in the name
		for raster in rasterLyrs:
			if str(raster).find('dem') > -1:
				dem = raster
				break

		if dem:
			self.params[7].value = dem

		# No 'dem' raster found; choose first floating point raster
		else:
			for raster in rasterLyrs:
				if arcpy.Describe(raster).pixelType in ('F32','F64'):
					dem = raster
					break
			if dem:
				self.params[7].value = dem
			else:
				self.params[7].value = ''

    except:
	self.params[7].value = ''		

    return

  def updateMessages(self):
    """Modify the messages created by internal validation for each tool
    parameter.  This method is called after internal validation."""
    return
