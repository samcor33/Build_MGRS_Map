This code streamlines the production of an MGRS ArcGIS Pro map and (optional) layout. By entering a set of coordinates in decimal format, the tool will:
- Create a new map and place a point at the location coordinates.
- Pull all the relevant MGRS layers from an local file repository (external hard drive) down to the 100m grid area.
- Set the spatial reference to the appropriate UTM projection.
- Identify the 100m grid location for quick identification of the point in MGRS format.
- (Optional) Create a layout of the map ready for export with proper grid markings and relevant cartography elements.

This tool was created using arcpy 3.x and only works in ArcGIS Pro. It was deleveloped entirely by me and has not been made compatible to download for other users. There is no promise that this is the best or most efficient code to perform the desired output. This code was desinged as a project to develop my coding abilities using arcpy and showcase my skills.
