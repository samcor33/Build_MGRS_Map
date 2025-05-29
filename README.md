This code streamlines the production of an MGRS ArcGIS Pro map and (optional) layout.
By entering a set of coordinates in decimal format, the tool will:
- Create a new map and place a point at the location coordinates.
- Pull all the relevant MGRS layers from a local file repository (external hard drive) down to the 100m grid area.
- Set the spatial reference to the appropriate UTM projection.
- Identify the 100m grid location for quick identification of the point in MGRS format.
- (Optional) Create a layout of the map ready for export with proper grid markings and relevant cartography elements.

This tool was created using ArcPy 3.x and only works in ArcGIS Pro. It was developed entirely by me and has not been made compatible to download for other users. There is no promise that this is the best or most efficient code to perform the desired output. This code was designed as a project to develop my coding abilities using arcpy and showcase my skills.


https://github.com/user-attachments/assets/47ac6076-aef1-466f-adc9-7813d16f24f3


Printed Layout - [tucson_az.pdf](https://github.com/user-attachments/files/19951800/tucson_az.pdf)


![map_screenshot](https://github.com/user-attachments/assets/35fd3713-2985-4cff-ab18-9dcd49114604)
