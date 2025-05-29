"""
Script documentation

- Tool parameters are accessed using arcpy.GetParameter() or 
                                     arcpy.GetParameterAsText()
- Update derived parameter values using arcpy.SetParameter() or
                                        arcpy.SetParameterAsText()
"""
import arcpy as ap
import ctypes
import os
import sys
import time

def script_tool(coord_input, map_name):
    aprx = ap.mp.ArcGISProject('CURRENT')
    m = []
    #selection_point = 'Location_100m'#ap.mp.LayerFile('C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Maps.gdb\\Location_100m')
    index_layer = 'mgrs_index_ftp_link'
    ex_hard_drive = 'D:\\MGRS_US_GZD'
    
    if os.path.exists(ex_hard_drive) == True:
        #ap.env.workspace = ex_hard_drive
        print('D: drive is connected')
    else:
        sys.exit('No Drive Connected')
    
    coord_str = (coord_input)
    coord_str = coord_str.split()
    coord_int = tuple(float(x) for x in coord_str)
    coords = (coord_int)
    coords_geo = ap.PointGeometry(ap.Point(coords[0], coords[1]))
    coords_array = [coords_geo.firstPoint.X, coords_geo.firstPoint.Y]
    
    GZD_00X = []
    mgrs_100kmSq = []
    mgrs_10km = []
    mgrs_1km = []
    mgrs_100m = []
    mgrs_00XX123456 = []
    
    def new_map(map_name):
        names = []
        for map in aprx.listMaps():
            names.append(map.name)
        try:
            for name in names:
                if map_name == name:
                    raise Exception('Map already exists')
        except:
            ctypes.windll.user32.MessageBoxW(0, "Map already exists. You must delete or rename the map with the duplicate title before trying to create a new map.", 'Warning', 1)
            sys.exit(1)
        else:
            aprx.createMap(map_name)
            global m
            m = aprx.listMaps(map_name)[0]
            m.openView()
            
            m.addBasemap('Imagery')
            basemap = ['World Imagery']
            bm = m.listLayers(basemap[0])[0]
            bm.visible = False
            m.addLayer(ap.mp.LayerFile('C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\Cartography_Layers\\Enhanced_Contrast_Base_(WGS84).lyrx'), add_position='BOTTOM')
            
            m.addDataFromPath('D:\\MGRS_US_GZD\\mgrs_index_ftp_link\\mgrs_index_ftp_link.shp')
            
            def change_symbology_std(layer, style):
                # Modify symbology
                symbology = layer.symbology
                if hasattr(symbology, "renderer"):
                    symbology.updateRenderer("SimpleRenderer")
                    symbology.renderer.symbol.applySymbolFromGallery(style)
                    layer.symbology = symbology  # Apply changes
                return
    
            change_symbology_std(m.listLayers(index_layer)[0], 'std_grid_line')
            return print('map created')   
        return
        
    def create_point(coords):
        aprx = ap.mp.ArcGISProject('CURRENT')
        m = aprx.listMaps(map_name)[0]
        template = ap.mp.Table('C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Maps.gdb\\MGRS_Template')
        
        ap.CreateTable_management(out_path='C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Maps.gdb', out_name='MGRS_selectTab', template=template)
        with ap.da.InsertCursor('MGRS_selectTab', ['lat', 'lng']) as cursor:
            cursor.insertRow(coords_array)
        ap.management.XYTableToPoint(in_table='MGRS_selectTab', y_field='lat', x_field='lng', out_feature_class='C:\\GIS\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Maps.gdb\\Location_100m')
        add_table = ap.mp.Table('C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Maps.gdb\\MGRS_selectTab')
        
        for lyr in m.listLayers():
            if lyr.name == 'Location_100m':
                print(True)
                break
            else:
                print('not here')
                m.addDataFromPath('C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Maps.gdb\\Location_100m')
                time.sleep(1)
                break
    
        if not m.listTables():
            m.addTable(add_table)
            print('table added')
        else:
            print('not table added')
        return
        
    def find_and_add_GZD(selection_point, search_layer, search_field):
        #aprx = ap.mp.ArcGISProject('CURRENT')
        m = aprx.listMaps(map_name)[0]
        ap.env.workspace = ex_hard_drive
        def locate():
            with ap.da.SearchCursor(search_layer, search_field) as cursor:
                for row in cursor:
                    return row[0]
    
        ap.management.SelectLayerByLocation(in_layer=search_layer, overlap_type='CONTAINS', select_features=selection_point)
        mgrs_folder = '\\mgrs_' + locate().lower() + '_100m'
        mgrs_folder_gdb = '\\MGRS_' + locate() + '_100m_esri93.gdb'
        MGRS_100kmSQ_file = '\\MGRS_100kmSQ_ID_' + locate()
        m.addDataFromPath(ap.env.workspace + mgrs_folder + mgrs_folder_gdb + MGRS_100kmSQ_file)
        GZD_00X.append(locate())
        mgrs_100kmSq.append(MGRS_100kmSQ_file.removeprefix('\\'))
        m.clearSelection()
    
        def change_symbology_std(layer, style):
            # Modify symbology
            symbology = layer.symbology
            if hasattr(symbology, "renderer"):
                symbology.updateRenderer("SimpleRenderer")
                symbology.renderer.symbol.applySymbolFromGallery(style)
                layer.symbology = symbology  # Apply changes
                return
    
        change_symbology_std(m.listLayers(mgrs_100kmSq[0])[0], 'std_grid_line')
        
        return
    
    def find_and_add_100m(selection_point, search_layer, search_field):
        #aprx = ap.mp.ArcGISProject('CURRENT')
        m = aprx.listMaps(map_name)[0]
        ap.env.workspace = ex_hard_drive
        def locate():
            with ap.da.SearchCursor(search_layer, search_field) as cursor:
                for row in cursor:
                    return row[0]
        
        ap.management.SelectLayerByLocation(in_layer=search_layer, overlap_type='CONTAINS', select_features=selection_point)
        mgrs_folder = '\\mgrs_' + GZD_00X[0].lower() + '_100m'
        mgrs_folder_gdb = '\\MGRS_' + GZD_00X[0] + '_100m_esri93.gdb'
        MGRS_100m_file = '\\MGRS_' + locate() + '_100m'
        m.addDataFromPath(ap.env.workspace + mgrs_folder + mgrs_folder_gdb + MGRS_100m_file)
        m.clearSelection()
        mgrs_100m.append(MGRS_100m_file.removeprefix('\\'))
    
        def change_symbology_100m(layer, style):
            # Modify symbology
            symbology = layer.symbology
            if hasattr(symbology, "renderer"):
                symbology.updateRenderer("SimpleRenderer")
                symbology.renderer.symbol.applySymbolFromGallery(style)
                layer.symbology = symbology  # Apply changes
                layer.maxThreshold = 0
                layer.minThreshold = 24000
                return
    
        change_symbology_100m(m.listLayers(mgrs_100m[0])[0], 'MGRS_100m')
        
        return
    
    def adjust_spatial_ref(reference_layer):  #'MGRS_00XXX_100m'   
        aprx = ap.mp.ArcGISProject('CURRENT')
        m = aprx.listMaps(map_name)[0]
        #get name of reference_layer
        desc = ap.Describe(reference_layer)
        spatial_ref = desc.spatialReference.name
        new_spatial_ref = ap.SpatialReference(spatial_ref)
        
        #change basemap spatial reference
        m.spatialReference = new_spatial_ref
        return
    
    
    def add_10_and_1_km(selection_point, search_layer, search_field):
        #aprx = ap.mp.ArcGISProject('CURRENT')
        m = aprx.listMaps(map_name)[0]
        ap.env.workspace = 'D:\\MGRS_US_GZD'
        def locate():
            with ap.da.SearchCursor(search_layer, search_field) as cursor:
                for row in cursor:
                    return row[0]
    
        ap.management.SelectLayerByLocation(in_layer=search_layer, overlap_type='CONTAINS', select_features=selection_point)
        mgrs_folder = '\\mgrs_' + locate().lower() + '_100m'
        mgrs_folder_gdb = '\\MGRS_' + locate() + '_100m_esri93.gdb'
        MGRS_10k_file = '\\MGRS_' + locate() + '_10km'
        MGRS_1k_file = '\\MGRS_' + locate() + '_1km'
        m.addDataFromPath(ap.env.workspace + mgrs_folder + mgrs_folder_gdb + MGRS_10k_file)
        m.addDataFromPath(ap.env.workspace + mgrs_folder + mgrs_folder_gdb + MGRS_1k_file)
        m.clearSelection()
        mgrs_10km.append(MGRS_10k_file.removeprefix('\\'))
        mgrs_1km.append(MGRS_1k_file.removeprefix('\\'))
    
        def change_symbology_10km(layer, style):
            #
            symbology = layer.symbology
            if hasattr(symbology, "renderer"):
                symbology.updateRenderer("SimpleRenderer")
                symbology.renderer.symbol.applySymbolFromGallery(style)
                layer.symbology = symbology  # Apply changes
                layer.maxThreshold = 100000
                layer.minThreshold = 1000000
            return
    
        change_symbology_10km(m.listLayers(mgrs_10km[0])[0], 'std_grid_line')
        
        def change_symbology_1km(layer, style):    
            #
            symbology = layer.symbology
            if hasattr(symbology, "renderer"):
                symbology.updateRenderer("SimpleRenderer")
                symbology.renderer.symbol.applySymbolFromGallery(style)
                layer.symbology = symbology  # Apply changes
                layer.minThreshold = 200000
            return
        change_symbology_1km(m.listLayers(mgrs_1km[0])[0], 'std_grid_line')
        return
    
    def find_100m_square(selection_point):
        #aprx = ap.mp.ArcGISProject('CURRENT')
        m = aprx.listMaps(map_name)[0]
        ap.management.SelectLayerByLocation(in_layer=m.listLayers()[1], overlap_type='CONTAINS', select_features=selection_point)
        
        with ap.da.SearchCursor(in_table=m.listLayers()[1], field_names='LABEL') as cursor:
            for row in cursor:
                mgrs_00XX123456.append(row[0])
        
        with ap.da.UpdateCursor(m.listLayers()[0], ['lat', 'lng', 'MGRS_100m']) as cursor:
                for row in cursor:
                    cursor.updateRow([coords_array[0], coords_array[1], mgrs_00XX123456[0]])
        return
    
    def group_layers():
        aprx = ap.mp.ArcGISProject('CURRENT')
        m = aprx.listMaps(map_name)[0]
        group_layer_name = 'MGRS_' + GZD_00X[0]
        add_group_layer = m.createGroupLayer(group_layer_name)
        group_layer = (m.listLayers()[0])
        add_group_layer
        layers = m.listLayers()
        
        for layer in layers:
            if layer.isFeatureLayer:
                m.addLayerToGroup(target_group_layer=group_layer, add_layer_or_layerfile=layer, add_position='BOTTOM')
                m.removeLayer(remove_layer=layer)
        return
    
    new_map(map_name) #map_name
    
    create_point(coords)
    
    #aprx = ap.mp.ArcGISProject('CURRENT')
    m = aprx.listMaps(map_name)[0]
    find_and_add_GZD(m.listLayers()[0], m.listLayers()[1], 'MGRS_UTM')
    add_10_and_1_km(m.listLayers()[0], m.listLayers()[2], 'MGRS_UTM')
    find_and_add_100m(m.listLayers()[0], m.listLayers()[3], 'MGRS')
    
    adjust_spatial_ref(m.listLayers(mgrs_100m[0])[0])
    
    aprx = ap.mp.ArcGISProject('CURRENT')
    m = aprx.listMaps(map_name)[0]
    find_100m_square(m.listLayers()[0])
    
    group_layers()
    ap.env.workspace = 'C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Maps.gdb'
    return


def create_layout(map_name, cam_scale):
    aprx = ap.mp.ArcGISProject('CURRENT')
    new_layout = ""
    m = ""
    #Import layout template
    aprx.importDocument("C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS_Layout_Template.pagx")
    
    #get and rename layout by matching map name
    for lyt in aprx.listLayouts():
        if lyt.name == "Layout":
            lyt.name = map_name
            new_layout = lyt
    
    #map
    m = aprx.listMaps(map_name)[0]
    #map frame
    mf = new_layout.listElements("MAPFRAME_ELEMENT")[0]
    
    #get point for extent reference
    point_layer = ""
    for lyr in m.listLayers():
        if lyr.name == "Location_100m":
            point_layer = lyr            
    
    def create_point():
        # Use a SearchCursor to grab the first feature's geometry
        with arcpy.da.SearchCursor(point_layer, ["SHAPE@"]) as cursor:
            for row in cursor:
                point_geom = row[0]  # This is an arcpy.PointGeometry
                #print(f"X: {point_geom.centroid.X}, Y: {point_geom.centroid.Y}")
                return point_geom # Only grab the first one

    create_point()
    
    #change map in mapframe from <None> to newly created map
    mf.map = m
    
    mf.camera.setExtent(create_point().extent)
    
    #fuction used to get the layout element
    def edit_layout_elements(element_title):
        for e in new_layout.listElements():
                if e.name == element_title:
                    return e
    
    coords_label = edit_layout_elements("point_coordinates")
    MGRS_100m_label = edit_layout_elements("MGRS_100m_label")
    
    #dynamic text URI element
    location_cim_path = f'"CIMPATH={m.name}/Location_100m2.json"'
    
    #update dynamic text to reflect the new map
    coords_label.text = f'<dyn type="table" property="value" mapFrame="MGRS_grid_labeled" mapMemberUri={location_cim_path} isDynamic="true" arcade="$feature.lat + &quot;, &quot; + $feature.lng"  delimiter=" "/>'
    MGRS_100m_label.text = f'<dyn type="table" property="value" mapFrame="MGRS_grid_labeled" mapMemberUri={location_cim_path} isDynamic="true" field="MGRS_100m" sortField="MGRS_100m" sortAscending="true" delimiter=" "/> (100m)'
    #set map frame scale with cam_scale variable
    mf.camera.scale = cam_scale
    
    grid_style = ""
    for style in aprx.listStyleItems("C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\MGRS.stylx"):
        if style.name == "1000m__MGRS_Grid":
            grid_style = style
    
    mf.removeGrids()
    mf.addGrid(grid_style)
    new_layout.openView()
    
    def change_symbology_lytpnt(layer, style):
        # Modify symbology
        symbology = layer.symbology
        if hasattr(symbology, "renderer"):
            symbology.updateRenderer("SimpleRenderer")
            symbology.renderer.symbol.applySymbolFromGallery(style)
            layer.symbology = symbology  # Apply changes
            return
            
    def rename_Location_100m(map):
        point_layer.name = f"Location_100_{map}"               
        ap.SaveToLayerFile_management(in_layer=point_layer, out_layer=f"C:\\GIS\\ArcGIS\\Projects\\MGRS_Maps\\Cartography_Layers\\Location_100m_layers\\Location_100_{map}.lyrx")
    
    vis_layers = m.listLayers()
    for lyr in vis_layers:
        lyr.visible = False
    vis_layers[0].visible = True
    vis_layers[1].visible = True
    vis_layers[8].visible = True
                          
    change_symbology_lytpnt(point_layer, "Reference_Point")
    rename_Location_100m(map_name)

if __name__ == "__main__":

    coord_input = ap.GetParameter(0)
    map_name = ap.GetParameter(1)
    create_lyt = ap.GetParameter(2)

    script_tool(coord_input, map_name)
    #ap.SetParameterAsText(2, result)
    if create_lyt == True:
        create_layout(map_name, 100000)
