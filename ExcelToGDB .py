# -*- coding: cp936 -*-
import arcpy
import xlrd

# 通过arcpy工具箱的方式获取文件路径
xls_path=arcpy.GetParameterAsText(0)
gdb_path=arcpy.GetParameterAsText(1)

xls_file = xlrd.open_workbook(xls_path)
sheet_num = len(xls_file.sheets())
sheet = xls_file.sheet_by_index(0)

arcpy.env.workspace = gdb_path

# shp文件名,坐标参考
# 下划线用空格代替
# 如CGCS2000_3_Degree_GK_Zone_36应为CGCS2000 3 Degree GK Zone 36

def create_shapefile(file_name, coordinate_system, geometry_type):
    spatial_reference = arcpy.SpatialReference(coordinate_system)
    arcpy.CreateFeatureclass_management(arcpy.env.workspace, file_name, geometry_type, spatial_reference=spatial_reference)

def add_fields_to_feature_class(fc,table):
    # 获取要素类的基本字段
    fields =[f.name.lower() for f in arcpy.ListFields(fc)]

    # 遍历Excel表格中的字段
    for ro in range(1,table.nrows):
        rowV=table.row_values(ro)
        field_name =rowV[0]
        field_aliasname = rowV[1]
        field_type = rowV[2]
        field_precision = rowV[3]
        field_is_nullable = rowV[4]
        field_length = rowV[5]
        field_scale = rowV[6]
        tmp=field_name.lower()
        if tmp in fields:
            continue;
        else:
            arcpy.AddField_management(fc, field_name,field_type,field_precision,field_scale,field_length,field_aliasname,field_is_nullable)            

    # 更新要素类
    # arcpy.Refresh(fc)



name="chushihua"

for i in range(sheet_num):
    sheet = xls_file.sheet_by_index(i)
    if i % 2 != 0:
        add_fields_to_feature_class(name,sheet)
    else:
        for j in range(1, sheet.nrows):
            row=sheet.row_values(j)
            name = row[0]
            # coordinate_system ="CGCS2000 3 Degree GK Zone 36"
            coordinate_system =row[1].replace('_'," ")
            geometry_type=row[2]
            create_shapefile(name, coordinate_system,geometry_type)
            
    
