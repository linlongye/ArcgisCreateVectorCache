# coding=utf-8
import os
import sys
import time

import arcpy


class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "切片工具箱"
        self.alias = "切片工具箱"

        # List of tool classes associated with this toolbox
        self.tools = [CacheTool,CacheToolVector]


class CacheTool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "制作切片"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        param1 = arcpy.Parameter(
            displayName="MXD文档位置",
            name="mxd_path",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        param1.filter.list = ['mxd']
        params = [param1]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        agspath = "C:/Users/Mr.HL/AppData/Roaming/ESRI/Desktop10.2/ArcCatalog/arcgis on localhost_6080 (admin).ags"
        mxdpath = parameters[0].valueAsText
        new_mxd = arcpy.mapping.MapDocument(mxdpath)
        dirname = os.path.dirname(mxdpath)
        mxdname = os.path.basename(mxdpath)
        dotindex = mxdname.index('.')
        servicename = mxdname[0:dotindex]

        messages.addMessage('service_name {0}'.format(servicename))

        sddratf = os.path.abspath(servicename + ".sddraft")
        sd = os.path.abspath(servicename + ".sd")
        if os.path.exists(sd):
            os.remove(sd)

        arcpy.CreateImageSDDraft(new_mxd, sddratf, servicename, "ARCGIS_SERVER", agspath, False, None, "Ortho Images",
                                 "dddkk")
        analysis = arcpy.mapping.AnalyzeForSD(sddratf)

        # 打印分析结果
        messages.addMessage("the following information wa returned during analysis of the MXD")
        for key in ('messages', 'warnings', 'errors'):
            messages.addMessage('---- {0} ----'.format(key.upper()))
            vars = analysis[key]
            for ((message, code), layerlist) in vars.iteritems():
                messages.addMessage(' message (CODE {0})    applies to:'.format(code))
                for layer in layerlist:
                    messages.addMessage(layer.name)

        if analysis['errors'] == {}:
            try:
                arcpy.StageService_server(sddratf, sd)
                arcpy.UploadServiceDefinition_server(sd, agspath)
                messages.addMessage('service successfully published')

                # 定义服务器缓存
                messages.addMessage("开始定义服务器缓存")
                # list of input variable for map service properties
                tiling_scheme_type = "NEW"
                scale_type = "CUSTOM"
                num_of_scales = 9
                scales = [7688, 7200, 6600, 6000, 5400, 4800, 4200, 3600, 3000]
                dot_per_inch = 96
                tile_origin = "-20037700 30198300"  # <X>-20037700</X> <Y>30198300</Y>
                tile_size = "256 x 256"
                cache_tile_format = "PNG8"
                tile_compression_quality = ""
                storage_format = "EXPLODED"
                predefined_tile_scheme = ""
                input_service = 'GIS Servers/arcgis on localhost_6080 (admin)/{0}.MapServer'.format(servicename)
                cache_path = "D:/arcgisserver/directories-new/arcgiscache"

                try:
                    start_time = time.clock()
                    result = arcpy.CreateMapServerCache_server(input_service, cache_path, tiling_scheme_type,
                                                               scale_type,
                                                               num_of_scales, dot_per_inch, tile_size,
                                                               predefined_tile_scheme, tile_origin, scales,
                                                               cache_tile_format, tile_compression_quality,
                                                               storage_format)
                    # message to a file
                    while result.status < 4:
                        time.sleep(0.2)
                    # result_value = result.getMessage()
                    # messages.addMessage('completed {0}'.format(str(result_value)))
                except Exception, e:
                    tb = sys.exc_info()[2]
                    messages.addMessage('Failed to stop 1 \n Line {0}'.format(tb.tb_lineno))
                    messages.addMessage(e.message)

                # 生成切片
                messages.addMessage("开始切片")
                scales = [7688, 7200, 6600, 6000, 5400, 4800, 4200, 3600, 3000]
                num_of_caching_service_instance = 2
                update_mode = "RECREATE_ALL_TILES"
                area_of_interest = ""
                wait_for_job_completion = "WAIT"
                update_extents = ""
                input_service = 'C:/Users/Mr.HL/AppData/Roaming/ESRI/Desktop10.2/ArcCatalog/arcgis on localhost_6080 (' \
                                'admin)/{0}.MapServer'.format(servicename)

                try:
                    result = arcpy.ManageMapServerCacheTiles_server(input_service, scales, update_mode,
                                                                    num_of_caching_service_instance, area_of_interest,
                                                                    update_extents, wait_for_job_completion)
                    while result.status < 4:
                        time.sleep(0.2)
                    # result_value = result.getMessage()
                    # messages.addMessage('completed {0}'.format(str(result_value)))
                    messages.addMessage('Created cache tiles for given schema successfully')
                except Exception, e:
                    tb = sys.exc_info()[2]
                    messages.addMessage('Failed at step 1 \n Line {0}'.format(tb.tb_lineno))
                    messages.addMessage(e.message)
                messages.addMessage("Created Map Service Cache Tiles")
            except Exception, e:
                messages.addMessage(e.message)
        else:
            messages.addMessage('Service could not be published because errors were found during analysis')

        return




class CacheToolVector(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "制作矢量切片"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Define parameter definitions"""
        param1 = arcpy.Parameter(
            displayName="MXD文档位置",
            name="mxd_path",
            datatype="DEFile",
            parameterType="Required",
            direction="Input"
        )
        param1.filter.list = ['mxd']
        params = [param1]
        return params

    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages):
        """The source code of the tool."""
        agspath = "C:/Users/Mr.HL/AppData/Roaming/ESRI/Desktop10.2/ArcCatalog/arcgis on localhost_6080 (admin).ags"
        mxdpath = parameters[0].valueAsText
        new_mxd = arcpy.mapping.MapDocument(mxdpath)
        dirname = os.path.dirname(mxdpath)
        mxdname = os.path.basename(mxdpath)
        dotindex = mxdname.index('.')
        servicename = mxdname[0:dotindex]

        messages.addMessage('service_name {0}'.format(servicename))

        sddratf = os.path.abspath(servicename + ".sddraft")
        sd = os.path.abspath(servicename + ".sd")
        if os.path.exists(sd):
            os.remove(sd)

        arcpy.CreateImageSDDraft(new_mxd, sddratf, servicename, "ARCGIS_SERVER", agspath, False, None, "Ortho Images",
                                 "dddkk")
        analysis = arcpy.mapping.AnalyzeForSD(sddratf)

        # 打印分析结果
        messages.addMessage("the following information wa returned during analysis of the MXD")
        for key in ('messages', 'warnings', 'errors'):
            messages.addMessage('---- {0} ----'.format(key.upper()))
            vars = analysis[key]
            for ((message, code), layerlist) in vars.iteritems():
                messages.addMessage(' message (CODE {0})    applies to:'.format(code))
                for layer in layerlist:
                    messages.addMessage(layer.name)

        if analysis['errors'] == {}:
            try:
                arcpy.StageService_server(sddratf, sd)
                arcpy.UploadServiceDefinition_server(sd, agspath)
                messages.addMessage('service successfully published')

                # 定义服务器缓存
                messages.addMessage("开始定义服务器缓存")
                # list of input variable for map service properties
                tiling_scheme_type = "NEW"
                scale_type = "CUSTOM"
                num_of_scales = 9
                # 根据实际情况修改比例尺参数
                scales = [8000000, 4000000, 2000000, 1000000, 500000, 250000, 125000, 64000, 32000]
                dot_per_inch = 96
                tile_origin = "-400 400"  # <X>-20037700</X> <Y>30198300</Y>
                tile_size = "256 x 256"
                cache_tile_format = "PNG8"
                tile_compression_quality = ""
                storage_format = "EXPLODED"
                predefined_tile_scheme = ""
                input_service = 'GIS Servers/arcgis on localhost_6080 (admin)/{0}.MapServer'.format(servicename)
                cache_path = "D:/arcgisserver/directories-new/arcgiscache"

                try:
                    start_time = time.clock()
                    result = arcpy.CreateMapServerCache_server(input_service, cache_path, tiling_scheme_type,
                                                               scale_type,
                                                               num_of_scales, dot_per_inch, tile_size,
                                                               predefined_tile_scheme, tile_origin, scales,
                                                               cache_tile_format, tile_compression_quality,
                                                               storage_format)
                    # message to a file
                    while result.status < 4:
                        time.sleep(0.2)
                    # result_value = result.getMessage()
                    # messages.addMessage('completed {0}'.format(str(result_value)))
                except Exception, e:
                    tb = sys.exc_info()[2]
                    messages.addMessage('Failed to stop 1 \n Line {0}'.format(tb.tb_lineno))
                    messages.addMessage(e.message)

                # 生成切片
                messages.addMessage("开始切片")
                # 根据实际情况修改比例尺参数
                scales = scales = [8000000, 4000000, 2000000, 1000000, 500000, 250000, 125000, 64000, 32000]
                num_of_caching_service_instance = 2
                update_mode = "RECREATE_ALL_TILES"
                area_of_interest = ""
                wait_for_job_completion = "WAIT"
                update_extents = ""
                input_service = 'C:/Users/Mr.HL/AppData/Roaming/ESRI/Desktop10.2/ArcCatalog/arcgis on localhost_6080 (' \
                                'admin)/{0}.MapServer'.format(servicename)

                try:
                    result = arcpy.ManageMapServerCacheTiles_server(input_service, scales, update_mode,
                                                                    num_of_caching_service_instance, area_of_interest,
                                                                    update_extents, wait_for_job_completion)
                    while result.status < 4:
                        time.sleep(0.2)
                    # result_value = result.getMessage()
                    # messages.addMessage('completed {0}'.format(str(result_value)))
                    messages.addMessage('Created cache tiles for given schema successfully')
                except Exception, e:
                    tb = sys.exc_info()[2]
                    messages.addMessage('Failed at step 1 \n Line {0}'.format(tb.tb_lineno))
                    messages.addMessage(e.message)
                messages.addMessage("Created Map Service Cache Tiles")
            except Exception, e:
                messages.addMessage(e.message)
        else:
            messages.addMessage('Service could not be published because errors were found during analysis')

        return