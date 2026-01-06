# Author: Chris Sostad, updated by Marina Cunningham
# chris.sostad@gov.bc.ca
# Ministry of WLRS
# Created Date: February 21st, 2024
# Updated Date: September 09, 2025



# Author: Chris Sostade
# Ministry, Division, Branch: WLRS, Fish and WIldlife Authorizations
# Created Date: February 21st, 2024
# Updated By: Marina Cunningham
# Updated Date: September 09, 2025

# Description:
 
'''
This ArcGIS Pro Python Toolbox provides three tools for quickly exporting map layouts to PDF or JPEG formats:

Export Layout From Project With Only One Layout - Automatically exports the single layout in your project with customizable resolution 

(150-600 DPI), format choice (PDF/JPEG), and optional georeferencing information.

From Multiple Layouts Export Single Layout - Allows you to select one layout from multiple layouts in your project and export 

it with the same customization options.

Export Multiple Layouts to Single File - Exports multiple selected layouts either into a single merged PDF document or as individual JPEGs in a folder, with options to control resolution and georeferencing.

All tools support High/Medium/Low resolution presets, automatic file extension handling, and the ability to include or remove georeferencing information (TFW and AUX.XML files). Overwrite is enabled by default for quick iterations.
'''

# - INPUTS - Layouts!

# - OUTPUTS - PDF or JPEG files

# --------------------------------------------------------------------------------
# * IMPROVEMENTS
# Stability of exporting multiple layouts to a single PDF could be improved.
# Move the field for searching for a directory to the bottom portion of each tool so they are more similar
#
# * Suggestions...
# Change the ordering of the fields
# --------------------------------------------------------------------------------

import arcpy
import traceback
import sys
import os

############################################################################################################################################################################################
#
# Global Functions used in all of the Export Scripts
#
############################################################################################################################################################################################
def exportLayout(layout, output_path, file_name, resolution, format_type, include_geo, messages):
    resolution_map = {
        "High (600 DPI)": 600,
        "Medium (300 DPI)": 300,
        "Low (150 DPI)": 150
    }
    dpi = resolution_map.get(resolution, 300)
    full_path = os.path.join(output_path, file_name)

    if format_type == "PDF":
        layout.exportToPDF(
            out_pdf=full_path,
            resolution=dpi,
            image_quality="BETTER",
            jpeg_compression_quality=80,
            image_compression="ADAPTIVE"
        )
    elif format_type == "JPEG":
        layout.exportToJPEG(
            out_jpg=full_path,
            resolution=dpi,
            jpeg_quality=80
        )

    # Remove georeferencing files if requested
    if not include_geo:
        base = os.path.splitext(full_path)[0]
        tfw = base + ".tfw"
        aux = full_path + ".aux.xml"
        for f in [tfw, aux]:
            if os.path.exists(f):
                os.remove(f)
                messages.addMessage(f"Removed georeferencing file: {f}")
        messages.addMessage("Georeferencing info removed.")
    else:
        messages.addMessage("Georeferencing info retained.")

############################################################################################################################################################################################
#
# Toolbox Definition
#
############################################################################################################################################################################################

class Toolbox(object):
    def __init__(self):
        self.label = "Toolbox"
        self.alias = ""
        self.tools = [ExportSingleLayout, FromMultipleExportSingleLayout, ExportMultipleLayoutsToSingleFile]

############################################################################################################################################################################################
#
# ExportSingleLayout Tool
#
############################################################################################################################################################################################

class ExportSingleLayout(object):
    def __init__(self):
        self.label = "Export Layout From Project With Only One Layout"
        self.description = "Export a selected layout from your project to a PDF or JPEG."
        self.canRunInBackground = False

    def getParameterInfo(self):
        workSpace = arcpy.Parameter(
            displayName="Navigate to the folder where you want to save your file (Warning: Overwrite set to true!)",
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )

        fileName = arcpy.Parameter(
            displayName="File name you want for your output",
            name="fileName",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        resolutionParam = arcpy.Parameter(
            displayName="Select vector resolution",
            name="vector_resolution",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        resolutionParam.filter.type = "ValueList"
        resolutionParam.filter.list = ["High (600 DPI)", "Medium (300 DPI)", "Low (150 DPI)"]
        resolutionParam.value = "Medium (300 DPI)"

        formatParam = arcpy.Parameter(
            displayName="Select export format",
            name="export_format",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        formatParam.filter.type = "ValueList"
        formatParam.filter.list = ["PDF", "JPEG"]
        formatParam.value = "PDF"

        geoParam = arcpy.Parameter(
            displayName="Include Georeferencing Information",
            name="export_geo",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        geoParam.value = True #Checked by default

        return [workSpace, fileName, resolutionParam, formatParam, geoParam]

    def updateParameters(self, parameters):
        if not parameters[1].altered:
            aprx = arcpy.mp.ArcGISProject("CURRENT")
            layouts = aprx.listLayouts()
            if len(layouts) == 1:
                parameters[1].value = f"{layouts[0].name}.pdf"
        return

    def execute(self, parameters, messages):
        try:
            workSpace_path = parameters[0].valueAsText
            file_name = parameters[1].valueAsText
            resolution_level = parameters[2].valueAsText
            format_type = parameters[3].valueAsText
            include_geo = parameters[4].value

            aprx = arcpy.mp.ArcGISProject("CURRENT")
            layout = aprx.listLayouts()[0]
            arcpy.env.overwriteOutput = True

            if not file_name.lower().endswith(f".{format_type.lower()}"):
                file_name += f".{format_type.lower()}"

            exportLayout(layout, workSpace_path, file_name, resolution_level, format_type, include_geo, messages)

        except arcpy.ExecuteError:
            arcpy.AddError(arcpy.GetMessages(2))
        except:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = ("PYTHON ERRORS:\nTraceback info:\n" + tbinfo +
                     "\nError Info:\n" + str(sys.exc_info()[1]))
            msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)

############################################################################################################################################################################################
#
# FromMultipleExportSingleLayout Tool
#
############################################################################################################################################################################################

class FromMultipleExportSingleLayout(object):
    def __init__(self):
        self.label = "From Multiple Layouts Export Single Layout"
        self.description = "Choose from multiple layouts and export a selected layout to PDF or JPEG."
        self.canRunInBackground = False

    def getParameterInfo(self):
        layoutList = arcpy.Parameter(
            displayName="Select a layout from a list of layouts",
            name="layoutList",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        layoutList.filter.type = "ValueList"
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        layouts = aprx.listLayouts()
        layoutList.filter.list = [layout.name for layout in layouts]

        fileName = arcpy.Parameter(
            displayName="File name you want for your output",
            name="fileName",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        fileName.value = "ChangeMe.pdf"

        workSpace = arcpy.Parameter(
            displayName="Navigate to the folder where you want to save your file (Warning: Overwrite set to true!)",
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )
        workSpace.filter.list = ["Local Database", "File System"]

        resolutionParam = arcpy.Parameter(
            displayName="Select vector resolution",
            name="vector_resolution",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        resolutionParam.filter.type = "ValueList"
        resolutionParam.filter.list = ["High (600 DPI)", "Medium (300 DPI)", "Low (150 DPI)"]
        resolutionParam.value = "Medium (300 DPI)"

        formatParam = arcpy.Parameter(
            displayName="Select export format",
            name="export_format",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        formatParam.filter.type = "ValueList"
        formatParam.filter.list = ["PDF", "JPEG"]
        formatParam.value = "PDF"

        geoParam = arcpy.Parameter(
            displayName="Include Georeferencing Information",
            name="export_geo",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        geoParam.value = True

        return [layoutList, fileName, workSpace, resolutionParam, formatParam, geoParam]

                        # #def updateParameters(self, parameters):
                        #     if parameters[0].altered:
                        #         parameters[1].value = parameters[0].value + ".pdf"
                        #     return

    def execute(self, parameters, messages):
        try:
            selected_layout_name = parameters[0].valueAsText
            file_name = parameters[1].valueAsText
            workSpace_path = parameters[2].valueAsText
            resolution_level = parameters[3].valueAsText
            format_type = parameters[4].valueAsText
            include_geo = parameters[5].value

            if not file_name.lower().endswith(f".{format_type.lower()}"):
                file_name += f".{format_type.lower()}"

            aprx = arcpy.mp.ArcGISProject("CURRENT")
            layout = aprx.listLayouts(selected_layout_name)[0]
            arcpy.env.overwriteOutput = True

            exportLayout(layout, workSpace_path, file_name, resolution_level, format_type, include_geo, messages)

            arcpy.AddMessage(f"Layout '{selected_layout_name}' exported to {workSpace_path}")

        except arcpy.ExecuteError:
            arcpy.AddError(arcpy.GetMessages(2))
        except:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
            msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)

############################################################################################################################################################################################
#
# ExportMultipleLayoutsToSingleFile Tool
#
############################################################################################################################################################################################

class ExportMultipleLayoutsToSingleFile(object):
    def __init__(self):
        self.label = "Export Multiple Layouts to Single File"
        self.description = "Export multiple layouts into a single PDF or a folder of JPEGs."
        self.canRunInBackground = False

    def getParameterInfo(self):
        layoutList = arcpy.Parameter(
            displayName="Select layouts to export",
            name="layoutList",
            datatype="GPString",
            parameterType="Required",
            direction="Input",
            multiValue=True
        )

        formatParam = arcpy.Parameter(
            displayName="Select export format",
            name="export_format",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        formatParam.filter.type = "ValueList"
        formatParam.filter.list = ["PDF", "JPEG"]
        formatParam.value = "PDF"

        fileName = arcpy.Parameter(
            displayName="Output file name (PDF or folder name for JPEGs)",
            name="fileName",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        resolutionParam = arcpy.Parameter(
            displayName="Select resolution",
            name="resolution",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        resolutionParam.filter.type = "ValueList"
        resolutionParam.filter.list = ["High (600 DPI)", "Medium (300 DPI)", "Low (150 DPI)"]
        resolutionParam.value = "Medium (300 DPI)"

        workSpace = arcpy.Parameter(
            displayName="Output folder",
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )

        geoParam = arcpy.Parameter(
            displayName="Include Georeferencing Information",
            name="export_geo",
            datatype="GPBoolean",
            parameterType="Optional",
            direction="Input"
        )
        geoParam.value = True

        return [layoutList, formatParam, fileName, resolutionParam, workSpace, geoParam]

    def updateParameters(self, parameters):
        try:
            aprx = arcpy.mp.ArcGISProject("CURRENT")
            layouts = aprx.listLayouts()
            parameters[0].filter.list = [layout.name for layout in layouts]
        except Exception as e:
            arcpy.AddWarning(f"Could not update layout list: {e}")
        return

    def execute(self, parameters, messages):
        try:
            layout_names = parameters[0].values
            format_type = parameters[1].valueAsText
            file_name = parameters[2].valueAsText
            resolution = parameters[3].valueAsText
            output_folder = parameters[4].valueAsText
            include_geo = parameters[5].value

            aprx = arcpy.mp.ArcGISProject("CURRENT")
            resolution_map = {
                "High (600 DPI)": 600,
                "Medium (300 DPI)": 300,
                "Low (150 DPI)": 150
            }
            dpi = resolution_map.get(resolution, 300)
            arcpy.AddMessage(f"Exporting layouts: {layout_names} to {format_type} in {output_folder} at {dpi} DPI")
            if format_type == "PDF":
                pdf_path = os.path.join(output_folder, file_name if file_name.endswith(".pdf") else file_name + ".pdf")
                pdf_doc = arcpy.mp.PDFDocumentCreate(pdf_path)

                for layout_name in layout_names:
                    layout = aprx.listLayouts(layout_name)[0]
                    temp_pdf = os.path.join(output_folder, layout_name + "_temp.pdf")
                    layout.exportToPDF(temp_pdf, resolution=dpi)
                    pdf_doc.appendPages(temp_pdf)
                    os.remove(temp_pdf)

                pdf_doc.saveAndClose()
                messages.addMessage(f"Exported {len(layout_names)} layouts to {pdf_path}")

            elif format_type == "JPEG":
                jpeg_folder = os.path.join(output_folder, file_name)
                if not os.path.exists(jpeg_folder):
                    os.makedirs(jpeg_folder)

                for layout_name in layout_names:
                    layout = aprx.listLayouts(layout_name)[0]
                    jpeg_path = os.path.join(jpeg_folder, layout_name + ".jpg")
                    layout.exportToJPEG(jpeg_path, resolution=dpi, jpeg_quality=80)

                    if not include_geo:
                        tfw = os.path.splitext(jpeg_path)[0] + ".tfw"
                        aux = jpeg_path + ".aux.xml"
                        for f in [tfw, aux]:
                            if os.path.exists(f):
                                os.remove(f)
                                messages.addMessage(f"Removed georeferencing file: {f}")

                messages.addMessage(f"Exported {len(layout_names)} layouts to folder {jpeg_folder}")

        except Exception as e:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(e)
            msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)

