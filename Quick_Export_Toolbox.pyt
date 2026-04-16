# Author: Chris Sostad and Marina Cunningham WLRS/Wildlife and Fisheries
# Date Created: February 21, 2024
# Date Updated: April 16th, 2026

# Update Log:
# April 16th, 2026 - Set default output folder to current project folder,
#   added georeferencing cleanup option, and added error handling for
#   georeferencing file deletion.

# Description:
#   ArcGIS Pro Python Toolbox with three tools for exporting layouts to PDF or
#   JPEG. Supports single-layout projects, selected layout export from
#   multi-layout projects, and multi-layout batch export. Includes resolution
#   presets (150-600 DPI) and optional georeferencing cleanup.

"""
ArcGIS Pro Python Toolbox for quick layout export to PDF or JPEG.
"""

import arcpy
import traceback
import sys
import os

# Global Functions
# -----------------------------------------------------------------------------------------------


def export_layout(layout, output_path, file_name, resolution, format_type, include_geo, messages):
    """
    Export a layout to PDF or JPEG and optionally remove georeferencing files.

    Args:
        layout: ArcGIS Pro layout object to export
        output_path (str): Target folder path
        file_name (str): Output file name including extension
        resolution (str): Resolution label (e.g., "Medium (300 DPI)")
        format_type (str): Export format, either "PDF" or "JPEG"
        include_geo (bool): Keep georeferencing sidecar files if True
        messages: ArcPy message object used by toolbox execution
    """
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

# Toolbox Definition
# -----------------------------------------------------------------------------------------------


class Toolbox(object):
    """ArcGIS toolbox container for quick layout export tools."""
    def __init__(self):
        self.label = "Toolbox"
        self.alias = ""
        self.tools = [
            ExportSingleLayout,
            FromMultipleExportSingleLayout,
            ExportMultipleLayoutsToSingleFile
        ]

# ExportSingleLayout Tool
# -----------------------------------------------------------------------------------------------


class ExportSingleLayout(object):
    """Export the only layout in the current project to PDF or JPEG."""
    def __init__(self):
        self.label = "Export Layout From Project With Only One Layout"
        self.description = "Export a selected layout from your project to a PDF or JPEG."
        self.canRunInBackground = False

    def getParameterInfo(self):
        workSpace = arcpy.Parameter(
            displayName=(
                "Navigate to the folder where you want to save your file "
                "(Warning: Overwrite set to true!)"
            ),
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )
        # Set default to current project folder
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        workSpace.value = os.path.dirname(aprx.filePath)

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
        geoParam.value = True  # Checked by default

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

            export_layout(
                layout, workSpace_path, file_name, resolution_level,
                format_type, include_geo, messages
            )

        except arcpy.ExecuteError:
            arcpy.AddError(arcpy.GetMessages(2))
        except Exception as e:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = (
                f"PYTHON ERRORS:\nTraceback info:\n{tbinfo}"
                f"\nError Info:\n{e}"
            )
            msgs = f"ArcPy ERRORS:\n{arcpy.GetMessages(2)}\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)

# FromMultipleExportSingleLayout Tool
# -----------------------------------------------------------------------------------------------


class FromMultipleExportSingleLayout(object):
    """Export one selected layout from a project with multiple layouts."""
    def __init__(self):
        self.label = "From Multiple Layouts Export Single Layout"
        self.description = (
            "Choose from multiple layouts and export a selected layout "
            "to PDF or JPEG."
        )
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
            displayName=(
                "Navigate to the folder where you want to save your file "
                "(Warning: Overwrite set to true!)"
            ),
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )
        workSpace.filter.list = ["Local Database", "File System"]
        # Set default to current project folder
        workSpace.value = os.path.dirname(aprx.filePath)

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

    def updateParameters(self, parameters):
        """Auto-populate filename when layout is selected."""
        if parameters[0].altered and parameters[0].value:
            # Get selected layout name
            layout_name = parameters[0].valueAsText
            # Get selected format (default to PDF if not set)
            format_type = parameters[4].valueAsText if parameters[4].value else "PDF"
            extension = ".pdf" if format_type == "PDF" else ".jpg"
            # Update filename parameter
            parameters[1].value = f"{layout_name}{extension}"
        return

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

            export_layout(
                layout, workSpace_path, file_name, resolution_level,
                format_type, include_geo, messages
            )

            arcpy.AddMessage(
                f"Layout '{selected_layout_name}' exported to {workSpace_path}"
            )

        except arcpy.ExecuteError:
            arcpy.AddError(arcpy.GetMessages(2))
        except Exception as e:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = f"PYTHON ERRORS:\nTraceback info:\n{tbinfo}\nError Info:\n{e}"
            msgs = f"ArcPy ERRORS:\n{arcpy.GetMessages(2)}\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)

# ExportMultipleLayoutsToSingleFile Tool
# -----------------------------------------------------------------------------------------------


class ExportMultipleLayoutsToSingleFile(object):
    """Export multiple layouts to one PDF or to individual JPEG files."""
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
        # Set default to current project folder
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        workSpace.value = os.path.dirname(aprx.filePath)

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
            arcpy.AddMessage(
                f"Exporting layouts: {layout_names} to {format_type} in "
                f"{output_folder} at {dpi} DPI"
            )
            if format_type == "PDF":
                pdf_file = (
                    file_name if file_name.endswith(".pdf")
                    else file_name + ".pdf"
                )
                pdf_path = os.path.join(output_folder, pdf_file)
                pdf_doc = arcpy.mp.PDFDocumentCreate(pdf_path)

                for layout_name in layout_names:
                    layout = aprx.listLayouts(layout_name)[0]
                    temp_pdf = os.path.join(output_folder, layout_name + "_temp.pdf")
                    layout.exportToPDF(temp_pdf, resolution=dpi)
                    pdf_doc.appendPages(temp_pdf)
                    os.remove(temp_pdf)

                pdf_doc.saveAndClose()
                messages.addMessage(
                    f"Exported {len(layout_names)} layouts to {pdf_path}"
                )

            elif format_type == "JPEG":
                jpeg_folder = os.path.join(output_folder, file_name)
                if not os.path.exists(jpeg_folder):
                    os.makedirs(jpeg_folder)

                for layout_name in layout_names:
                    layout = aprx.listLayouts(layout_name)[0]
                    jpeg_path = os.path.join(jpeg_folder, f"{layout_name}.jpg")
                    layout.exportToJPEG(jpeg_path, resolution=dpi, jpeg_quality=80)

                    if not include_geo:
                        tfw = f"{os.path.splitext(jpeg_path)[0]}.tfw"
                        aux = f"{jpeg_path}.aux.xml"
                        for f in [tfw, aux]:
                            if os.path.exists(f):
                                os.remove(f)
                                messages.addMessage(
                                    f"Removed georeferencing file: {f}"
                                )

                messages.addMessage(
                    f"Exported {len(layout_names)} layouts to folder {jpeg_folder}"
                )

        except Exception as e:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = f"PYTHON ERRORS:\nTraceback info:\n{tbinfo}\nError Info:\n{e}"
            msgs = f"ArcPy ERRORS:\n{arcpy.GetMessages(2)}\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)

