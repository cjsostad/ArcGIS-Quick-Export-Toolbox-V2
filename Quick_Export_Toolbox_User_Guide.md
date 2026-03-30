# Quick Export Toolbox — User Guide

**Ministry of Water, Land and Resource Stewardship**

---

## What Is the Quick Export Toolbox?

The Quick Export Toolbox is a set of three tools that let you quickly export your map layouts from ArcGIS Pro to **PDF** or **JPEG** files. It considerably speeds up the process of exporting PDF/JPG maps from ArcGIS Pro, which is known to take several minutes per map.

---

## How to Add the Toolbox to Your ArcGIS Pro Project

1. Open your ArcGIS Pro project.
2. In the **Catalog** pane, right-click on **Toolboxes**.
3. Click **Add Toolbox**.
4. Browse to `\\giswhse.env.gov.bc.ca\whse_np\corp\cartographic_resources\resources\tools\Quick_Export_Toolbox.pyt` and select it.
5. Click **OK**. The toolbox will now appear under **Toolboxes** in your Catalog pane.

---

## The Three Tools

The toolbox contains three tools. Choose the one that fits your situation:

| Tool Name | When to Use It |
|---|---|
| **Export Layout From Project With Only One Layout** | Your project has exactly one layout and you want to export it. There is no drop-down menu to select the layout because there is only one. |
| **From Multiple Layouts Export Single Layout** | Your project has multiple layouts, and you want to export just one of them. There will be a drop-down parameter where you can select which layout to export. |
| **Export Multiple Layouts to Single File** | You want to export several different layouts at the same time into a single PDF containing multiple pages. |

---

## Tool 1: Export Layout From Project With Only One Layout

Use this tool when your project contains a single layout. It will automatically detect that layout and export it.

### Parameters

| Parameter | Description |
|---|---|
| **Navigate to the folder where you want to save your file** | Click the folder icon to browse to the folder where the exported file will be saved. **Warning:** If a file with the same name already exists in that folder, it will be overwritten without asking. |
| **File name you want for your output** | Type the name you want for your exported file. The tool auto-fills this with your layout name by default. The correct file extension (`.pdf` or `.jpg`) will be added automatically if you don't include it. |
| **Select vector resolution** | Choose the image quality of your export: |
| | • **High (600 DPI)** — Best quality; larger file size. Good for high-quality print and publication. |
| | • **Medium (300 DPI)** — Recommended default. Suitable for most print and screen uses. |
| | • **Low (150 DPI)** — Smallest file size; lower quality. Suitable for quick drafts or screen-only use. |
| **Select export format** | Choose the output file type: |
| | • **PDF** — Best for sharing, printing, and multi-page documents. |
| | • **JPEG** — Good for inserting into presentations or documents. |
| **Include Georeferencing Information** | When checked (the default), the exported PDF will include georeferencing information. |

### Steps

1. Double-click the tool to open it.
2. Browse to your desired output folder.
3. Enter or confirm the output file name.
4. Select your preferred resolution and format.
5. Check or uncheck the georeferencing option.
6. Click **Run**.

---

## Tool 2: From Multiple Layouts Export Single Layout

Use this tool when your project contains more than one layout and you want to export just one of them.

### Parameters

| Parameter | Description |
|---|---|
| **Select a layout from a list of layouts** | A dropdown list showing all the layouts in your current project. Click the dropdown and select the layout you want to export. |
| **File name you want for your output** | Type the name you want for the exported file. Defaults to `ChangeMe.pdf` — be sure to replace this with a meaningful name. The correct file extension will be added automatically if you don't include it. |
| **Navigate to the folder where you want to save your file** | Click the folder icon to browse to the folder where the exported file will be saved. **Warning:** If a file with the same name already exists, it will be overwritten. |
| **Select vector resolution** | Choose the image quality: **High (600 DPI)**, **Medium (300 DPI)**, or **Low (150 DPI)**. See descriptions in Tool 1 above. |
| **Select export format** | Choose **PDF** or **JPEG**. |
| **Include Georeferencing Information** | When checked (the default), the exported PDF will include georeferencing information. |

### Steps

1. Double-click the tool to open it.
2. Select the layout you want to export from the dropdown.
3. Change the PDF file name to something meaningful.
4. Browse to your desired output folder.
5. Select your preferred resolution and format.
6. Check or uncheck the georeferencing option.
7. Click **Run**.

---

## Tool 3: Export Multiple Layouts to Single File

Use this tool when you want to export two or more layouts to a single PDF that contains multiple pages — one layout per page. (This saves you from having to use Adobe Acrobat later to join the PDFs.)

*Note: JPEGs will be exported as individual files in a folder with the name provided in the output file name parameter.*

- If you choose **PDF**, all selected layouts are merged into a **single multi-page PDF** file.
- If you choose **JPEG**, each selected layout is exported as an **individual JPEG** inside a new folder.

### Parameters

| Parameter | Description |
|---|---|
| **Select layouts to export** | A drop-down menu where you can choose the layouts you want to export. You can select as many as you need; after choosing the first layout you will be prompted to choose subsequent layouts. |
| **Select export format** | Choose PDF or JPEG: |
| | • **PDF** — All layouts are combined into one multi-page PDF. |
| | • **JPEG** — Each layout is saved as a separate `.jpg` file inside a folder. |
| **Output file name** | If exporting to PDF, this is the file name for the combined PDF (e.g., `AllMaps`). If exporting to JPEG, this is the name of the folder that will be created to hold the individual JPEG files. The `.pdf` extension is added automatically for PDF exports. |
| **Select resolution** | Choose the image quality: **High (600 DPI)**, **Medium (300 DPI)**, or **Low (150 DPI)**. |
| **Output folder** | Browse to the folder where the PDF file or JPEG folder will be saved. |
| **Include Georeferencing Information** | Check to include georeferencing information. |

### Steps

1. Double-click the tool to open it.
2. Select one or more layouts from the list.
3. Choose your export format (PDF or JPEG).
4. Enter an output file or folder name.
5. Select your preferred resolution.
6. Browse to the output folder location.
7. Check or uncheck the georeferencing option.
8. Click **Run**.

---

## Frequently Asked Questions

### What does DPI mean?

DPI stands for **Dots Per Inch**. It controls how detailed and sharp your exported image will be. Higher DPI means better quality but a larger file size. For most uses, **300 DPI (Medium)** is a good balance.

### My file was overwritten — can I get it back?

All three tools have **overwrite enabled by default**. If you export to a file name that already exists in the same folder, the old file will be replaced. Always double-check your file name and output folder before clicking Run. There is no undo for overwritten files.

### The tool doesn't show any layouts in the dropdown — what's wrong?

The tool reads layouts from your **current ArcGIS Pro project**. If no layouts appear, it means your project does not have any layouts set up yet. To create a layout, go to the **Insert** tab on the ribbon and click **New Layout**.

### Can I use this tool with map frames instead of layouts?

No. These tools export **layouts**, not individual map frames. A layout is the page view that contains your map frame along with titles, legends, scale bars, and other map elements. Make sure your map content is placed inside a layout before using this toolbox.
