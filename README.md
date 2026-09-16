# excel-image-assistant
Excel VBA add-in for automatically inserting images beside worksheet codes from local folders or URLs, with batch insertion, comments, image sizing, and configurable left/right placement.

Excel Image Assistant - Mohamed

A lightweight Microsoft Excel add-in for inserting and managing images directly from worksheet codes, file names, local folders, URLs, and comments.

Current stable version: v1.7
Author / Info: Mohamed Abd Elnasser - IT Manager
Platform: Microsoft Excel Desktop for Windows
Activation / License Manager: None
QLM / external activation DLLs: Not required

Overview

Excel Image Assistant - Mohamed is an Excel VBA add-in designed to make image-heavy Excel workbooks faster and easier to manage.

The add-in can use a worksheet value such as:

2611029

as an image reference.

If the image folder contains:

2611029.jpg

or another supported image extension, Excel Image Assistant can find the matching image automatically and place it beside the selected code.

The original code cell remains untouched. The image can be inserted either to the left or right of the reference cell according to your settings.

Main Features

Pictures

Insert one image using the selected cell as the reference.

Insert images for multiple selected cells.

Insert images for a selected code column.

Match worksheet codes to image file names automatically.

Delete inserted pictures.

Fit and center pictures inside destination cells.

Replace an existing EIA image in the same destination instead of stacking duplicate pictures.

Keep the original code/reference cell visible.

Choose whether the image is inserted to the left or right of the code.

Comments / Notes

Insert an image into a cell comment/note.

Insert comment images in bulk.

Delete image comments.

Configure comment image height.

URL Images

Insert an image from a URL.

Insert multiple URL images.

Delete inserted URL pictures.

Uses Windows COM components rather than architecture-specific API declarations.

File Names

Insert a selected image file name into Excel.

Insert multiple image file names.

Picture Compression

Access Excel's picture compression command directly from the add-in.

Settings

The add-in includes configurable options for:

Image folder path.

Row offset.

Column offset.

Insert into cells or comments.

Comment image height.

Picture placement behavior:

Move and size with cells.

Move but don't size with cells.

Don't move or size with cells.

Inserting direction:

Down.

Right.

Insert image beside code:

Left.

Right.

Example Workflow

Assume your worksheet contains:

Code

Image

2611029



2611030



2611031



And your image folder contains:

2611029.jpg
2611030.png
2611031.jpeg

If Insert image beside code = Right

The result is:

Code

Image

2611029

[2611029 image]

2611030

[2611030 image]

2611031

[2611031 image]

If Insert image beside code = Left

The result is:

Image

Code

[2611029 image]

2611029

[2611030 image]

2611030

[2611031 image]

2611031

The picture is not inserted over the code cell.

Supported Image Matching

When a selected cell contains a code such as:

2611029

the add-in searches the configured folder for a matching image.

Common supported extensions include:

.jpg
.jpeg
.png
.bmp
.gif
.ico
.tif
.tiff
.webp

For best results, use image file names that exactly match the worksheet code.

Example:

Excel cell: 2611029
Image file: 2611029.jpg

Ribbon

After installation, Excel displays a new tab:

Excel Image Assistant

The Ribbon contains the following groups:

Pictures

Insert

Insert All

Delete

Comments

Insert

Insert All

Delete

URL

Insert

Insert All

Delete

Compress

Compress Pictures

File Names

Insert

Insert All

Excel Image Assistant

EIA

Settings

Info

Version v1.7 uses unique internal Ribbon IDs to reduce conflicts with older builds or other add-ins.

A fallback Excel Image Assistant menu is also created under Excel's Add-ins interface when macros are available but the custom Ribbon is not rendered.

Requirements

Windows 10 or Windows 11 recommended.

Microsoft Excel Desktop.

VBA/macros enabled.

Access to the VBA project object model is required only when building the add-in from source.

PowerShell is required for the included build/install scripts.

This project is intended for the Windows desktop version of Excel. Excel for the web does not run VBA add-ins.

Installation

Recommended Installation

Download or clone this repository.

Close all Excel windows.

Open the project folder.

Run:

Build Add-in.cmd

Wait until the build reports:

XLAM package + Ribbon relationship validation: PASS
DONE - v1.7 validated build

Run:

Install Add-in.cmd

Close Excel completely if it was opened in the background.

Open Excel again.

Look for the Excel Image Assistant tab.

Building From Source

The repository contains a PowerShell-based builder that creates the final .xlam add-in.

Before running the builder, Excel may require this setting:

Open Excel.

Go to File > Options > Trust Center.

Click Trust Center Settings.

Open Macro Settings.

Enable:

Trust access to the VBA project object model

Close Excel.

Run Build Add-in.cmd.

You can disable that setting again after the build if you do not normally need programmatic access to VBA projects.

Using Pictures > Insert

Open Excel Image Assistant > Settings.

Set the folder containing your images.

Choose whether images should be inserted on the Left or Right side of the code.

Select a cell containing an image code.

Click:

Pictures > Insert

The add-in searches the configured folder and inserts the matching image in the adjacent destination cell.

Using Pictures > Insert All

Set your image folder in Settings.

Select the cells containing the image codes.

For example:

2611029
2611030
2611031
2611032

Click:

Pictures > Insert All

The add-in processes the selected reference cells row by row.

If a whole column is selected, the add-in limits processing to the used worksheet area instead of scanning all 1,048,576 Excel rows.

Blank cells are skipped.

If one image cannot be found or inserted, the remaining references can continue processing.

Image Placement

The reference cell is used only to identify the image.

For example:

Selected code cell = D25
Value = 2611029

With Right selected:

D25 = 2611029
E25 = Image

With Left selected:

C25 = Image
D25 = 2611029

The image is fitted and centered inside the destination cell or merged-cell area.

Picture Placement Properties

You can select how the inserted picture behaves when worksheet cells change:

Move and size with cells

The picture moves and resizes with its cells.

Move but don't size with cells

The picture follows the cells but keeps its own size.

Don't move or size with cells

The picture stays independent of worksheet cell sizing.

Comments / Notes Mode

Excel Image Assistant can also place pictures inside legacy Excel comments/notes.

Available options include:

Single comment image insertion.

Bulk comment image insertion.

Comment image deletion.

Configurable image height.

Note that newer versions of Excel distinguish between threaded Comments and legacy Notes. VBA's classic comment APIs normally operate on the legacy note-style object.

URL Images

The URL tools can retrieve an image from a web address and insert it into Excel.

The current implementation avoids architecture-specific URLDownloadToFile VBA declarations and uses late-bound Windows COM components instead, which improves compatibility across common Office installations.

Internet access and the target server's permissions are still required.

Project Structure

EIA_Mohamed_Clone_v1.7/
│
├── Build Add-in.cmd
├── Build-Addin.ps1
├── Install Add-in.cmd
├── Install-Addin.ps1
├── EIA_Core.bas
├── customUI.xml
├── customUI14.xml
└── README.txt

EIA_Core.bas

Main VBA functionality, including:

Ribbon callbacks.

Image insertion.

Batch image insertion.

Image matching.

Image destination logic.

Comments/notes.

URL image retrieval.

File-name tools.

Picture deletion.

Settings storage.

Fallback Excel Add-ins menu.

customUI.xml

Ribbon definition for compatible Office versions.

customUI14.xml

Ribbon definition using the newer Office custom UI schema.

Build-Addin.ps1

Creates the .xlam add-in, builds the VBA forms, injects the Ribbon XML, and validates the resulting package.

Install-Addin.ps1

Installs the generated add-in into the Excel add-ins environment and helps avoid old cached builds.

Troubleshooting

The Excel Image Assistant tab does not appear

Try the following:

Close all Excel processes.

Make sure the new build was installed rather than an older .xlam.

Run Install Add-in.cmd again.

Open Excel.

Check File > Options > Add-ins.

At the bottom, choose Excel Add-ins and click Go.

Confirm the add-in is enabled.

Version v1.7 uses unique internal Ribbon IDs to reduce Ribbon cache conflicts.

The add-in also includes a fallback menu under Excel's Add-ins tab when available.

Build fails because VBA project access is blocked

Enable:

File
> Options
> Trust Center
> Trust Center Settings
> Macro Settings
> Trust access to the VBA project object model

Then close Excel completely and run the builder again.

Images are not found

Check:

The configured image folder.

The value inside the selected Excel cell.

The image file name.

The image extension.

Leading/trailing spaces in codes or filenames.

Whether the file is actually inside the configured folder.

Recommended convention:

Cell value: 2611029
File name: 2611029.jpg

The image is inserted in the wrong location

Open Settings and check:

Insert image beside code

Choose:

Left

or:

Right

The code cell itself is not intended to be the image destination.

Excel blocks macros

Depending on your organization's security settings, files downloaded from the internet may be blocked by Windows or Microsoft Office.

Use only files you trust.

If necessary, check the file's Windows properties or your organization's Excel Trust Center policy.

Security

The add-in does not require:

QLM.

Activation servers.

License keys.

IsLicense50.dll.

QlmLicenseLib.dll.

QlmLicenseLib.tlb.

The add-in does use VBA and local Windows/Excel automation, so users should inspect the source and only run code they trust.

For organizational deployments, review the scripts and VBA code according to your IT security policy before distributing the add-in.

Privacy

Normal local image insertion is performed on the user's computer.

The add-in does not require an account or activation service.

When using the URL feature, Excel/Windows connects to the URL supplied by the user in order to retrieve the image.

Version

v1.7

Current stable base build.

Highlights:

Stable custom Ribbon installation.

Unique Ribbon IDs to reduce conflicts with previous builds.

Automatic image matching from worksheet code values.

Left/right image placement beside the reference code.

Prevents pictures from being intentionally placed over the code cell.

Batch insertion for selected references.

Whole-column selection limited to the worksheet's used range.

Existing EIA image replacement in the destination cell.

Local folder image support.

Comment/note image support.

URL image tools.

File-name tools.

Picture compression shortcut.

Excel Add-ins fallback menu.

No QLM or external activation/license DLL dependencies.

Author

Mohamed Abd Elnasser
IT Manager

The add-in's Info screen and workbook metadata identify:

Mohamed Abd Elnasser - IT Manager

Repository Description

Suggested GitHub repository description:

Excel VBA add-in for automatically inserting images beside worksheet codes from local folders or URLs, with batch insertion, comments, image sizing, and configurable left/right placement.

Suggested GitHub Topics

excel
excel-vba
vba
excel-addin
xlam
images
image-insertion
automation
microsoft-excel
office
powershell
productivity

Suggested Repository Name

excel-image-assistant

Alternative:

excel-image-assistant-mohamed

Suggested Release Title

Excel Image Assistant v1.7

Suggested release summary:

Stable Windows Excel add-in for inserting images automatically from worksheet codes. Includes single/batch image insertion, configurable left/right placement, comments, URLs, file-name tools, picture compression, custom Ribbon integration, and no activation system.

Contributing

Issues and pull requests are welcome.

When reporting a problem, please include:

Excel version.

Office 32-bit or 64-bit.

Windows version.

The action being performed.

A screenshot of the error.

The exact build/install console output when relevant.

Do not include confidential workbooks, private URLs, credentials, or sensitive company data in public issues.

Disclaimer

This project is provided as an independent Excel automation utility.

Microsoft, Excel, Windows, and Office are trademarks of Microsoft Corporation. This project is not affiliated with or endorsed by Microsoft.

Before using the add-in on important workbooks, keep a backup and test the workflow on a copy of your file.

License

No software activation or license-key system is built into the add-in.

Important for GitHub: the absence of an activation system is different from an open-source copyright license. If you want other people to legally use, modify, and redistribute the source code, add a repository license such as the MIT License, Apache-2.0, or another license appropriate for your project.
