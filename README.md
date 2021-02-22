# NoInk

Removes Ink Annotations from PowerPoint presentations

# DESCRIPTION

Removes the Ink Annotations from the file specified. 

If no files are specified it acts on all the the PowerPoint presentations found in the current directory.

If the file is marked **Final** it removes the attribute.

## Usage

From the **Powershell** command line:

**.\NoInk.ps1** 
- It removes the Ink Annotation from all the .pptx files found in the current directory

**.\NoInk.ps1 myPresentation.pptx**
- It removes the Ink Annotation from the specified file

# NOTES
The script has been tested on Windows 7 x64 SP1 and on Windows 10 with Office Plus 2013 and Office 2019

The PowerShell version was: 5.1
# DISCLAIMER & LICENSE
The script is given **AS IS** and it is under the **AGPL Aferro license 3.0**.