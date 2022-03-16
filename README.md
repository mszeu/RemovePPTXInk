# NoInk

Removes Ink Annotations from PowerPoint presentations

# DESCRIPTION

Removes the Ink Annotations from the file specified.

If no files are specified it acts on all the the PowerPoint presentations found in the current directory.

If the file is marked **Final** it removes the attribute.

## Usage

From the **PowerShell** command line:

**.\NoInk.ps1** 
- It removes the Ink Annotations from all the .pptx files found in the current directory

**.\NoInk.ps1 myPresentation.pptx**
- It removes the Ink Annotations from the specified file

**.\NoInk.ps1 C:\Users\mzuppone\Desktop**
- Removes the ink annotations from all the slide decks (.pptx files) found in **C:\Users\mzuppone\Desktop**

### Optional Parameters

**-ShowAll** If specified all the objests found in every slides are enumerated but only the objects of type msoInk and msoInkComment are removed

**-DryRun** If the parameter is specified the objects of type msoInk and msoInkComment are not deleted

# NOTES
The script has been tested on **Windows 7 x64 SP1** and on **Windows 10** with **Office Plus 2013** and **Office 2019**

It was tested with **Windows PowerShell 5.1**
# DISCLAIMER & LICENSE
The script is given **AS IS** and it is under the **AGPL Aferro license 3.0**.