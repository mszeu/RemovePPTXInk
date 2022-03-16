<#
    .SYNOPSIS
     Removes the Ink Annotations from the file specified.
    .DESCRIPTION
     Removes the Ink Annotations from the file specified. If no files are specified it acts on all the the 
     PowerPoint Presentation found in the current directory.
     If the file is marked Final it removes the attribute
    .PARAMETER FileName
     The PowerPoint presentation from where you'd like to remove the ink annotations
    .PARAMETER ShowAll
     If the parameter is specified all the objests found in every slides are enumerated but only the objects of type msoInk and msoInkComment are removed
    .PARAMETER DryRun
    If the parameter is specified the objects of type msoInk and msoInkComment are not deleted
     .INPUTS
     None
    .OUTPUTS
     None
    .EXAMPLE
     .\NoInk.ps1
     Removes the ink annotations from all the slide decks (.pptx files) found in the current directory
    .EXAMPLE
     .\NoInk.ps1 myPresentation.pptx
     Removes the ink annotations from the file myPresentation.pptx
    .EXAMPLE
     .\NoInk.ps1 C:\Users\mzuppone\Desktop
     Removes the ink annotations from all the slide decks (.pptx files) found in C:\Users\mzuppone\Desktop
    .EXAMPLE
     .\NoInk.ps1 C:\Users\mzuppone\Desktop -DryRun
     All the slide decks (.pptx files) found in C:\Users\mzuppone\Desktop are processed but no shapers are deleted from them
    .EXAMPLE
     .\NoInk.ps1 myPresentation.pptx -ShowAll
     Removes the ink annotations from the file myPresentation.pptx and shows all the shapes type found in every slide
    .NOTES
      Author: Marco S. Zuppone - msz@msz.eu - https://msz.eu
      Version: 0.2
      License: AGPL 3.0 - Plese abide to the Aferro AGPL 3.0 license rules! It's free but give credits to the author :-)
     
#>
param (
    
    [parameter(Position = 0)][string]$FileName,
    [switch]$ShowAll,
    [switch]$DryRun
) 
enum MsoShapeType {
    # For maximum compatibility with all the Microsoft Office installation I defined here the MsoShapeType
    # avoiding to import it form the COM objects of Office
    mso3DModel = 30	        #3D model
    msoAutoShape = 1        #AutoShape
    msoCallout = 2	        #Callout
    msoCanvas = 20	        #Canvas
    msoChart = 3	        #Chart
    msoComment = 4	        #Comment
    msoContentApp = 27	    #Content Office Add-in
    msoDiagram = 21	        #Diagram
    msoEmbeddedOLEObject = 7	#Embedded OLE object
    msoFormControl = 8	    #Form control
    msoFreeform = 5	        #Freeform
    msoGraphic = 28	        #Graphic
    msoGroup = 6	        #Group
    msoIgxGraphic = 24	    #SmartArt graphic
    msoInk = 22	            #Ink
    msoInkComment = 23	    #Ink comment
    msoLine = 9	            #Line
    msoLinked3DModel = 31	#Linked 3D model
    msoLinkedGraphic = 29	#Linked graphic
    msoLinkedOLEObject = 10	#Linked OLE object
    msoLinkedPicture = 11	#Linked picture
    msoMedia = 16	        #Media
    msoOLEControlObject = 12	#OLE control object
    msoPicture = 13	        #Picture
    msoPlaceholder = 14	    #Placeholder
    msoScriptAnchor = 18	#Script anchor
    msoShapeTypeMixed = -2	#Mixed shape type
    msoSlicer = 25	        #Slicer
    msoTable = 19	        #Table
    msoTextBox = 17	        #Text box
    msoTextEffect = 15	    #Text effect
    msoWebVideo = 26	    #Web video
} 
try {
    $officeObj = New-Object -ComObject PowerPoint.Application
}
catch { 
    Write-Host "It was not possible to instantiate the PowerPoint COM object."
    Write-Host "Terminating"
    Write-Error "Error"
    Exit -1
}
$presentation_in_dir = Get-ChildItem -Filter '*.pptx' -File $FileName
if ($null -ne $presentation_in_dir ) {
    foreach ($presentation_file in $presentation_in_dir ) {
        Write-Host "I'm starting to work on presentation" $presentation_file.FullName
        $presentation = $officeObj.Presentations.Open($presentation_file.FullName)
        $slides = $presentation.Slides
        foreach ($slide in $slides) {
            foreach ($shape in $slide.Shapes) {
                if ($ShowAll) {
                    Write-Host "Shape detected of type" $shape.Type ([MsoShapeType].GetEnumName($shape.Type)) "at SlideIndex" $slide.SlideIndex "Slide" $slide.SlideNumber
                    if (($shape.Type -eq 23) -or ($shape.Type -eq 22)) {
                        if (-not $DryRun) {
                            $shape.Delete()
                            Write-Host "Shape deleted"
                        }
                    }
                }
                elseif (($shape.Type -eq 23) -or ($shape.Type -eq 22)) {
                    #The shape type is the MsoShapeType enumeration, the documentation can be found at https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype
                    Write-Host "Shape detected of type" $shape.Type ([MsoShapeType].GetEnumName($shape.Type)) "at SlideIndex" $slide.SlideIndex "Slide" $slide.SlideNumber
                    if (-not $DryRun) {
                        $shape.Delete()
                        Write-Host "Shape deleted"
                    } 
                }
            }
        }
        $presentation.Final = $false
        $presentation.Save()
        $presentation.Close()
        Write-Host $presentation_file.FullName "has been processed"
        Write-Host ""
    }
}
else { Write-Host "The file specified was not found" }
$officeObj.Quit()
$officeObj = $null
