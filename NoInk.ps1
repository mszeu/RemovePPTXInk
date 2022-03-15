<#
    .SYNOPSIS
     Removes the Ink Annotations from the file specified.
    .DESCRIPTION
    Removes the Ink Annotations from the file specified. If no files are specified it acts on all the the 
    PowerPoint Presentation found in the current directory.
    If the file is marked Final it removes the attribute
    .PARAMETER FileName
     The PowerPoint presentation from where you'd like to remove the ink annotations
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
    .NOTES
      Author: Marco S. Zuppone - msz@msz.eu - https://msz.eu
      Version: 0.2
      License: AGPL 3.0 - Plese abide to the Aferro AGPL 3.0 license rules! It's free but give credits to the author :-)
     
#>
param (
    
    [parameter(Position = 0)][string]$FileName
)   
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
        $presentation = $officeObj.Presentations.Open($presentation_file.FullName)
        $slides = $presentation.Slides
        foreach ($slide in $slides) {
            foreach ($shape in $slide.Shapes) {
                if ($shape.Type -eq 23) {
                    Write-Host $shape #TODO: Explore what are the properties of the SHAPE object
					Write-Host $slide.SlideIndex " Slide" $slide.SlideNumber
                    $shape.Delete()
                }
            }
        }
        $presentation.Final = $false
        $presentation.Save()
        $presentation.Close()
        Write-Host $presentation_file.FullName "has been processed"
    }
}
else { Write-Host "The file specified was not found" }
$officeObj.Quit()
$officeObj = $null


