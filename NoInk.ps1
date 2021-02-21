<#
    .SYNOPSIS
     Removes the Ink Annotations from the file specified.
    .DESCRIPTION
    Removes the Ink Annotations from the file specified. If no files are specified it acts on all the the 
    PowerPoint Presentation found in the current directory.
    If the file is marked Final it removes the attribute
    .PARAMETER FileName
     The PowerPoint presentation from where you'd like to remove the ink annotations
    .EXAMPLE
     NoInk.ps1
     NoInk.ps1 myPresentation.pptx
    .NOTES
      Author: Marco S. Zuppone - msz@msz.eu - https://msz.eu
      Version: 0.1
      License: AGPL 3.0 - Plese abide to the Aferro AGPL 3.0 license rules! It's free but give credits to the author :-)
      
#>
param (
    
    [parameter(Position = 0)][string]$FileName
)   
$officeObj = New-Object -ComObject PowerPoint.Application
$presentation_in_dir = Get-ChildItem -Filter '*.pptx' -File $FileName
if ($null -ne $presentation_in_dir ) {
    foreach ($presentation_file in $presentation_in_dir ) {
        $presentation = $officeObj.Presentations.Open($presentation_file.FullName)
        $slides = $presentation.Slides
        foreach ($slide in $slides) {
            foreach ($shape in $slide.Shapes) {
                if ($shape.Type -eq 23) {
                    $shape.Delete()
                }
            }
        }
        $presentation.Final=$false
        $presentation.Save()
        $presentation.Close()
        Write-Host $presentation_file.FullName "has been processed"
    }
}
else { Write-Host "The file specified was not found" }
$officeObj = $null


