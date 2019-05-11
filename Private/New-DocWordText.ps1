function New-DocWordText {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $WordObject
    )


    if ($WordObject.Text) {
        Add-WordText -WordDocument $WordDocument -Text $WordObject.Text -Color $WordObject.Color -Supress $true
    }
    if ($WordObject.LineBreak) {
        Add-WordParagraph -WordDocument $WordDocument -Supress $True
    }
}