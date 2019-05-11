function New-DocWordNumbering {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $WordObject
    )

    $TOC = Add-WordTocItem -WordDocument $WordDocument -Text $WordObject.Text -ListLevel $WordObject.Level -ListItemType $WordObject.Type -HeadingType $WordObject.Heading

    if ($WordObject.Content) {
        #$Content = Invoke-Command -ScriptBlock $WordObject.Content
        New-WordProcessing -Content $WordObject.Content -WordDocument $WordDocument
    }
}