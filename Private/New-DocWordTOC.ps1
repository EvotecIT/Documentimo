function New-DocWordTOC {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $WordObject
    )

    Add-WordToc -WordDocument $WordDocument -Title $WordObject.Title -Switches $WordObject.Switches -RightTabPos $WordObject.RightTabPos -Supress $True
}