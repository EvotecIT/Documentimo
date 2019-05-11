function New-WordProcessing {
    [CmdletBinding()]
    param(
        [ScriptBlock] $Content,
        [Xceed.Words.NET.Container] $WordDocument
    )
    if ($null -ne $Content) {
        $Array = Invoke-Command -ScriptBlock $Content
        foreach ($WordObject in $Array) {
            if ($WordObject.ObjectType -eq 'List') {
                New-DocWordList -WordDocument $WordDocument -WordObject $WordObject
            }
            if ($WordObject.ObjectType -eq 'Table') {
                New-DocWordTable -WordDocument $WordDocument -WordObject $WordObject
            }
            if ($WordObject.ObjectType -eq 'TOC') {
                New-DocWordTOC -WordDocument $WordDocument -WordObject $WordObject
            }
            if ($WordObject.ObjectType -eq 'Text') {
                New-DocWordText -WordDocument $WordDocument -WordObject $WordObject
            }
            if ($WordObject.ObjectType -eq 'TocItem') {
                New-DocWordNumbering -WordDocument $WordDocument -WordObject $WordObject
            }
            if ($WordObject.ObjectType -eq 'PageBreak') {
                New-DocWordPageBreak -WordDocument $WordDocument -WordObject $WordObject
            }
            if ($WordObject.ObjectType -eq 'ChartPie') {
                New-DocWordChart -WordDocument $WordDocument -WordObject $WordObject
            }
        }
    }
}