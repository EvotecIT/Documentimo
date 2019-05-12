﻿function New-DocWordNumbering {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )

    $TOC = Add-WordTocItem -WordDocument $WordDocument -Text $Parameters.Text -ListLevel $Parameters.Level -ListItemType $Parameters.Type -HeadingType $Parameters.Heading

    if ($Parameters.Content) {
        $Content = Invoke-Command -ScriptBlock $Parameters.Content
        New-WordProcessing -Content $Content -WordDocument $WordDocument
    }
}