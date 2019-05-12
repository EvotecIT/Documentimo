﻿function New-DocWordPageBreak {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )

    $i = 0
    While ($i -lt $Parameters.Count) {
        Write-Verbose "New-WordBlockPageBreak - PageBreak $i"
        Add-WordPageBreak -Supress $True -WordDocument $WordDocument
        $i++
    }
}