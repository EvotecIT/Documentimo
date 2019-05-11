function New-DocWordPageBreak {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $WordObject
    )

    $i = 0
    While ($i -lt $WordObject.Count) {
        Write-Verbose "New-WordBlockPageBreak - PageBreak $i"
        Add-WordPageBreak -Supress $True -WordDocument $WordDocument
        $i++
    }
}