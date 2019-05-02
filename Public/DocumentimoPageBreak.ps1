function DocumentimoPageBreak {
    [alias('New-DocumentimoPageBreak', 'DocPageBreak')]

    param(
        [int] $Count = 1
    )
    $i = 0
    While ($i -lt $Count) {
        Write-Verbose "New-WordBlockPageBreak - PageBreak $i"
        Add-WordPageBreak -Supress $True -WordDocument $Script:Parameters.WordDocument
        $i++
    }
}