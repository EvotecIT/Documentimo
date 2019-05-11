function New-DocWordTable {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $WordObject
    )


    if ($WordObject.OverWriteTitle) {
        #[nullable[int]] $TableMaximumColumns = 5,
        # [nullable[bool]] $TableTitleMerge,
        #[string] $TableTitleText,
        [Alignment] $TitleAlignment = [Alignment]::center
        [System.Drawing.Color] $TitleColor = [System.Drawing.Color]::Black


        $Table = Add-WordTable -WordDocument $WordDocument -Supress $false -DataTable $WordObject.DataTable -Design $WordObject.Design -AutoFit $WordObject.AutoFit #-DoNotAddTitle
        $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll  # -ColumnNrStart 0 -ColumnNrEnd 1
        $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
        $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $WordObject.OverwriteTitle -Alignment $TitleAlignment -Color $TitleColor
    } else {
        $Table = Add-WordTable -WordDocument $WordDocument -Supress $true -DataTable $WordObject.DataTable -Design $WordObject.Design -AutoFit $WordObject.AutoFit
    }
}