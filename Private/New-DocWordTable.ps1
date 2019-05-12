function New-DocWordTable {
    [CmdletBinding()]
    param(
        [Xceed.Words.NET.Container] $WordDocument,
        [PSCustomObject] $Parameters
    )


    if ($Parameters.OverWriteTitle) {
        #[nullable[int]] $TableMaximumColumns = 5,
        # [nullable[bool]] $TableTitleMerge,
        #[string] $TableTitleText,
        [Alignment] $TitleAlignment = [Alignment]::center
        [System.Drawing.Color] $TitleColor = [System.Drawing.Color]::Black


        $Table = Add-WordTable -WordDocument $WordDocument -Supress $false -DataTable $Parameters.DataTable -Design $Parameters.Design -AutoFit $Parameters.AutoFit #-DoNotAddTitle
        $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll  # -ColumnNrStart 0 -ColumnNrEnd 1
        $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
        $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $Parameters.OverwriteTitle -Alignment $TitleAlignment -Color $TitleColor
    } else {
        $Table = Add-WordTable -WordDocument $WordDocument -Supress $true -DataTable $Parameters.DataTable -Design $Parameters.Design -AutoFit $Parameters.AutoFit
    }

}

foreach ($Domain in $ADForest.FoundDomains.Keys | Where-Object { $_ -ne 'ad.evotec.pl' }) {

    Write-Color $ADForest.FoundDomains.$Domain
}