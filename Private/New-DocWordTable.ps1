function New-DocWordTable {
    [CmdletBinding()]
    param(
        [Container] $WordDocument,
        [PSCustomObject] $Parameters
    )

    if ($Parameters.OverWriteTitle) {
        [Alignment] $TitleAlignment = $Parameters.OverwriteTitleAlignment
        [System.Drawing.Color] $TitleColor = $Parameters.OverwriteTitleColor

        $Table = Add-WordTable -WordDocument $WordDocument -Supress $false -DataTable $Parameters.DataTable -Design $Parameters.Design -AutoFit $Parameters.AutoFit #-DoNotAddTitle
        $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll  # -ColumnNrStart 0 -ColumnNrEnd 1
        $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
        $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $Parameters.OverwriteTitle -Alignment $TitleAlignment -Color $TitleColor
    } else {
        $Table = Add-WordTable -WordDocument $WordDocument -Supress $true -DataTable $Parameters.DataTable -Design $Parameters.Design -AutoFit $Parameters.AutoFit
    }

}