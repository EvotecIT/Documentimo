function DocumentimoTable {
    [alias('New-DocumentimoTable', 'DocTable')]
    param(
        [Array] $DataTable,
        [AutoFit] $AutoFit,
        [TableDesign] $Design,
        [Direction] $Direction,
        [switch] $BreakPageAfterTable,
        [switch] $BreakPageBeforeTable,
        [nullable[bool]] $BreakAcrossPages,
        [nullable[int]] $MaximumColumns,
        [string[]]$Titles = @('Name', 'Value'),
        [switch] $DoNotAddTitle,
        [alias ("ColummnWidth")][float[]] $ColumnWidth = @(),
        [nullable[float]] $TableWidth = $null,
        [bool] $Percentage,
        [alias ("C")] [System.Drawing.Color[]]$Color = @(),
        [alias ("S")] [double[]] $FontSize = @(),
        [alias ("FontName")] [string[]] $FontFamily = @(),
        [alias ("B")] [nullable[bool][]] $Bold = @(),
        [alias ("I")] [nullable[bool][]] $Italic = @(),
        [alias ("U")] [UnderlineStyle[]] $UnderlineStyle = @(),
        [alias ('UC')] [System.Drawing.Color[]]$UnderlineColor = @(),
        [alias ("SA")] [double[]] $SpacingAfter = @(),
        [alias ("SB")] [double[]] $SpacingBefore = @(),
        [alias ("SP")] [double[]] $Spacing = @(),
        [alias ("H")] [highlight[]] $Highlight = @(),
        [alias ("CA")] [CapsStyle[]] $CapsStyle = @(),
        [alias ("ST")] [StrikeThrough[]] $StrikeThrough = @(),
        [alias ("HT")] [HeadingType[]] $HeadingType = @(),
        [int[]] $PercentageScale = @(), # "Value must be one of the following: 200, 150, 100, 90, 80, 66, 50 or 33"
        [Misc[]] $Misc = @(),
        [string[]] $Language = @(),
        [int[]]$Kerning = @(), # "Value must be one of the following: 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48 or 72"
        [nullable[bool][]]$Hidden = @(),
        [int[]]$Position = @(), #  "Value must be in the range -1585 - 1585"
        [single[]] $IndentationFirstLine = @(),
        [single[]] $IndentationHanging = @(),
        [Alignment[]] $Alignment = @(),
        [Direction[]] $DirectionFormatting = @(),
        [ShadingType[]] $ShadingType = @(),
        [Script[]] $Script = @(),
        [nullable[bool][]] $NewLine = @(),
        [switch] $KeepLinesTogether,
        [switch] $KeepWithNextParagraph,
        [switch] $ContinueFormatting,
        [alias('Rotate', 'RotateData', 'TransposeColumnsRows', 'TransposeData')][switch] $Transpose,
        [string[]] $ExcludeProperty,
        [switch] $NoAliasOrScriptProperties,
        [switch] $DisplayPropertySet,
        [string] $OverwriteTitle
    )


    if ($OverWriteTitle) {
        #[nullable[int]] $TableMaximumColumns = 5,
        # [nullable[bool]] $TableTitleMerge,
        #[string] $TableTitleText,
        [Alignment] $TitleAlignment = [Alignment]::center
        [System.Drawing.Color] $TitleColor = [System.Drawing.Color]::Black


        $Table = Add-WordTable -WordDocument $Script:Parameters.WordDocument -Supress $false -DataTable $DataTable -Design $Design -AutoFit $AutoFit #-DoNotAddTitle
        $Table = Set-WordTableRowMergeCells -Table $Table -RowNr 0 -MergeAll  # -ColumnNrStart 0 -ColumnNrEnd 1
        $TableParagraph = Get-WordTableRow -Table $Table -RowNr 0 -ColumnNr 0
        $TableParagraph = Set-WordText -Paragraph $TableParagraph -Text $OverwriteTitle -Alignment $TitleAlignment -Color $TitleColor
    } else {
        $Table = Add-WordTable -WordDocument $Script:Parameters.WordDocument -Supress $true -DataTable $DataTable -Design $Design -AutoFit $AutoFit
    }
}