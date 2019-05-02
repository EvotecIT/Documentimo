function DocumentimoText {
    [alias('New-DocumentimoText', 'DocText')]
    param(
        [Parameter(Mandatory = $false, Position = 0)][ScriptBlock] $TextBlock,
        [alias ("T")] [String[]]$Text,
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
        [nullable[bool][]]$NewLine = @(),
        # [switch] $KeepLinesTogether, # not done
        # [switch] $KeepWithNextParagraph, # not done
        [single[]] $IndentationFirstLine = @(),
        [single[]] $IndentationHanging = @(),
        [Alignment[]] $Alignment = @(),
        [Direction[]] $Direction = @(),
        [ShadingType[]] $ShadingType = @(),
        [System.Drawing.Color[]]$ShadingColor = @(),
        [Script[]] $Script = @(),
        [Switch] $ContinueFormatting,
        [alias ("Append")][Switch] $AppendToExistingParagraph,
        [bool] $Supress = $false,
        [switch] $LineBreak
    )
    if ($TextBlock) {
        $Text = (Invoke-Command -ScriptBlock $TextBlock)
        #if ($Text.Count) {
        #    $LineBreak = $false
        #}
    }
    Add-WordText -WordDocument $Script:Parameters.WordDocument -Text $Text -Color $Color -Supress $true
    if ($LineBreak) {
        Add-WordParagraph -WordDocument $Script:Parameters.WordDocument -Supress $True
    }
}