function DocTable {
    [CmdletBinding()]
    [alias('DocumentimoTable', 'New-DocumentimoTable')]
    param(
        [Array] $DataTable,
        [Xceed.Words.NET.AutoFit] $AutoFit = [Xceed.Words.NET.AutoFit]::Window,
        [Xceed.Words.NET.TableDesign] $Design,
        [Xceed.Words.NET.Direction] $Direction = [Xceed.Words.NET.Direction]::LeftToRight,
        [int] $MaximumColumns = 5,
        [string] $OverwriteTitle,
        [Xceed.Words.NET.Alignment] $OverwriteTitleAlignment = [Xceed.Words.NET.Alignment]::center,
        [System.Drawing.Color] $OverwriteTitleColor = [System.Drawing.Color]::Black
    )
    [PSCustomObject] @{
        ObjectType              = 'Table'
        DataTable               = $DataTable
        AutoFit                 = $AutoFit
        Design                  = $Design
        Direction               = $Direction
        MaximumColumns          = $MaximumColumns
        OverwriteTitle          = $OverwriteTitle
        OverwriteTitleAlignment = $OverwriteTitleAlignment
        OverwriteTitleColor     = $OverwriteTitleColor
    }
}