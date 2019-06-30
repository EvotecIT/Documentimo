function DocTable {
    [CmdletBinding()]
    [alias('DocumentimoTable', 'New-DocumentimoTable')]
    param(
        [Array] $DataTable,
        [AutoFit] $AutoFit = [AutoFit]::Window,
        [TableDesign] $Design = [TableDesign]::LightGrid,
        [Direction] $Direction = [Direction]::LeftToRight,
        [int] $MaximumColumns = 5,
        [string] $OverwriteTitle,
        [Alignment] $OverwriteTitleAlignment = [Alignment]::center,
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