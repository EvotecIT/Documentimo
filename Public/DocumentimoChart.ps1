function DocChart {
    [CmdletBinding()]
    [alias('DocumentimoChart', 'New-DocumentimoChart')]
    param(
        [Array] $DataTable,
        [string] $Title,
        [string] $Key,
        [string] $Value,
        [Xceed.Words.NET.ChartLegendPosition] $LegendPosition = [Xceed.Words.NET.ChartLegendPosition]::Bottom,
        [bool] $LegendOverlay
    )

    [PSCustomObject] @{
        ObjectType     = 'ChartPie'
        DataTable      = $DataTable
        Title          = $Title
        Key            = $Key
        Value          = $Value
        LegendPosition = $LegendPosition
        LegendOverlay  = $LegendOverlay
    }
}