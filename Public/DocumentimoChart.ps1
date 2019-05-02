function DocumentimoChart {
    [alias('DocChart', 'New-DocumentimoChart')]
    param(
        [string] $ChartTitle,
        [string[]] $ChartKeys,
        $ChartValues,
        [ChartLegendPosition] $ChartLegendPosition = [ChartLegendPosition]::Bottom,
        [bool] $ChartLegendOverlay
    )

    if (($ChartKeys.Count -eq 0) -or ($ChartValues.Count -eq 0)) {
        # If chart had no values or keys it would create an empty chart and prevent saving of document in Word
        # Handling this case with TextNoData above
    } else {
        Add-WordPieChart -WordDocument $WordDocument -ChartName $ChartTitle -Names $ChartKeys -Values $ChartValues -ChartLegendPosition $ChartLegendPosition -ChartLegendOverlay $ChartLegendOverlay
    }
}