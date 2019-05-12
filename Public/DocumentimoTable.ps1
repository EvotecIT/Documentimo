function DocumentimoTable {
    [alias('New-DocumentimoTable', 'DocTable')]
    param(
        [Array] $DataTable,
        [Xceed.Words.NET.AutoFit] $AutoFit = [Xceed.Words.NET.AutoFit]::Window,
        [Xceed.Words.NET.TableDesign] $Design,
        [Xceed.Words.NET.Direction] $Direction = [Xceed.Words.NET.Direction]::LeftToRight,
        [int] $MaximumColumns = 5,
        [string] $OverwriteTitle
    )

    [PSCustomObject] @{
        ObjectType     = 'Table'
        DataTable      = $DataTable
        AutoFit        = $AutoFit
        Design         = $Design
        Direction      = $Direction
        MaximumColumns = $MaximumColumns
        OverwriteTitle = $OverwriteTitle
    }
}