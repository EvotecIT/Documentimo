function DocumentimoPageBreak {
    [alias('New-DocumentimoPageBreak', 'DocPageBreak')]

    param(
        [int] $Count = 1
    )

    [PSCustomObject] @{
        ObjectType = 'PageBreak'
        Count      = $Count
    }
}

