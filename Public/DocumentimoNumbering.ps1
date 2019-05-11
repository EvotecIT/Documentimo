function DocumentimoNumbering {
    [alias('DocNumbering', 'New-DocumentimoNumbering')]
    param(
        [Parameter(Position = 0)][ScriptBlock] $Content,
        [string] $Text,
        [int] $Level = 0,
        [ListItemType] $Type = [ListItemType]::Numbered,
        [HeadingType] $Heading = [HeadingType]::Heading1
    )

    #$TOC = Add-WordTocItem -WordDocument $Script:Parameters.WordDocument -Text $Text -ListLevel $Level -ListItemType $ItemType -HeadingType $HeadingType

    # if ($Content) {
    #      Invoke-Command -ScriptBlock $Content
    # }

    [PSCustomObject] @{
        ObjectType = 'TocItem'
        Text       = $Text
        Content    = $Content
        Level      = $Level
        Type       = $Type
        Heading    = $Heading
    }
}