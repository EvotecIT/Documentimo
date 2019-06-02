function DocNumbering {
    [CmdletBinding()]
    [alias('DocumentimoNumbering', 'New-DocumentimoNumbering')]
    param(
        [Parameter(Position = 0)][ScriptBlock] $Content,
        [string] $Text,
        [int] $Level = 0,
        [ListItemType] $Type = [ListItemType]::Numbered,
        [HeadingType] $Heading = [HeadingType]::Heading1
    )
    [PSCustomObject] @{
        ObjectType = 'TocItem'
        Text       = $Text
        Content    = & $Content
        Level      = $Level
        Type       = $Type
        Heading    = $Heading
    }
}