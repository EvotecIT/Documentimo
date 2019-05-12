function DocNumbering {
    [CmdletBinding()]
    [alias('DocumentimoNumbering', 'New-DocumentimoNumbering')]
    param(
        [Parameter(Position = 0)][ScriptBlock] $Content,
        [string] $Text,
        [int] $Level = 0,
        [Xceed.Words.NET.ListItemType] $Type = [Xceed.Words.NET.ListItemType]::Numbered,
        [Xceed.Words.NET.HeadingType] $Heading = [Xceed.Words.NET.HeadingType]::Heading1
    )
    [PSCustomObject] @{
        ObjectType = 'TocItem'
        Text       = $Text
        Content    = $Content
        Level      = $Level
        Type       = $Type
        Heading    = $Heading
    }
}