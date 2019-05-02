function DocumentimoNumbering {
    [alias('DocNumbering', 'New-DocumentimoNumbering')]
    param(
        [Parameter(Position = 0)][ScriptBlock] $Content,
        [string] $Text,
        [int] $Level = 0,
        [ListItemType] $ItemType = [ListItemType]::Numbered,
        [HeadingType] $HeadingType = [HeadingType]::Heading1
    )

    $TOC = Add-WordTocItem -WordDocument $Script:Parameters.WordDocument -Text $Text -ListLevel $Level -ListItemType $ItemType -HeadingType $HeadingType

    if ($Content) {
        Invoke-Command -ScriptBlock $Content
    }
}