function DocList {
    [CmdletBinding()]
    [alias('DocumentimoList', 'New-DocumentimoList')]
    param(
        [ScriptBlock] $ListItems,
        [alias('ListType')][Xceed.Words.NET.ListItemType] $Type = [Xceed.Words.NET.ListItemType]::Bulleted
    )

    [PSCustomObject] @{
        ObjectType = 'List'
        ListItems  = Invoke-Command -ScriptBlock $ListItems
        Type       = $Type
    }
}