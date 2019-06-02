function DocList {
    [CmdletBinding()]
    [alias('DocumentimoList', 'New-DocumentimoList')]
    param(
        [ScriptBlock] $ListItems,
        [alias('ListType')][ListItemType] $Type = [ListItemType]::Bulleted
    )

    [PSCustomObject] @{
        ObjectType = 'List'
        ListItems  = Invoke-Command -ScriptBlock $ListItems
        Type       = $Type
    }
}