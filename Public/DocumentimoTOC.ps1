function DocToc {
    [CmdletBinding()]
    [alias('DocumentimoTOC', 'New-DocumentimoTOC')]
    param(
        [string] $Title,
        [int] $RightTabPos,
        [TableOfContentsSwitches] $Switches
    )
    [PSCustomObject] @{
        ObjectType  = 'TOC'
        Title       = $Title
        RightTabPos = $RightTabPos
        Switches    = $Switches
    }
}