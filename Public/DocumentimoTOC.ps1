function DocumentimoTableOfContents {
    [alias('New-DocumentimoTOC', 'DocToc', 'DocumentimoTOC', 'DocTableOfContents')]
    param(
        [string] $Title,
        [int] $RightTabPos,
        [Xceed.Words.Net.TableOfContentsSwitches] $Switches
    )
    [PSCustomObject] @{
        ObjectType  = 'TOC'
        Title       = $Title
        RightTabPos = $RightTabPos
        Switches    = $Switches
    }

    #Add-WordToc -WordDocument $Script:Parameters.WordDocument -Title $Title -Switches $Switches -RightTabPos $RightTabPos -Supress $True
}