function DocumentimoTableOfContents {
    [alias('New-DocumentimoTOC', 'DocToc', 'DocumentimoTOC', 'DocTableOfContents')]
    param(
        [string] $Title,
        [int] $RightTabPos,
        [TableOfContentsSwitches[]] $Switches
    )
    Add-WordToc -WordDocument $Script:Parameters.WordDocument -Title $Title -Switches $Switches -RightTabPos $RightTabPos -Supress $True
}