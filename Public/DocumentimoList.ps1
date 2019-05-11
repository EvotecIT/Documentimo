function DocumentimoList {
    [alias('DocList', 'New-DocumentimoList')]
    param(
        [ScriptBlock] $ListItems,
        [alias('ListType')][Xceed.Words.NET.ListItemType] $Type = [Xceed.Words.NET.ListItemType]::Bulleted
        #[parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.Container] $WordDocument,
        #[parameter(ValueFromPipelineByPropertyName, ValueFromPipeline)][Xceed.Words.NET.InsertBeforeOrAfter] $Paragraph
        # [ListItemType]$ListType = [ListItemType]::Bulleted,
        #  [alias('DataTable')][object] $ListData = $null,
        #   [InsertWhere] $InsertWhere = [InsertWhere]::AfterSelf,
        #   $BehaviourOption = 0,
        #      [Array] $ListLevels = @(),
        #   [bool] $Supress = $false
    )

    [PSCustomObject] @{
        ObjectType = 'List'
        ListItems  = Invoke-Command -ScriptBlock $ListItems
        Type       = $Type
    }

    <#
    $ListItemsArray = & $ListItems
    $List = $null
    foreach ($Item in $ListItemsArray) {
        if ($null -eq $List) {
            $List = $($Script:WordDocument).AddList($Item.Text, $Item.Level, $Item.Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
            #$Paragraph = $List.Items[$List.Items.Count - 1]
        } else {
            $List = $($Script:Parameters.WordDocument).AddListItem($List, $Item.Text, $Item.Level, $Item.Type, $Item.StartNumber, $Item.TrackChanges, $Item.ContinueNumbering)
            # $Paragraph = $List.Items[$List.Items.Count - 1]
        }
    }
    Add-WordListItem -WordDocument ($Script:Parameters.WordDocument) -List $List
    #>
}

function DocumentimoListItem {
    [alias('DocListItem', 'New-DocumentimoListItem')]
    param(
        [ValidateRange(0, 8)] [int] $Level,
        [string] $Text,
        [nullable[int]] $StartNumber,
        [bool]$TrackChanges = $false,
        [bool]$ContinueNumbering = $false,
        [bool]$Supress = $false
    )

    [PSCustomObject] @{
        ObjectType        = 'ListItem'
        Level             = $Level
        Text              = $Text
        StartNumber       = $StartNumber
        TrackChanges      = $TrackChanges
        ContinueNumbering = $ContinueNumbering
    }

    <#
    $ListItem = New-WordListItem -WordDocument $Script:WordDocument -List $Script:List -ListLevel $Level -ListItemType $Type -ListValue $Text
    $Paragraph = Get-WordListItemParagraph -List $ListItem -LastItem
    Set-WordText -Paragraph $Paragraph -Color $Color -FontFamily Calibri -Supress $True
    return $ListItem
#>

}



<#

function Test {
    param(
        [scriptblock] $Test
    )
    $Text = 'Test'
    $Array = & $Test
    foreach ($A in $Array) {
        Write-Color $Text, ' ', $A.MyOtherText -Color Red, Blue, Blue
    }
}
function Test5 {
    param(
        $Text,
        $MyOtherText
    )
    $PSBoundParameters
    #Write-Color $Text, ' ', $MyOtherText -Color Red, Blue, Blue
}


Test {

    Test5 -MyOtherText 'Testooro'
    Test5 -MyOtherText 'Testooro1'
    Test5 -MyOtherText 'Testooro2'
}


#>