function Documentimo {
    [alias('Doc', 'New-Documentimo')]
    param(
        [Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Documentimo requires opening and closing brace."),
        [string] $FilePath,
        [switch] $Open
    )
    $WordDocument = New-WordDocument -FilePath $FilePath
    #New-WordProcessing -Content $Content -WordDocument $WordDocument

    if ($null -ne $Content) {
        $Array = Invoke-Command -ScriptBlock $Content
        foreach ($Parameters in $Array) {
            if ($Parameters.ObjectType -eq 'List') {
                New-DocWordList -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'Table') {
                New-DocWordTable -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'TOC') {
                New-DocWordTOC -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'Text') {
                New-DocWordText -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'TocItem') {
                New-DocWordNumbering -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'PageBreak') {
                New-DocWordPageBreak -WordDocument $WordDocument -Parameters $Parameters
            } elseif ($Parameters.ObjectType -eq 'ChartPie') {
                New-DocWordChart -WordDocument $WordDocument -Parameters $Parameters
            } else {
                $Parameters
            }
        }
    }

    Save-WordDocument -WordDocument $WordDocument -Supress $true -Language 'en-US' -Verbose -OpenDocument:$Open
}


