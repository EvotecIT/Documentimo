function Documentimo {
    [alias('Doc', 'New-Documentimo')]
    param(
        [Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Documentimo requires opening and closing brace."),
        [string] $FilePath,
        [switch] $Open
    )
    $WordDocument = New-WordDocument -FilePath $FilePath
    New-WordProcessing -Content $Content -WordDocument $WordDocument
    Save-WordDocument -WordDocument $WordDocument -Supress $true -Language 'en-US' -Verbose -OpenDocument:$Open
}


