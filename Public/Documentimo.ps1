function Documentimo {
    [alias('Doc', 'New-Documentimo')]
    param(
        [Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Documentimo requires opening and closing brace."),
        [string] $FilePath,
        [switch] $Open
    )
    $WordDocument = New-WordDocument -FilePath $FilePath

    $Script:Parameters.WordDocument = $WordDocument
    Invoke-Command -ScriptBlock $Content

    Save-WordDocument -WordDocument $WordDocument -Supress $true -Language 'en-US' -Verbose -OpenDocument:$Open
}