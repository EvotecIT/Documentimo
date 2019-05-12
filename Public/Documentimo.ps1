function Documentimo {
    [CmdletBinding()]
    [alias('Doc', 'New-Documentimo')]
    param(
        [Parameter(Position = 0)][ValidateNotNull()][ScriptBlock] $Content = $(Throw "Documentimo requires opening and closing brace."),
        [string] $FilePath,
        [switch] $Open
    )
    $WordDocument = New-WordDocument -FilePath $FilePath
    if ($null -ne $Content) {
        $Array = Invoke-Command -ScriptBlock $Content
        New-WordProcessing -Content $Array -WordDocument $WordDocument
    }
    Save-WordDocument -WordDocument $WordDocument -Supress $true -Language 'en-US' -Verbose -OpenDocument:$Open
}


