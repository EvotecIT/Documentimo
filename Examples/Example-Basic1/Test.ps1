Import-Module .\Documentimo.psd1 -Force

$Table = Get-Process | Select-Object -First 5

Documentimo -FilePath $PSScriptRoot\Test.docx {
    DocTableOfContents -Title 'Table of content'

    DocPageBreak -Verbose

    DocNumbering -Text 'My document' -Level 0 -ItemType Numbered -HeadingType Heading1 {
        DocText -Text 'Test', ' Test2' -Color Yellow
        DocTable -DataTable $Table -Design ColorfulGrid
        DocPageBreak
    }
} -Open