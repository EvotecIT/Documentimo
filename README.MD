<center>

[![PowerShellGallery Version](https://img.shields.io/powershellgallery/v/Documentimo.svg)](https://www.powershellgallery.com/packages/Documentimo)

[![PowerShellGallery Platform](https://img.shields.io/powershellgallery/p/Documentimo.svg)](https://www.powershellgallery.com/packages/Documentimo)
[![PowerShellGallery Preview Version](https://img.shields.io/powershellgallery/vpre/Documentimo.svg?label=powershell%20gallery%20preview&colorB=yellow)](https://www.powershellgallery.com/packages/Documentimo)

![Top Language](https://img.shields.io/github/languages/top/evotecit/Documentimo.svg)
![Code](https://img.shields.io/github/languages/code-size/evotecit/Documentimo.svg)
[![PowerShellGallery Downloads](https://img.shields.io/powershellgallery/dt/Documentimo.svg)](https://www.powershellgallery.com/packages/Documentimo)

</center>

# Documentimo - PowerShell Module
`Documentimo` is a Powershell module that has a single goal creating easy to use Word based documentation

### Overview:
Simplified version of [PSWriteWord](https://github.com/EvotecIT/PSWriteWord) PowerShell Module. Easy to use, easier to remember syntax. It's still in very alpha stage.

### Changelog
- 0.5.0 - 23.06.2019
  - Support for PSWriteWord 1.0
  - Support for PSSharedGoods 0.0.79
- 0.4.0 - 15.05.2019
  - Fix for problems with content within content
- 0.3.0 - 12.05.2019
  - First public release


### Install

```
Install-Module Documentimo -Force
```

### Example 1

![image](https://evotec.xyz/wp-content/uploads/2019/05/img_5cd7ffc9bab8f.png)

```
$Table = Get-Process | Select-Object -First 5

$TableForCharts = @(
    [PSCustomObject] @{ Name = 'Test 1'; SomeValue = 1 }
    [PSCustomObject] @{ Name = 'Test 2'; SomeValue = 5 }
    [PSCustomObject] @{ Name = 'Test 3'; SomeValue = 6 }
)

Documentimo -FilePath $PSScriptRoot\Test.docx {
    DocTOC -Title 'Table of content'

    DocNumbering -Text 'My document' -Level 0 -Type Numbered -Heading Heading1 {
        DocText -Text 'Test', ' Test2' -Color Yellow
        DocTable -DataTable $Table -Design ColorfulGrid
    }

    DocNumbering -Text 'Chart' {
        DocChart -Title 'Processes' -DataTable $Table  -Key 'ProcessName' -Value 'Handles'

    }
    DocNumbering -Text 'AnotherChart' {
        DocChart -Title 'Priviliged Group Members' -DataTable $TableForCharts  -Key 'Name' -Value 'SomeValue'
    }
    $Table1 = Get-Process | Select-Object -First 5

    DocTable -DataTable $Table -Design ColorfulGrid
    DocList {
        DocListItem -Text 'Test' -Level 0
        DocListItem -Text 'Test1' -Level 2
    }

    DocList -Type Numbered {
        DocListItem -Text 'Test' -Level 0
        DocListItem -Text 'Test1' -Level 2
    }

    DocTable -DataTable $Table1 -Design ColorfulGrid
} -Open
```


### Example 2

Sample use case is described here: https://evotec.xyz/what-do-we-say-to-writing-active-directory-documentation/

Example code: https://github.com/EvotecIT/Documentimo/blob/master/Examples/Example-PSWinDocumentation.AD/Starter-AD.ps1

Output:

![Image](https://evotec.xyz/wp-content/uploads/2019/05/img_5cd8006aabfde.png)
![Image](https://evotec.xyz/wp-content/uploads/2019/05/img_5cd8008d694d9.png)
![Image](https://evotec.xyz/wp-content/uploads/2019/05/img_5cd800a741a03.png)
![Image](https://evotec.xyz/wp-content/uploads/2019/05/img_5cd800bf42ba4.png)
![Image](https://evotec.xyz/wp-content/uploads/2019/05/img_5cd800ef8c547.png)