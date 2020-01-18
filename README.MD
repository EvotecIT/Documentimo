<p align="center">
  <a href="https://www.powershellgallery.com/packages/Documentimo"><img src="https://img.shields.io/powershellgallery/v/Documentimo.svg"></a>
  <a href="https://www.powershellgallery.com/packages/Documentimo"><img src="https://img.shields.io/powershellgallery/vpre/Documentimo.svg?label=powershell%20gallery%20preview&colorB=yellow"></a>
  <a href="https://github.com/EvotecIT/Documentimo"><img src="https://img.shields.io/github/license/EvotecIT/Documentimo.svg"></a>
</p>

<p align="center">
  <a href="https://www.powershellgallery.com/packages/PSWriteHTML"><img src="https://img.shields.io/powershellgallery/p/Documentimo.svg"></a>
  <a href="https://github.com/EvotecIT/Documentimo"><img src="https://img.shields.io/github/languages/top/evotecit/Documentimo.svg"></a>
  <a href="https://github.com/EvotecIT/Documentimo"><img src="https://img.shields.io/github/languages/code-size/evotecit/Documentimo.svg"></a>
  <a href="https://www.powershellgallery.com/packages/Documentimo"><img src="https://img.shields.io/powershellgallery/dt/Documentimo.svg"></a>
</p>

<p align="center">
  <a href="https://twitter.com/PrzemyslawKlys"><img src="https://img.shields.io/twitter/follow/PrzemyslawKlys.svg?label=Twitter%20%40PrzemyslawKlys&style=social"></a>
  <a href="https://evotec.xyz/hub"><img src="https://img.shields.io/badge/Blog-evotec.xyz-2A6496.svg"></a>
  <a href="https://www.linkedin.com/in/pklys"><img src="https://img.shields.io/badge/LinkedIn-pklys-0077B5.svg?logo=LinkedIn"></a>
</p>

# Documentimo - PowerShell Module

`Documentimo` is a Powershell module that has a single goal creating easy to use Word based documentation

## Overview

Simplified version of [PSWriteWord](https://github.com/EvotecIT/PSWriteWord) PowerShell Module. Easy to use, easier to remember syntax. It's still in very alpha stage.
Since version 0.0.7 it is now integrated back to **PSWriteWord** and this project is archived. If you want to use Documentimo features simply uninstall Documentimo and install PSWriteWord.

```PowerShell
Uninstall-Module Documentimo -AllVersions
Install-Module PSWriteWord -Force
```

This project is archived. For development go to https://github.com/EvotecIT/PSWriteWord

### Changelog

- 0.0.7 - 2020.01.18
  - Added TableDesign default
  - Migrated back to PSWriteWord. This module is now meta module with only PSD1 file.
- 0.0.6 - 2019.06.23
  - Support for PSWriteWord 1.0
  - Support for PSSharedGoods 0.0.79
- 0.0.5 - 2019.05.15
  - Fix for problems with content within content
- 0.0.4 - 2019.05.12
  - First public release

### Install

```powershell
Install-Module Documentimo -Force
```

### Example 1

![image](https://evotec.xyz/wp-content/uploads/2019/05/img_5cd7ffc9bab8f.png)

```powershell
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