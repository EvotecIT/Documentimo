@{
    Copyright = 'Evotec (c) 2011-2019. All rights reserved.'
    PrivateData = @{
        PSData = @{
            Tags = 'documentation', 'windows', 'word'
            ProjectUri = 'https://github.com/EvotecIT/Documentimo'
            IconUri = 'https://evotec.xyz/wp-content/uploads/2019/05/Documentimo.png'
            Prerelease = 'Preview1'
        }
    }
    Description = 'Simple project generating documentation to Microsoft Word.'
    PowerShellVersion = '5.1'
    FunctionsToExport = 'Documentimo', 'DocChart', 'DocList', 'DocListItem', 'DocNumbering', 'DocPageBreak', 'DocTable', 'DocText', 'DocToc'
    Author = 'Przemyslaw Klys'
    RequiredModules = @{
        ModuleVersion = '0.7.2'
        ModuleName = 'PSWriteWord'
        GUID = '6314c78a-d011-4489-b462-91b05ec6a5c4'
    }, @{
        ModuleVersion = '0.0.76'
        ModuleName = 'PSSharedGoods'
        GUID = 'ee272aa8-baaa-4edf-9f45-b6d6f7d844fe'
    }
    GUID = '90b8cc01-acf1-4ad3-8e4b-23fbe326cbc0'
    RootModule = 'Documentimo.psm1'
    AliasesToExport = 'Doc', 'New-Documentimo', 'DocumentimoChart', 'New-DocumentimoChart', 'DocumentimoList', 'New-DocumentimoList', 'DocumentimoListItem', 'New-DocumentimoListItem', 'DocumentimoNumbering', 'New-DocumentimoNumbering', 'DocumentimoPageBreak', 'New-DocumentimoPageBreak', 'DocumentimoTable', 'New-DocumentimoTable', 'DocumentimoText', 'New-DocumentimoText', 'DocumentimoTOC', 'New-DocumentimoTOC'
    ModuleVersion = '0.0.5'
    CompanyName = 'Evotec'
}