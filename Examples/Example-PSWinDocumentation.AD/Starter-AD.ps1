Import-Module PSWinDocumentation.AD -Force
Import-Module .\Documentimo.psd1 -Force
Import-Module PSWriteWord -Force

if ($null -eq $ADForest) {
    $ADForest = Get-WinADForestInformation -Verbose -PasswordQuality
}

$CompanyName = 'Evotec'

Documentimo -FilePath "$PSScriptRoot\Starter-AD.docx" {
    DocTableOfContents -Title 'Table of content'

    DocPageBreak

    DocText {
        "This document provides low-level documentation of Active Directory infrastructure in Evotec organization. This document contains general data that has been exported from Active Directory and provides an overview of the whole environment."
    }
    DocNumbering -Text 'General Information - Forest Summary' -Level 0 -Type Numbered -Heading Heading1 {
        DocText {
            "Active Directory at $CompanyName has a forest name $($ADForest.Forest). Following table contains forest summary with important information:"
        }
        DocTable -DataTable $ADForest.ForestInformation -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Forest Summary'
        DocText -LineBreak

        DocText -Text 'Following table contains FSMO servers:'
        DocTable -DataTable $ADForest.ForestFSMO -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'FSMO Roles'
        DocText -LineBreak

        DocText -Text 'Following table contains optional forest features:'
        DocTable -DataTable $ADForest.ForestOptionalFeatures -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Optional Features'
        DocText -LineBreak

        DocText -Text 'Following UPN suffixes were created in this forest:'
        DocTable -DataTable $ADForest.ForestUPNSuffixes -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'UPN Suffixes'
        DocText -LineBreak

        if ($ADForest.ForestSPNSuffixes) {
            DocText -Text 'Following SPN suffixes were created in this forest:'
            DocTable -DataTable $ADForest.ForestSPNSuffixes -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'UPN Suffixes'
        } else {
            DocText -Text 'No SPN suffixes were created in this forest.'
        }
        DocPageBreak

        DocNumbering -Text 'General Information - Forest Sites' -Level 1 -Type Numbered -Heading Heading1 {
            DocText -Text "Forest Sites list can be found below:"
            DocTable -DataTable $ADForest.ForestSites1 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
            DocText -LineBreak

            DocText -Text "Forest Sites list can be found below:"
            DocTable -DataTable $ADForest.ForestSites2 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
            #  DocText -LineBreak
        }

        DocNumbering -Text 'General Information - Subnets' -Level 1 -Type Numbered -Heading Heading1 {
            DocText -Text "Table below contains information regarding relation between Subnets and sites"
            DocTable -DataTable $ADForest.ForestSubnets1 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
            DocText -LineBreak

            DocText -Text "Table below contains information regarding relation between Subnets and sites"
            DocTable -DataTable $ADForest.ForestSubnets2 -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
            # DocText -LineBreak
        }

        DocNumbering -Text 'General Information - Site Links' -Level 1 -Type Numbered -Heading Heading1 {
            DocText -Text "Forest Site Links information is available in table below"
            DocTable -DataTable $ADForest.ForestSiteLinks -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
            # DocText -LineBreak
        }
    }
    foreach ($Domain in $ADForest.FoundDomains.Keys) {
        DocPageBreak

        DocNumbering -Text "General Information - Domain $Domain" -Level 0 -Type Numbered -Heading Heading1 {

            DocNumbering -Text 'General Information - Domain Summary' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Following domain exists within forest $($ADForest.ForestName):"

                DocList -Type Bulleted {
                    DocListItem -Level 0 -Text "Domain $($ADForest.FoundDomains.'ad.evotec.xyz'.DomainInformation.DistinguishedName)"
                    DocListItem -Level 1 -Text "Name for fully qualified domain name (FQDN): $($ADForest.FoundDomains.'ad.evotec.xyz'.DomainInformation.DNSRoot)"
                    DocListItem -Level 1 -Text "Name for NetBIOS: $($ADForest.FoundDomains.'ad.evotec.xyz'.DomainInformation.NetBIOSName)"
                }

                # TO DO: DocList
                #DocText -LineBreak
                <#
                SectionDomainIntroduction                         = [ordered] @{
                    Use                   = $true
                    TocEnable             = $True
                    TocText               = 'General Information - Domain Summary'
                    TocListLevel          = 1
                    TocListItemType       = [ListItemType]::Numbered
                    TocHeadingType        = [HeadingType]::Heading1
                    Text                  = "Following domain exists within forest <ForestName>:"
                    ListBuilderContent    = "Domain <DomainDN>", 'Name for fully qualified domain name (FQDN): <Domain>', 'Name for NetBIOS: <DomainNetBios>'
                    ListBuilderLevel      = 0, 1, 1
                    ListBuilderType       = [ListItemType]::Bulleted, [ListItemType]::Bulleted, [ListItemType]::Bulleted
                    EmptyParagraphsBefore = 0
                }
                #>
            }

            DocNumbering -Text  'General Information - Domain Controllers' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text 'Following table contains domain controllers'
                DocTable -DataTable $ADForest.FoundDomains.'ad.evotec.xyz'.DomainControllers -Design ColorfulGridAccent5 -AutoFit Window #-OverwriteTitle 'Forest Summary'
                DocText -LineBreak

                DocText -Text "Following table contains FSMO servers with roles for domain $Domain"
                DocTable -DataTable $ADForest.FoundDomains.'ad.evotec.xyz'.DomainFSMO -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "FSMO Roles for $Domain"
            }

            DocNumbering -Text 'General Information - Password Policies' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Following table contains password policies for all users within $Domain"
                DocTable -DataTable $ADForest.FoundDomains.'ad.evotec.xyz'.DomainDefaultPasswordPolicy -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle "Default Password Policy for $Domain"
                #DocText -LineBreak
            }

            DocNumbering -Text 'General Information - Fine-grained Password Policies' -Level 1 -Type Numbered -Heading Heading1 {
                if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainFineGrainedPolicies) {
                    DocText -Text 'Following table contains Fine-grained password policies'
                    DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainFineGrainedPolicies -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle  "Fine-grained Password Policy for <Domain>"
                } else {
                    DocText {
                        "Following section should cover fine-grained password policies. " `
                            + "There were no fine-grained password polices defined in $Domain. There was no formal requirement to have them set up."
                    }
                }
            }
            DocNumbering -Text 'General Information - Group Policies' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Following table contains group policies for $Domain"
                DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainGroupPolicies -Design ColorfulGridAccent5 -AutoFit Window
                #DocText -LineBreak
            }
            DocNumbering -Text 'General Information - Group Policies Details' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Following table contains group policies for $Domain"
                DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainGroupPoliciesDetails -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 6
                #DocText -LineBreak
            }

            DocNumbering -Text 'General Information - DNS A/SRV Records' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Following table contains SRV records for Kerberos and LDAP"
                DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainDNSSRV -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 10
                DocText -LineBreak

                DocText -Text "Following table contains A records for Kerberos and LDAP"
                DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainDNSA -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 10
            }


            DocNumbering -Text 'General Information - Trusts' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Following table contains trusts established with domains..."
                DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainTrusts -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 10
                DocText -LineBreak
            }

            DocNumbering -Text 'General Information - Organizational Units' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text "Following table contains all OU's created in $Domain"
                DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainOrganizationalUnits -Design ColorfulGridAccent5 -AutoFit Window -MaximumColumns 4
                DocText -LineBreak
            }

            DocNumbering -Text 'General Information - Priviliged Groups' -Level 1 -Type Numbered -Heading Heading1 {
                DocText -Text 'Following table contains list of priviliged groups and count of the members in it.'
                DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainGroupsPriviliged -Design ColorfulGridAccent5 -AutoFit Window
                DocChart -Title 'Priviliged Group Members' -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainGroupsPriviliged -Key 'Group Name' -Value 'Member Count'
            }

            DocNumbering -Text "General Information - Domain Users in $Domain" -Level 1 -Type Numbered -Heading Heading1 {

                #DocText -Text 'Following table contains list of priviliged groups and count of the members in it.'


                DocNumbering -Text 'General Information - Users Count' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText -Text "Following table and chart shows number of users in its categories"
                    DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainUsersCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Users Count'
                    DocChart -Title 'Servers Count' -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainUsersCount
                }

                DocNumbering -Text 'General Information - Domain Administrators' -Level 2 -Type Numbered -Heading Heading2 {

                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainAdministratorsRecursive) {
                        DocText -Text 'Following users have highest priviliges and are able to control a lot of Windows resources.'
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainAdministratorsRecursive -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText -Text 'No Domain Administrators users were defined for this domain.'
                    }
                }


                DocNumbering -Text  'General Information - Enterprise Administrators' -Level 2 -Type Numbered -Heading Heading2 {

                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainEnterpriseAdministratorsRecursive) {
                        DocText -Text 'Following users have highest priviliges across Forest and are able to control a lot of Windows resources.'
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainEnterpriseAdministratorsRecursive -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText -Text 'No Enterprise Administrators users were defined for this domain.'
                    }
                }
            }

            DocNumbering -Text "General Information - Computer Objects in $Domain" -Level 1 -Type Numbered -Heading Heading1 {


                DocNumbering -Text 'General Information - Computers' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText -Text "Following table and chart shows number of computers and their versions"
                    DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainComputersCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Computers Count'
                    DocChart -Title 'Servers Count' -DataTable $ADForest.FoundDomains.'ad.evotec.xyz'.DomainComputersCount -Key 'System Name' -Value  'System Count'
                }
                DocNumbering -Text 'General Information - Servers' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText -Text "Following table and chart shows number of servers and their versions"
                    DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainServersCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Servers Count'
                    DocChart -Title 'Servers Count' -DataTable $ADForest.FoundDomains.'ad.evotec.xyz'.DomainServersCount -Key 'System Name' -Value  'System Count'
                }
                DocNumbering -Text 'General Information - Unknown Computers' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText -Text "Following table and chart shows number of unknown object computers in domain."
                    DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainComputersUnknownCount -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Unknown Computers Count'
                    DocChart -Title 'Servers Count' -DataTable $ADForest.FoundDomains.'ad.evotec.xyz'.DomainComputersUnknownCount -Key 'System Name' -Value  'System Count'
                }
            }

            DocNumbering -Text 'Domain Password Quality' -Level 1 -Type Numbered -Heading Heading1 {
                Doctext {
                    "This section provides overview about password quality used in $Domain. One should review if all those potentially" `
                        + " dangerous approaches to password quality should be left as is or addressed in one way or another."
                }

                DocNumbering -Text 'Password Quality - Passwords with Reversible Encryption' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText -Text 'Passwords of these accounts are stored using reversible encryption.'

                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordClearTextPassword) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordClearTextPassword -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText -Text 'There are no accounts that have passwords stored using reversible encryption.'
                    }
                }
                DocNumbering -Text 'Password Quality - Passwords with LM Hash' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'LM-hashes is the oldest password storage used by Windows, dating back to OS/2 system.' `
                            + ' Due to the limited charset allowed, they are fairly easy to crack. Following accounts are affected:'
                    }

                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordLMHash) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordLMHash -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            'LM-hashes is the oldest password storage used by Windows, dating back to OS/2 system.' `
                                + ' There were no accounts found that use LM Hashes.'
                        }
                    }
                }
                DocNumbering -Text 'Password Quality - Empty Passwords' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText -Text  'Following accounts have no password set:'

                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordEmptyPassword) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordEmptyPassword -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText -Text "There are no accounts in $Domain that have no password set."
                    }
                }
                DocNumbering -Text 'Password Quality - Known passwords' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        "Passwords of these accounts have been found in given dictionary. It's highly recommended to " `
                            + "notify those users and ask them to change their passwords asap!"
                    }

                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordWeakPassword) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordWeakPassword -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText -Text 'There were no passwords found that match given dictionary.'
                    }
                }
                DocNumbering -Text 'Password Quality - Default Computer Password' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText -Text 'These computer objects have their password set to default:'

                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDefaultComputerPassword) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDefaultComputerPassword -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText -Text 'There were no accounts found that match default computer password criteria.'
                    }
                }
                DocNumbering -Text 'Password Quality - Password Not Required' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'These accounts are not required to have a password. For some accounts it may be perfectly acceptable ' `
                            + ' but for some it may not. Those accounts should be reviewed and accepted or changed to proper security.'
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordPasswordNotRequired) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordPasswordNotRequired -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            'There were no accounts found that does not require password.'
                        }
                    }
                }
                DocNumbering -Text  'Password Quality - Non expiring passwords' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'Following account have do not expire password policy set on them. Those accounts should be reviewed whether ' `
                            + 'allowing them to never expire is good idea and accepted risk.'
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordPasswordNeverExpires) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordPasswordNeverExpires -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            "There are no accounts in $Domain that never expire."
                        }
                    }
                }
                DocNumbering -Text 'Password Quality - AES Keys Missing' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'Following accounts have their Kerberos AES keys missing'
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordAESKeysMissing) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordAESKeysMissing -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            'There are no accounts that hvae their Kerberos AES keys missing.'
                        }
                    }
                }
                DocNumbering -Text 'Password Quality - Kerberos Pre-Auth Not Required' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'Kerberos pre-authentication is not required for these accounts'
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordPreAuthNotRequired) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordPreAuthNotRequired -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            'There were no accounts found that do not require pre-authentication.'
                        }
                    }
                }
                DocNumbering -Text 'Password Quality - Only DES Encryption Allowed' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'Only DES encryption is allowed to be used with these accounts'
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDESEncryptionOnly) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDESEncryptionOnly -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            'There are no account that require only DES encryption.'
                        }
                    }
                }
                DocNumbering -Text 'Password Quality - Delegatable to Service' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'These accounts are allowed to be delegated to a service:'
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDelegatableAdmins) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDelegatableAdmins -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            'No accounts were found that are allowed to be delegated to a service.'
                        }
                    }
                }
                DocNumbering -Text 'Password Quality - Groups of Users With Same Password' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        'Following groups of users have same passwords:'
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDuplicatePasswordGroups) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordDuplicatePasswordGroups -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            "There are no 2 passwords that are the same in $Domain."
                        }
                    }
                }
                DocNumbering -Text 'Password Quality - Leaked Passwords' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        "Passwords of these accounts have been found in given HASH dictionary (https://haveibeenpwned.com/). It's highly recommended to " `
                            + "notify those users and ask them to change their passwords asap!"
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordHashesWeakPassword) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordHashesWeakPassword -Design ColorfulGridAccent5 -AutoFit Window
                    } else {
                        DocText {
                            'There were no passwords found that match in given dictionary.'
                        }
                    }
                }
                DocNumbering -Text  'Password Quality - Statistics' -Level 2 -Type Numbered -Heading Heading2 {
                    DocText {
                        "Following table and chart shows password statistics"
                    }
                    if ($ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordStats) {
                        DocTable -DataTable  $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordStats -Design ColorfulGridAccent5 -AutoFit Window -OverwriteTitle 'Password Quality - Statistics'
                    } else {
                        DocText {
                            'There were no passwords found that match in given dictionary.'
                        }
                    }
                    DocChart -Title 'Password Statistics' -DataTable $ADForest.FoundDomains.'ad.evotec.xyz'.DomainPasswordStats # Hashtables don't require Key/Value pair
                }
            }
        }
    }
} -Open