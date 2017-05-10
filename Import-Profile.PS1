<#
    .DESCRIPTION 
        Reads an Outlook for Mac 2016 Profile, and imports the contact lists to a currently open Outlook window


    .LINK
        http://ramblingcookiemonster.github.io/SQLite-and-PowerShell/

    
    .LINK
        https://github.com/RamblingCookieMonster/PSSQLite


#>

Import-Module "$PSScriptRoot\PSSQLite\PSSQLite.psd1"

#region Constants

#endregion

#region Functions

    Function Get-BinaryDistributionListMembers
    {
        Param(
            $Path
        )
        $delim  = "$([char]00)" * 11
        $delim2=  "$([char]00)" * 2
        $regex = [regex] '\W(\w.*?@\w.*?\.\w.*?)\W{4}([\w|\W]*)*?\W\W\W\W'
        $List = Get-Content -Raw -Path $Path
        $lines = $list -split $delim
        For ($i=1;$i -lt $lines.count;$i++)
        {
            $Line = $Lines[$i]
            #$Line | Set-Content "$PSScriptRoot\$i.bin"
            Try{
                $emailLen = [byte]$($line[0xa])
                $email = $line.Substring(14,$emailLen)
                $displayNameStart = 14+$emailLen
                $displayNameLen = "{0:d}" -f [byte]$($Line[$displayNameStart])
                $displayName = $line.Substring($displayNameStart+4,$displayNameLen) -replace("$([char]00)","")
            }
            catch {
                #Yes, I'm ashamed of myself for this.
            }
            finally
            {
                if($Email -like "*@*")
                {
                    New-Object -TypeName psobject -Property @{
                        Name = $displayName;
                        Email = $email
                    }
                }
            
            }
           
           

            <#
            }#>
            #Write-Host $($lines[$i]|ConvertTo-JSON)
           # Write-Host $([byte][char]$split[5])
            #Write-Host $($Split | ConvertTo-JSON)
            #Write-Host "$($Entry[0]) =============== $($Entry[1])"

            <#
            if ($lines[$i] -match $regex)
            {
                Write-Host $lines[$i]
               
            }#>

        
        }

    }

    Function Get-DistributionLists 
    {
        Param(
            $SQLiteFile
        )

        $Query = "select PathToDataFile, Contact_DisplayName from Contacts where Contacts.Contact_ContactRecType = '3'"


        $ContactLists =  Invoke-SqliteQuery -Query $Query -DataSource $SQLiteFile
        Foreach ($ContactList in $ContactLists)
        {
            $Path = "{0}{1}" -f $($SQLiteFile | Split-Path), $ContactList.PathToDataFile
            New-Object -TypeName psobject -Property @{
                ListName = $ContactList.Contact_DisplayName;
                Members = $(Get-BinaryDistributionListMembers -Path $Path)
            }
        }
    }


    Function Set-DistributionList
    {
        Param(
            $ListName,
            $Members
        )
        $outlook = new-object -com Outlook.Application
        $contacts = $outlook.Session.GetDefaultFolder(10)
        $session = $outlook.Session
        $session.Logon("Outlook")
        $namespace = $outlook.GetNamespace("MAPI")
        $DL = $contacts.Items.Add("IPM.DistList")
        $DL.DLName = $ListName
        Foreach($Member in $Members)
        {
            $recipient = $namespace.CreateRecipient($Member.Email)  # this has to be an exsiting contact
            $recipient.Resolve()  # check if this returns true
            $DL.AddMember($recipient)
        }
       
        $DL.Save()
    }    
#end region

$Lists = Get-DistributionLists -SQLiteFile "D:\Outlook.sqlite"

Foreach( $List in $Lists)
{
    Set-DistributionList -ListName $List.ListName -Members $List.Members

}