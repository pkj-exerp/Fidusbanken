# Default values
$defaultTitle = "Fidusbanken user management"
$defaultSelectionMsg = "Please make a selection"
$defaultEmailDomain = "fidusbanken.dk"
$defaultDomainController = "DC=omit,DC=localhost"

<#
showMenu
Shows the menu for the selection
#>
function showMenu
{
    $title = $defaultTitle
    Clear-Host
    Write-Host "================ $title ================"
    
    Write-Host "1: Press '1' to create a single user."
    Write-Host "2: Press '2' to create multiple users."
    Write-Host "3: Press '3' to create users from CSV."
    Write-Host "4: Press '4' to update a already exisiting user."
    Write-Host "Q: Press 'Q' to quit."
}
<#
selection
Creates the selection menu
#>
function selection 
{
    switch ($selectionMenu)
    {
        '1' 
        {
            singleUser
        }

        '2'
        {
            multipleUsers
        }

        '3' 
        {
            usersCSV
        }

        '4' 
        {
            updateUser
        }

        'q' 
        {
            return
        }
        default
        {
            Write-Warning -Message "Invalid option, try again"
            $selectionMenu = Read-Host $defaultSelectionMsg
            selection
        }
    }
}
<#
Get-RandomPassword
Creates a random password with 3 of each characters
Will randomize the the password
#>
function Get-RandomPassword 
{
    $length = 3

    $numbers = 1..9
    $lettersLower = 'abcdefghijklmnopqrstuvwxyz'.ToCharArray()
    $lettersUpper = 'ABCEDEFHIJKLMNOPQRSTUVWXYZ'.ToCharArray()
    $special = '!@#$%^&*()=+[{}]/?<>'.ToCharArray()

    $numbers = Get-Random -Count $length $numbers
    $lettersLower = Get-Random -Count $length $lettersLower
    $lettersUpper = Get-Random -Count $length $lettersUpper
    $special = Get-Random -Count $length $special

    $password = $numbers+$lettersLower+$lettersUpper+$special | Sort-Object {Get-Random}
    $password = -join $password

    return $password
}
<#
Get-OU's
Gets all the OU that exist in the AD
#>
function Get-OUs
{
    $OUList = Get-ADOrganizationalUnit -Filter * | Select-Object -Property Name
    foreach ($item in $OUList)
    {
        '{0} - {1}' -f ($OUList.IndexOf($item) + 1), $item.Name
    }
}
<#
Set-OU
Sets the OU based on the choice from end user
#>
function Set-OU
{
    $choice = ""
    $OUList = Get-ADOrganizationalUnit -Filter * | Select-Object -Property Name
    while ([string]::IsNullOrEmpty($choice))
    {
        $choice = Read-Host 'Please choose an item by number'
        if ($choice -notin 1..$OUList.Count)
        {
            Write-Warning ('Your choice [ {0} ] is not valid.' -f $choice)
            Write-Warning ('The valid choices are 1 thru {0}.' -f $OUList.Count)

            $choice = ''
        }
    }

    [string]$ou = $OUList[$choice - 1]
    $ou = $ou.Substring($ou.IndexOf('=')+1)
    $ou = $ou.Substring(0, $ou.Length-1)

    return $ou
}
<#
SingleUser function
Creates a singular user with input from end user
#>
function singleUser
{
    $title = $defaultTitle + " - " + "Single user creation"
    Clear-Host
    Write-Host "================ $title ================"
    $fullname = Read-Host "Insert Full name"
    $fullnameSplit = $fullname.Split()

    $firstname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[0]) # -Name + -GivenName
    $surname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[1]) # -Surname
    $initials = $firstname.Substring(0,2) + $surname.Substring(0,2)
    $initials = $initials.ToLower() # -SamAccountName
    $email = $initials + $defaultEmailDomain # -UserPrincipalName
    Get-OUs
    $ou = Set-OU
    $path = "OU="+$ou+","+$defaultDomainController # -Path
    $password = get-RandomPassword | ConvertTo-SecureString -AsPlainText -Force # -AccountPassword
    New-ADUser -Name $firstname -Surname $surname -GivenName $firstname -SamAccountName $initials -UserPrincipalName $email -AccountPassword $password -path $path -Enabled $true
}

function multipleUsers
{
    $title = $defaultTitle + " - " + "Multiple user creation"
    Clear-Host
    Write-Host "================ $title ================"
    Get-OUs
    $ou = Set-OU
    $path = "OU="+$ou+","+$defaultDomainController
    write-host $path
}
function usersCSV
{
    $title = $defaultTitle + " - " + "CSV user creation"
    Clear-Host
    Write-Host "================ $title ================"
}

function updateUser
{
    $title = $defaultTitle + " - " + "Update already exisiting user"
    Clear-Host
    Write-Host "================ $title ================"
}

<#
Initializes the script
#>
showMenu
$selectionMenu = Read-Host $defaultSelectionMsg
selection