# Default values
$defaultTitle = "Fidusbanken user management"
$defaultSelectionMsg = "Please make a selection"
$defaultEmailDomain = "fidusbanken.dk"
$defaultDomainController = "DC=Oddermose-IT,DC=localhost"

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
    Write-Host "5: Press '5' to activate a user."
    Write-Host "6: Press '6' to deactivate a user."
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

        '5' 
        {
            activateUser
        }

        '6' 
        {
            deactivateUser
        }

        't'
        {
            test
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
function showMenuUpdateUser($initials)
{
    $title = $defaultTitle + " - " + "update for $initials"
    Clear-Host
    Write-Host "================ $title ================"
    
    Write-Host "1: Press '1' to update username."
    Write-Host "2: Press '2' to update password."
    Write-Host "3: Press '3' to update department."
    Write-Host "4: Press '4' to update fullname"
    Write-Host "Q: Press 'Q' to quit."
}
<#
selection
Creates the selection menu
#>
function selectionUpdateUser($initials)
{
    switch ($selectionMenuUpdateUser)
    {
        '1' 
        {
            updateUserUsername($initials)
        }

        '2'
        {
            updateUserPassword($initials)
        }

        '3' 
        {
            updateUserDepartment($initials)
        }

        '4' 
        {
            updateUserFullname($initials)
        }

        'q' 
        {
            return
        }

        default
        {
            Write-Warning -Message "Invalid option, try again"
            $selectionMenuUpdateUser = Read-Host $defaultSelectionMsg
            selection
        }
    }
}
function selectionTryAgain($tryagain)
{
    switch ($selectionTryAgain)
    {
        'y' 
        {
            if ($tryagain -eq "active")
            {
                activateUser
            }
            elseif ($tryagain -eq "deactive")
            {
                deactivateUser
            }
            elseif ($tryagain -eq "updateUser")
            {
                updateUser
            }
        }

        'q' 
        {
            return
        }
        default
        {
            Write-Warning -Message "Invalid option, try again"
            $selectionTryAgain = Read-Host $defaultSelectionMsg
            selectionTryAgain($tryagain)
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
IsSamAccountNameValid function
Takes in parameter initials and checks if the samaccountname exists
If the username exists it will call the function updateSamAccountName
#>
function IsSamAccountNameValid($initials)
{
    if (!(Get-ADUser -Filter "SamAccountName -eq '$($initials)'"))
    {
        #User doesn't exists
        #write-host "User doesn't exist - $initials"
        return $initials
    }
    else 
    {
        #write-host "user exists - $initials"
        #User exists
        #write-host "User exists" $initials
        updateSamAccountName($initials)
    }
}
<#
IsSamAccountNameValidUpdateUser function
Takes in parameter initials and checks if the samaccountname exists
If the username exists it will call the function updateSamAccountName
#>
function IsSamAccountNameValidUpdateUser($initials)
{
    $valid = $false
    if (!(Get-ADUser -Filter "SamAccountName -eq '$($initials)'"))
    {
        #User doesn't exists
        #write-host "User doesn't exist - $initials"
        return $valid
    }
    else 
    {
        #write-host "user exists - $initials"
        #User exists
        #write-host "User exists" $initials
        $valid = $true
        return $valid
    }
}
<#
updateSamAccountName function
Adds a number to the username
IF that username also exists it will increase the number until it doesn't exist anymore
#>
function updateSamAccountName($initials)
{
    $isvalid = $false
    $number = 1
    do
    {
        if (!(Get-ADUser -Filter "SamAccountName -eq '$($initials)'"))
        {
            #User doesn't exists
            #write-host "User is now valid " $initials
            $isvalid = $true
            IsSamAccountNameValid($initials)
        }
        else 
        {
            if ($initials -match ".*\d+.*")
            {
                $initials = $initials -split '(?<=\D)(?=\d)'
                $number = $initials[1] -as [int]
                $number++
                $initials = $initials[0]+$number
                #write-host "split: " $initials
            }
            else 
            {
                #write-host "Dosen't contain number " $initials
                $initials = $initials+$number
            }
            #Write-Host "In update function" $initials
        }
    }
    while(!$isvalid)
}
<#
IsOUVaid function
Checks if the OU exists
#>
function IsOUValid($ou)
{
    $ouExist = $false
    try 
    {
        Get-ADOrganizationalUnit -Identity $ou | Out-Null
        #Write-Host "OU exists - $ou"
        $ouExist = $true
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] 
    {
        #write-host "Ou DOESN'T exist - $ou"
        $ouExist = $false
    }

    return $ouExist
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

    $firstname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[0]) # -GivenName
    $surname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[1]) # -Surname

    $initials = $firstname.Substring(0,2) + $surname.Substring(0,2)
    $initials = $initials.ToLower() # -SamAccountName

    $initials = IsSamAccountNameValid($initials)

    $email = $initials + "@" + $defaultEmailDomain # -UserPrincipalName

    Get-OUs
    $ou = Set-OU
    $path = "OU="+$ou+","+$defaultDomainController # -Path
    $password = get-RandomPassword
    $password_secured = ConvertTo-SecureString($password) -AsPlainText -Force # -AccountPassword


    New-ADUser -GivenName $firstname -Surname $surname -Name $initials  -SamAccountName $initials -UserPrincipalName $email -AccountPassword $password_secured -path $path -Enabled $true
    Write-Host -ForegroundColor red "$fullname has been created"
    Write-Host -ForegroundColor red "Username: $initials"
    Write-Host -ForegroundColor red "Password: $password"
}

function multipleUsers
{
    $title = $defaultTitle + " - " + "Multiple user creation"
    Clear-Host
    Write-Host "================ $title ================"

    $hash_table = $null
    $hash_table = @{}
    $amount = Read-Host "How many users do you wish to create?"
    $count = 1..$amount
    foreach ($i in $count)
    {
        $fullname = Read-Host "Insert Full name"
        $fullnameSplit = $fullname.Split()

        $firstname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[0]) # -GivenName
        $surname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[1]) # -Surname

        $initials = $firstname.Substring(0,2) + $surname.Substring(0,2)
        $initials = $initials.ToLower() # -SamAccountName

        $initials = IsSamAccountNameValid($initials)

        $email = $initials + "@" + $defaultEmailDomain # -UserPrincipalName

        Get-OUs
        $ou = Set-OU
        $path = "OU="+$ou+","+$defaultDomainController # -Path
        $password = get-RandomPassword
        $password_secured = ConvertTo-SecureString($password) -AsPlainText -Force # -AccountPassword

        New-ADUser -GivenName $firstname -Surname $surname -Name $initials  -SamAccountName $initials -UserPrincipalName $email -AccountPassword $password_secured -path $path -Enabled $true
        $hash_table.add($initials,$password)
    }

    $results = $hash_table.GetEnumerator() |
        ForEach-Object {
            [PSCustomObject]@{
                Username = $_.Key
                Password = $_.Value
            }
        }

    $results
}
function usersCSV
{
    $title = $defaultTitle + " - " + "CSV user creation"
    Clear-Host
    Write-Host "================ $title ================"
    write-warning "The full path is needed - Example: C:\Users\Documents\users.csv"
    $csvfile = Read-Host "Insert full path to CSV file" | Import-Csv
    $hash_table = $null
    $hash_table = @{}

    foreach ($user in $csvfile)
    {
        $fullname =  $user.fullname

        $fullnameSplit = $fullname.Split()

        $firstname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[0]) # -GivenName
        $surname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[1]) # -Surname

        $initials = $firstname.Substring(0,2) + $surname.Substring(0,2)
        $initials = $initials.ToLower() # -SamAccountName

        $initials = IsSamAccountNameValid($initials)

        $email = $initials + "@" + $defaultEmailDomain # -UserPrincipalName

        $ou = $user.department
        $path = "OU="+$ou+","+$defaultDomainController # -Path
        $password = get-RandomPassword
        $password_secured = ConvertTo-SecureString($password) -AsPlainText -Force # -AccountPassword

        if (IsOUValid($path))
        {
            New-ADUser -GivenName $firstname -Surname $surname -Name $initials  -SamAccountName $initials -UserPrincipalName $email -AccountPassword $password_secured -path $path -Enabled $true
            $hash_table.add($initials,$password)
        }
        else 
        {
            Write-Warning "The OU does not exist for $initials"
            Write-Warning "$initials will not be created!"
        }

        $results = $hash_table.GetEnumerator() |
            ForEach-Object {
                [PSCustomObject]@{
                    Username = $_.Key
                    Password = $_.Value
                }
            }
        
    }
    $results
}
<#
updateUser function
Creates menu to ask what the user needs to get updated
#>
function updateUser
{
    $title = $defaultTitle + " - " + "Update already exisiting user"
    Clear-Host
    Write-Host "================ $title ================"

    $initials = Read-Host "Insert initials for the user to update"
    $check = $true

    if(IsSamAccountNameValidUpdateUser($initials))
    {
        showMenuUpdateUser($initials)
        $selectionMenuUpdateUser = Read-Host $defaultSelectionMsg
        selectionUpdateUser($initials)
    }
    else
    {
        Write-Host "User --- $initials --- does not exist, do you wanna try again?"
        $selectionTryAgain = Read-Host "Press 'y' to try again, press 'q' to quit."
        $tryagain = "updateUser"
        selectionTryAgain($tryagain)
    }
}
<#
updateUserUsername function
Updates the SamAccountName and the email
#>
function updateUserUsername($initials)
{
    $newInitials = Read-Host "Insert the new initials for $initials"
    $email = $newInitials + "@" + $defaultEmailDomain
    try
    {
        Set-ADUser $initials -Replace @{UserPrincipalName="$email"}
        Set-ADUser $initials -Replace @{samaccountname="$newInitials"}
    }
    catch 
    {
        Write-Warning "FAILED - The inserted username already exists"
    }
}

<#
updateUserPassword function
Updates the Password
#>
function updateUserPassword($initials)
{
    $password = get-RandomPassword
    $password_secured = ConvertTo-SecureString($password) -AsPlainText -Force # -AccountPassword

    Set-ADAccountPassword -Identity $initials -Reset -NewPassword $password_secured
    Write-Host "New password for $initials has been set!"
    Write-Host "Password: $password"
}

<#
updateUserDepartment function
Updates the Department/OU
#>
function updateUserDepartment($initials)
{
    Get-OUs
    $ou = Set-OU
    $path = "OU="+$ou+","+$defaultDomainController # -Path
    Get-ADUser $initials | Move-ADObject -TargetPath $path
}

<#
updateUserFullname function
Updates the Fullname
#>
function updateUserFullname($initials)
{
    $fullname = Read-Host "Insert Full name"
    $fullnameSplit = $fullname.Split()

    $firstname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[0]) # -GivenName
    $surname = (Get-Culture).TextInfo.ToTitleCase($fullnameSplit[1]) # -Surname

    Set-ADUser -Identity $initials -GivenName $firstname
    Set-ADUser -Identity $initials -Surname $surname
}

<#
activateUser function
activates the user in the AD specified with the SamAccountName
#>
function activateUser
{
    $title = $defaultTitle + " - " + "Activate user."
    Clear-Host
    Write-Host "================ $title ================"

    $initials = Read-Host "Insert initials for the user to activate"

    try
    {
        Set-ADUser -Identity $initials -Enabled $true
    }
    catch 
    {
        Write-Host "User --- $initials --- does not exist, do you wanna try again?"
        $selectionTryAgain = Read-Host "Press 'y' to try again, press 'q' to quit."
        $tryagain = "activate"
        selectionTryAgain($tryagain)
    }
}
<#
deactivateUser function
activates the user in the AD specified with the SamAccountName
#>
function deactivateUser
{
    $title = $defaultTitle + " - " + "Deactivate user."
    Clear-Host
    Write-Host "================ $title ================"

    $initials = Read-Host "Insert initials for the user to deactivate"

    try
    {
        Set-ADUser -Identity $initials -Enabled $false
    }
    catch 
    {
        Write-Host "User --- $initials --- does not exist, do you wanna try again?"
        $selectionTryAgain = Read-Host "Press 'y' to try again, press 'q' to quit."
        $tryagain = "deactivate"
        selectionTryAgain($tryagain)
    }
}

function test
{
    $title = $defaultTitle + " - " + "test."
    Clear-Host
    Write-Host "================ $title ================"

    $ou = "OU=Test,DC=omit,DC=localhost"

    IsOUValid($ou)
}

<#
Initializes the script
#>
showMenu
$selectionMenu = Read-Host $defaultSelectionMsg
selection
