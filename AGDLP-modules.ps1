function Add-OU {
    param(
        [Parameter(Mandatory=$true)]
        [string]$name,
        [Parameter(Mandatory=$true)]
        [string]$path
    )
    try {
        New-ADOrganizationalUnit -Path $path -Name $name
        Write-Verbose "Created OU named '$name'"
    } catch {
        if($_ -match "name that is already in use") {Write-Verbose "OU named '$name' already exists"}
        else {Write-Error "An error occured while trying to create OU named '$name' -> $_"}
    }
}


function Add-GG {
    param (
        [Parameter(Mandatory=$true)]
        [string]$name,
        [Parameter(Mandatory=$true)]
        [string]$path
    ) 
    try {
        New-ADGroup -Path $path -Name "GG - $name" -GroupCategory Security -GroupScope Global
        Write-Verbose "Created GG named '$name'"
    } catch {
        if($_ -match "already exists") {Write-Verbose "GG named '$name' already exists"}
        else {Write-Error "An error occured while trying to create GG named '$name' -> $_"}
    }
}


function Add-GL {
    param (
        [Parameter(Mandatory=$true)]
        [string]$name,
        [Parameter(Mandatory=$true)]
        [string]$path
    ) 
    try {
        New-ADGroup -Path $path -Name "GL - $name - R" -GroupCategory Security -GroupScope DomainLocal
        New-ADGroup -Path $path -Name "GL - $name - RW" -GroupCategory Security -GroupScope DomainLocal
        Write-Verbose "Created GL named '$name'"
    } catch {
        if($_ -match "already exists") {Write-Verbose "GL named '$name' already exists"}
        else {Write-Error "An error occured while trying to create GL named '$name' -> $_"}
    }
}


function Add-UPNSuffix {
    param (
        [Parameter(Mandatory=$true)]
        [string]$name,
        [Parameter(Mandatory=$true)]
        [string]$path
    )
    try {
        Get-ADForest | Set-ADForest -UPNSuffixes @{Add="$name"}
        Write-Verbose "Created UPN Suffix named '$name'"
    }
    catch {
        if($_ -match "already exists") {Write-Verbose "UPN Suffix named '$name' already exists"}
        else {Write-Error "An error occured while trying to create UPN Suffix named '$name' -> $_"}
    }
    
}

function Add-Folder {
    param (
        [Parameter(Mandatory=$true)]
        [string]$name,
        [Parameter(Mandatory=$true)]
        [string]$path
    ) 
    try {
        if (-not (Test-Path -Path "$path\$name")) {
        New-Item -Path "$path\$name" -ItemType Directory
        Write-Verbose "Created folder named '$name'"
        } else {Write-Verbose "Folder named '$name' already exists"}
    } catch {Write-Error "An error occured while trying to create folder named '$name' -> $_"}
}


function Add-User {
    param (
        [Parameter(Mandatory=$true)]
        [string]$id,
        [Parameter(Mandatory=$true)]
        [string]$firstname,
        [Parameter(Mandatory=$true)]
        [string]$lastname,        
        [Parameter(Mandatory=$true)]
        [string]$password,
        [Parameter(Mandatory=$true)]
        [string]$office,
        [Parameter(Mandatory=$true)]
        [string]$description,
        [Parameter(Mandatory=$true)]
        [string]$number,
        [Parameter(Mandatory=$true)]
        [string]$suffix,
        [Parameter(Mandatory=$true)]
        [string]$group,
        [Parameter(Mandatory=$true)]
        [string]$path
    )
    try {
        $userDetails = @{
            Enabled = $true
            AccountPassword = (ConvertTo-SecureString $password -AsPlainText -Force)
            Name = $id
            SamAccountName = $id
            GivenName = $firstname
            Surname = $lastname
            UserPrincipalName = "$id@$suffix"
            Description = $description
            Office = $office
            Path = $path
        }
        New-ADUser @userDetails -OtherAttributes @{'ipPhone' = $number}
        # Add to group
        Add-ADGroupMember -Identity "GG - $group" -Members $id
        Write-Verbose "Created user named '$name'"
    }
    catch {
        if($_ -match "already exists") {Write-Verbose "User '$id' already exists."}
        else {Write-Error "An error occured while trying to create user named '$($id)' -> $_"}
    }
}