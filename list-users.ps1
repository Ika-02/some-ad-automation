Import-Module ActiveDirectory

function Get-Users {
    param (
        [Parameter(Mandatory=$true)]
        [string]$ou,
        [string]$domain = "france.lan"
    )
    $dn = $domain -split "\."
    # Get all the users in the OU and display their full name, login and group in a table
    Get-ADUser -filter * -searchbase "OU=$ou,DC=$($dn[0]),DC=$($dn[1])" -properties * | `
    Sort-Object -Property SN | `
    Format-Table @{Name = "Full Name"; Expression = {"$($_.SN) $($_.GivenName)"}}, `
    @{Name = "Logon Name"; Expression = {$_.userPrincipalName}}, `
    @{Name = "Group"; Expression = {(($_.MemberOf -split ",")[0]) -replace "CN=GG - ", ""}} -A
}

Get-Users -ou "ressources-humaines"