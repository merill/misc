<#
  .SYNOPSIS
  Adds users from a home tenant to partner tenants as guests without requiring
  invitation redemption by the user.

  This script needs to be run by the user account that has been added to the 
  Guest Inviter role in the partner tenant. For details on setting this up see
  https://docs.microsoft.com/en-us/azure/active-directory/external-identities/add-user-without-invite

  .DESCRIPTION
  As a quick summary the user running this script needs to be
  1. Invited to the partner tenant as a guest
  2. Added to the Guest Inviter role in the partner tenant

  To reduce the time taken to query For optimum performance (especially with large groups), 
  this script uses a delta query and stores the nextLink in the cache folder.

  To run a full query instead of the delta, delete the corresponding group file in the ./cache/ folder.

  .EXAMPLE
  PS> .\Update-Month.ps1

  .EXAMPLE
  PS> .\Update-Month.ps1 -inputpath C:\Data\January.csv

  .EXAMPLE
  PS> .\Update-Month.ps1 -inputpath C:\Data\January.csv -outputPath C:\Reports\2009\January.csv
#>

[CmdletBinding()]
param (
    # The Group ID in the host tenant that contains the users that need to be added to the partner tenant.
    [Parameter(Mandatory = $true)]
    [string] $HomeTenantGroupId,

    # The tenant to which the guests will be added to. 
    [Parameter(Mandatory = $true)]
    [string] $PartnerTenantId
)

function GetGroupCacheFilePath($groupId){
    return ".\cache\delta-query-group-$groupId.txt"
}

# Uses the delta query if a next link
function GetGroupNextLink($groupId){
    $fileName = GetGroupCacheFilePath($groupId)

    if(Test-Path -Path $fileName -PathType Leaf){
        Write-Verbose "Getting delta list of users using cache data at $fileName"
        return Get-Content $fileName
    }
    else {
        Write-Debug "No cache found at $fileName"
        return $null
    }
}

function SaveGroupNextLink($groupId, $nextLink){
    if($null -eq $nextLink -or $nextLink.Length -eq 0){
        Write-Verbose "Invalid nextLink, not saving to cache"
    }
    else{
        $fileName = GetGroupCacheFilePath($groupId)
        Write-Verbose "Saving nextLink at $fileName"
    
        New-Item -Path $fileName -ItemType File -Value $nextLink -Force | Out-Null    
    }
}

function GetGuestUsers($groupId){
    Write-Verbose "`n`nGetting list of users from group $groupId ---------------"
    $graphQuery = GetGroupNextLink($groupId)

    if($null -eq $graphQuery){ # first time run
        $graphQuery = "https://graph.microsoft.com/v1.0/groups/delta?`$filter=+id+eq+'{0}'" -f $groupId
        Write-Verbose "Getting full list of users with`n$graphQuery"
    }

    $members = @()
    $hasData = $true
    
    while($hasData){
        $graphResult = Invoke-GraphRequest -Method GET -Uri $graphQuery

        if($null -ne $graphResult.Value -and $graphResult.Value.'members@delta'.Length -gt 0){
            $members += $graphResult.Value.'members@delta'
            Write-Verbose "Found $($members.Length) members, checking for more guests"
        }
        $hasData = $null -ne $graphResult.'@odata.nextLink' # is there another page of data?
        if($hasData){
            $graphQuery = $graphResult.'@odata.nextLink'
        }
    }
    
    $deltaLink = $graphResult.'@odata.deltaLink' # Get the delta link for the  next query
    Write-Debug "DeltaLink"
    Write-Debug $deltaLink
    return $deltaLink, $members
}

function GetUserDetails($members){
    Write-Verbose "Getting user information for $($members.Length) guests"
    $usersToAdd = @()
    $usersToRemove = @()
    foreach ($member in $members) {
        if($null -ne $member.'@removed'){
            $userProps = @{
                Id = $member.id
            }
            $usersToRemove += $userProps
        }
        else{
            Write-Verbose "Getting user information for user $($member.Id)"
            $user = Get-MgUser -UserId $member.Id -Property Id, UserPrincipalName, Mail, DisplayName, GivenName, Surname
            $userProps = [ordered]@{
                Id = $user.Id
                UserPrincipalName = $user.UserPrincipalName
                Mail = $user.Mail
                DisplayName = $user.DisplayName
                GivenName = $user.GivenName
                Surname = $user.Surname
            }
            $usersToAdd += $userProps
        }
    }
    return $usersToAdd, $usersToRemove
}

function InviteUsers($usersToAdd, $inviteRedirectUrl){
    Write-Information "`n`nInviting $($usersToAdd.Length) users---------------"
    foreach($user in $usersToAdd){
        Write-Verbose "Inviting user $($user.Mail)"
        New-MgInvitation `
            -InvitedUserEmailAddress $user.Mail `
            -InvitedUserDisplayName $user.DisplayName `
            -InviteRedirectUrl $inviteRedirectUrl `
            -SendInvitationMessage:$false | Out-Null
    }
}

$ErrorActionPreference = "Stop"

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Authentication

Write-Host "Connecting to Home Tenant to get list of users to invite as guest"
Connect-MgGraph -Scopes GroupMember.Read.All | Out-Null # Connect to the logged in user's home tenant
$logFolder = ".\Logs\" + $((Get-Date).ToString("yyyyMMdd-HHmmss"))
New-Item -Path $logFolder -ItemType Directory | Out-Null

Write-Host "Getting list of guests to invite"
$nextLink, $members = GetGuestUsers -groupId $HomeTenantGroupId
$usersToAdd, $usersToRemove = GetUserDetails -members $members

$userCount = 0
if($null -ne $usersToAdd) { $userCount = $usersToAdd.Length }
Write-Host "Found $userCount users to invite"

Write-Host "Exporting log of guests to be invited to $logFolder"
$usersToAdd | Export-Csv -Path (Join-Path $logFolder -ChildPath "UsersAdd.csv")
$usersToRemove | Export-Csv -Path (Join-Path $logFolder -ChildPath "UsersRemove.csv")

if($usersToAdd.Length -gt 0){
    Write-Host "Connecting to Partner Tenant $PartnerTenantId to perform invites"
    Connect-MgGraph -TenantId $PartnerTenantId -Scopes User.Invite.All | Out-Null
    $inviteRedirectUrl = "https://myapps.microsoft.com?tenantId=$PartnerTenantId"
    
    Write-Host "Inviting users"
    InviteUsers -usersToAdd $usersToAdd -inviteRedirectUrl $inviteRedirectUrl
}

Write-Host "Saving cache"
SaveGroupNextLink -groupId $HomeTenantGroupId -nextLink $nextLink