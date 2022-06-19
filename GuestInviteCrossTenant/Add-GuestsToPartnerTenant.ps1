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
    $HomeTenantGroupId,
    # The tenant to which the guests will be added to. 
    $PartnerTenantId
)

# Uses the delta query if a next link
function GetGroupNextLink($groupId){
    $fileName = ".\cache\delta-query-group-$groupId.txt";
    if(Test-Path -Path $fileName -PathType Leaf){
        return Get-Content $fileName
    }
    else {
        return $null
    }
}

function GetGuestUsers($groupId){
    $graphQuery = GetGroupNextLink($groupId)

    if($null -eq $graphQuery){ # first time run
        $graphQuery = "https://graph.microsoft.com/v1.0/groups/delta?`$filter=+id+eq+'{0}'" -f $groupId
    }

    $graphResult = Invoke-GraphRequest -Method GET -Uri $graphQuery
    $members = @()
    while($graphResult.Value.Length -gt 0){
        $members += $graphResult.Value.'members@delta'

        $graphQuery = $graphResult.'@odata.nextLink'
        $graphResult = Invoke-GraphRequest -Method GET -Uri $graphQuery
    }
    return $graphResult.'@odata.nextLink', $members
}

function GetUserDetails($members){
    $usersToAdd = @()
    $usersToRemove = @()
    foreach ($member in $members) {
        if($member.'@removed'.reason -eq 'deleted'){
            $userProps = @{
                Id = $member.id
            }
            $usersToRemove += $userProps
        }
        else{
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

Import-Module Microsoft.Graph.Users
Import-Module Microsoft.Graph.Authentication

Connect-MgGraph -Scopes GroupMember.Read.All # Connect to the logged in user's home tenant
$logFolder = ".\Logs\" + $((Get-Date).ToString("yyyyMMdd-HHmmss"))
New-Item -Path $logFolder -ItemType Directory

$nextLink, $members = GetGuestUsers($HomeTenantGroupId)
$usersToAdd, $usersToRemove = GetUserDetails($members)

$usersToAdd | Export-Csv -Path (Join-Path $logFolder -ChildPath "UsersAdd.csv")
$usersToRemove | Export-Csv -Path (Join-Path $logFolder -ChildPath "UsersRemove.csv")

Connect-MgGraph -TenantId $PartnerTenantId -Scopes User.Invite.All
$InviteRedirectUrl = "https://myapps.microsoft.com?tenantId=$PartnerTenantId"

foreach($user in $usersToAdd){
    New-MgInvitation `
        -InvitedUserEmailAddress $user.Mail `
        -InvitedUserDisplayName $user.DisplayName `
        -InviteRedirectUrl $InviteRedirectUrl `
        -SendInvitationMessage:$false `
}

