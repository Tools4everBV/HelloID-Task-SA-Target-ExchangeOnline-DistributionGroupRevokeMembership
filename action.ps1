# HelloID-Task-SA-Target-ExchangeOnline-DistributionGroupRevokeMembership
#########################################################################
# Form mapping
$formObject = @{
    GroupIdentity = $form.GroupIdentity
    UsersToRemove = $form.UsersToRemove.Name
}

[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnline action: [DistributionGroupRevokeMembership] for: [$($formObject.GroupIdentity)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Remove-DistributionGroupMember', 'Disconnect-ExchangeOnline'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToRemove) {
        try {
            $removeDistributionGroupMember = @{
                Identity                        = $formObject.GroupIdentity
                Member                          = $user
                BypassSecurityGroupManagerCheck = $true
            }
            $null = Remove-DistributionGroupMember @removeDistributionGroupMember -Confirm:$false -ErrorAction Stop
            $auditLog = @{
                Action            = 'RevokeMembership'
                System            = 'ExchangeOnline'
                TargetIdentifier  = $formObject.GroupIdentity
                TargetDisplayName = $formObject.GroupIdentity
                Message           = "ExchangeOnline action: [DistributionGroupRevokeMembership][$($user)] from group [$($formObject.GroupIdentity)] executed successfully"
                IsError           = $false
            }
            Write-Information -Tags 'Audit' -MessageData $auditLog
            Write-Information "ExchangeOnline action: [DistributionGroupRevokeMembership][$($user)] from group [$($formObject.GroupIdentity)] executed successfully"

        } catch {
            if ($_.Exception.ErrorRecord.CategoryInfo.Reason -eq 'MemberNotFoundException') {
                $auditLog = @{
                    Action            = 'RevokeMembership'
                    System            = 'ExchangeOnline'
                    TargetIdentifier  = $formObject.GroupIdentity
                    TargetDisplayName = $formObject.GroupIdentity
                    Message           = "ExchangeOnline action: [DistributionGroupRevokeMembership][$($user)] from group [$($formObject.GroupIdentity)] Already Removed"
                    IsError           = $false
                }
                Write-Information -Tags 'Audit' -MessageData $auditLog
                Write-Information "ExchangeOnline action: [DistributionGroupRevokeMembership][$($user)] from group [$($formObject.GroupIdentity)] Already Removed"
                Write-Information "Warning message: $($_.Exception.Message)"
                continue
            }
            throw $_
        }
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'RevokeMembership'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.Identity
        TargetDisplayName = $formObject.Identity
        Message           = "Could not execute ExchangeOnline action: [DistributionGroupRevokeMembership] for: [$($formObject.Identity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [DistributionGroupRevokeMembership] for: [$($formObject.Identity)], error: $($ex.Exception.Message)"

} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}

#########################################################################
