# MailboxToolsTabbedGUI.ps1
# Exchange Online Mailbox Troubleshooting GUI
#
# Tabs:
#   - Mailbox Summary        (auto on lookup)
#   - Mailbox Diagnostics    (auto on lookup)
#   - Inbox Rules            (auto on lookup)
#   - Permissions Summary    (auto on lookup: Full Access, Send As, Send on Behalf,
#                             Inbox / Calendar / Deleted Items folder permissions)
#   - Forwarding Check       (auto on lookup: red/yellow/green status rows,
#                             plus on-demand transport rule scan)
#   - Executive Calendar     (auto on lookup: delegates, calendar processing)
#   - Distribution Groups    (on-demand button - group memberships)
#   - Send Rights Analyzer   (own input box - built for shared mailboxes)
#   - Audit Investigation    (date range + Search-UnifiedAuditLog buttons)
#   - Message Trace          (Get-MessageTraceV2/Get-MessageTrace, double-click a row for detail)
#
# Export:
#   - "Export Tab" button    exports the tab you are currently looking at
#   - "Export..." button     opens a picker: choose which tabs, all rows vs selected rows,
#                            separate CSV per tab or one combined CSV with a Section column
#
# Run with:
# powershell.exe -STA -ExecutionPolicy Bypass -File .\MailboxToolsTabbedGUI.ps1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# =============================
# Global Variables
# =============================

$script:ConnectedToEXO = $false
$script:ResolvedMailbox = $null
$script:DefaultAdminUPN = "TDPORD8A@prod.mtb.com"
$script:ConnectedAdminUPN = ""
$script:LoadedEXOModuleVersion = ""
$script:LoadedEXOModulePath = ""
$script:TraceCmdletUsed = ""

# =============================
# Helper Functions
# =============================

function New-InfoRow {
    param(
        [string]$Message
    )

    return @(
        [PSCustomObject]@{
            Message = $Message
        }
    )
}

function Format-ToolValue {
    param(
        [object]$Value
    )

    if ($null -eq $Value) {
        return ""
    }

    if ($Value -is [array]) {
        return ($Value -join "; ")
    }

    return $Value.ToString()
}

function ConvertTo-DataTable {
    param(
        [Parameter(Mandatory)]
        [object[]]$InputObject
    )

    $dataTable = New-Object System.Data.DataTable

    if (-not $InputObject -or $InputObject.Count -eq 0) {
        return ,$dataTable
    }

    $properties = $InputObject[0].PSObject.Properties.Name

    foreach ($property in $properties) {
        [void]$dataTable.Columns.Add($property)
    }

    foreach ($item in $InputObject) {
        $row = $dataTable.NewRow()

        foreach ($property in $properties) {
            $value = $item.$property

            if ($null -eq $value) {
                $row[$property] = ""
            }
            elseif ($value -is [array]) {
                $row[$property] = ($value -join "; ")
            }
            else {
                $row[$property] = $value.ToString()
            }
        }

        [void]$dataTable.Rows.Add($row)
    }

    return ,$dataTable
}

# =============================
# Import Exchange Online Module
# =============================

function Import-EXOModuleTool {
    try {
        # Direct search locations.
        # This includes the PowerShell 7 module path where your module is currently installed.
        $possibleRoots = @(
            "$env:USERPROFILE\OneDrive - M&T Bank\Documents\PowerShell\Modules\ExchangeOnlineManagement",
            "$env:USERPROFILE\OneDrive - M&T Bank\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement",
            "$env:USERPROFILE\Documents\PowerShell\Modules\ExchangeOnlineManagement",
            "$env:USERPROFILE\Documents\WindowsPowerShell\Modules\ExchangeOnlineManagement",
            "C:\Program Files\PowerShell\Modules\ExchangeOnlineManagement",
            "C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement"
        )

        $moduleCandidates = New-Object System.Collections.Generic.List[object]

        foreach ($root in $possibleRoots) {
            if (Test-Path $root) {
                $foundManifests = Get-ChildItem -Path $root -Recurse -Filter "ExchangeOnlineManagement.psd1" -ErrorAction SilentlyContinue

                foreach ($manifest in $foundManifests) {
                    try {
                        $manifestData = Test-ModuleManifest -Path $manifest.FullName -ErrorAction Stop

                        $moduleCandidates.Add([PSCustomObject]@{
                            Version = $manifestData.Version
                            Path    = $manifest.FullName
                        })
                    }
                    catch {
                        # Ignore bad manifest reads and keep searching.
                    }
                }
            }
        }

        if ($moduleCandidates.Count -eq 0) {
            $message = @"
ExchangeOnlineManagement module was not found by direct path search.

Checked locations:

$($possibleRoots -join "`r`n")

Your module may be installed somewhere else, or this script may be running under a different Windows user profile.

Current USERPROFILE:
$env:USERPROFILE

Current PowerShell module path:
$($env:PSModulePath -replace ';', "`r`n")
"@

            [System.Windows.Forms.MessageBox]::Show(
                $message,
                "Module Not Found",
                "OK",
                "Error"
            )

            return $false
        }

        $selectedModule = $moduleCandidates |
            Sort-Object Version -Descending |
            Select-Object -First 1

        Import-Module $selectedModule.Path -Force -ErrorAction Stop

        $script:LoadedEXOModuleVersion = $selectedModule.Version.ToString()
        $script:LoadedEXOModulePath = $selectedModule.Path

        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Unable to import ExchangeOnlineManagement.`r`n`r`n$($_.Exception.Message)",
            "Module Import Error",
            "OK",
            "Error"
        )

        return $false
    }
}

# =============================
# Connect to Exchange Online
# =============================

function Connect-EXOTool {
    try {
        $moduleLoaded = Import-EXOModuleTool

        if (-not $moduleLoaded) {
            return $false
        }

        $adminUPN = [Microsoft.VisualBasic.Interaction]::InputBox(
            "Enter the Exchange admin account to connect with:",
            "Connect to Exchange Online",
            $script:DefaultAdminUPN
        )

        if ([string]::IsNullOrWhiteSpace($adminUPN)) {
            [System.Windows.Forms.MessageBox]::Show(
                "No admin account was entered. The tool will not connect.",
                "Connection Cancelled",
                "OK",
                "Warning"
            )

            return $false
        }

        Connect-ExchangeOnline `
            -UserPrincipalName $adminUPN `
            -ShowProgress $true `
            -ShowBanner:$false `
            -ErrorAction Stop

        $script:ConnectedAdminUPN = $adminUPN

        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Exchange Online connection failed.`r`n`r`n$($_.Exception.Message)",
            "Connection Error",
            "OK",
            "Error"
        )

        return $false
    }
}

# =============================
# Resolve Mailbox
# =============================

function Resolve-MailboxIdentity {
    param(
        [Parameter(Mandatory)]
        [string]$InputIdentity
    )

    $exoProperties = @(
        "DisplayName",
        "Alias",
        "PrimarySmtpAddress",
        "UserPrincipalName",
        "RecipientTypeDetails",
        "ForwardingSmtpAddress",
        "ForwardingAddress",
        "DeliverToMailboxAndForward",
        "HiddenFromAddressListsEnabled",
        "GrantSendOnBehalfTo",
        "LitigationHoldEnabled",
        "ArchiveStatus",
        "ArchiveDatabase",
        "EmailAddresses",
        "MessageCopyForSentAsEnabled",
        "MessageCopyForSendOnBehalfEnabled",
        "DistinguishedName"
    )

    try {
        $mailbox = Get-EXOMailbox `
            -Identity $InputIdentity `
            -Properties $exoProperties `
            -ErrorAction Stop

        return $mailbox
    }
    catch {
        try {
            $mailboxMatches = @(Get-EXOMailbox `
                -Anr $InputIdentity `
                -ResultSize 10 `
                -Properties $exoProperties `
                -ErrorAction Stop)

            if ($mailboxMatches.Count -eq 0) {
                throw "No mailbox found for '$InputIdentity'. Try the full email address."
            }

            if ($mailboxMatches.Count -gt 1) {
                $matchList = ($mailboxMatches | Select-Object DisplayName, PrimarySmtpAddress | Out-String)

                throw "More than one mailbox matched '$InputIdentity'. Please use the full email address.`r`n`r`nMatches:`r`n$matchList"
            }

            return $mailboxMatches[0]
        }
        catch {
            throw "Could not resolve mailbox '$InputIdentity'. Try the full email address instead.`r`n`r`n$($_.Exception.Message)"
        }
    }
}

# =============================
# Mailbox Summary
# =============================

function Get-MailboxSummaryTool {
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $identity = $Mailbox.PrimarySmtpAddress.ToString()

    try {
        $stats = Get-EXOMailboxStatistics -Identity $identity -ErrorAction Stop
    }
    catch {
        $stats = $null
    }

    return @(
        [PSCustomObject]@{
            DisplayName                = Format-ToolValue $Mailbox.DisplayName
            Alias                      = Format-ToolValue $Mailbox.Alias
            UserPrincipalName          = Format-ToolValue $Mailbox.UserPrincipalName
            PrimarySmtpAddress         = Format-ToolValue $Mailbox.PrimarySmtpAddress
            RecipientTypeDetails       = Format-ToolValue $Mailbox.RecipientTypeDetails
            HiddenFromAddressLists     = Format-ToolValue $Mailbox.HiddenFromAddressListsEnabled
            LitigationHoldEnabled      = Format-ToolValue $Mailbox.LitigationHoldEnabled
            ArchiveStatus              = Format-ToolValue $Mailbox.ArchiveStatus
            ForwardingSmtpAddress      = Format-ToolValue $Mailbox.ForwardingSmtpAddress
            ForwardingAddress          = Format-ToolValue $Mailbox.ForwardingAddress
            DeliverToMailboxAndForward = Format-ToolValue $Mailbox.DeliverToMailboxAndForward
            TotalItemSize              = if ($stats) { Format-ToolValue $stats.TotalItemSize } else { "" }
            ItemCount                  = if ($stats) { Format-ToolValue $stats.ItemCount } else { "" }
            DeletedItemCount           = if ($stats) { Format-ToolValue $stats.DeletedItemCount } else { "" }
            LastLogonTime              = if ($stats) { Format-ToolValue $stats.LastLogonTime } else { "" }
        }
    )
}

# =============================
# Inbox Rules
# =============================

function Get-InboxRulesTool {
    param(
        [Parameter(Mandatory)]
        [string]$MailboxIdentity
    )

    try {
        $rules = @(Get-InboxRule -Mailbox $MailboxIdentity -IncludeHidden -ErrorAction Stop)

        if ($rules.Count -eq 0) {
            return New-InfoRow -Message "No inbox rules found."
        }

        return $rules | Select-Object `
            Name,
            Enabled,
            Priority,
            @{
                Name = "ForwardTo"
                Expression = { Format-ToolValue $_.ForwardTo }
            },
            @{
                Name = "RedirectTo"
                Expression = { Format-ToolValue $_.RedirectTo }
            },
            @{
                Name = "ForwardAsAttachmentTo"
                Expression = { Format-ToolValue $_.ForwardAsAttachmentTo }
            },
            @{
                Name = "DeleteMessage"
                Expression = { Format-ToolValue $_.DeleteMessage }
            },
            @{
                Name = "MoveToFolder"
                Expression = { Format-ToolValue $_.MoveToFolder }
            },
            @{
                Name = "StopProcessingRules"
                Expression = { Format-ToolValue $_.StopProcessingRules }
            },
            @{
                Name = "RuleDescription"
                Expression = { Format-ToolValue $_.Description }
            }
    }
    catch {
        return New-InfoRow -Message "Error pulling inbox rules: $($_.Exception.Message)"
    }
}

# =============================
# Folder Permissions
# =============================

function Get-FolderPermissionTool {
    param(
        [Parameter(Mandatory)]
        [string]$MailboxIdentity,

        [Parameter(Mandatory)]
        [string]$FolderName
    )

    try {
        $folderIdentity = "${MailboxIdentity}:\$FolderName"

        $perms = @(Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop)

        if ($perms.Count -eq 0) {
            return New-InfoRow -Message "No permissions found on $FolderName."
        }

        return $perms | Select-Object `
            @{
                Name = "Folder"
                Expression = { $FolderName }
            },
            @{
                Name = "User"
                Expression = { $_.User.ToString() }
            },
            @{
                Name = "AccessRights"
                Expression = { ($_.AccessRights -join ", ") }
            },
            @{
                Name = "SharingPermissionFlags"
                Expression = {
                    if ($_.PSObject.Properties.Name -contains "SharingPermissionFlags") {
                        Format-ToolValue $_.SharingPermissionFlags
                    }
                    else {
                        ""
                    }
                }
            },
            @{
                Name = "IsDefaultOrAnonymous"
                Expression = {
                    if ($_.User.ToString() -in @("Default", "Anonymous")) {
                        "Yes"
                    }
                    else {
                        "No"
                    }
                }
            }
    }
    catch {
        return New-InfoRow -Message "Error pulling $FolderName permissions: $($_.Exception.Message)"
    }
}

# =============================
# Mailbox Delegation
# =============================

function Get-MailboxDelegationTool {
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $mailboxIdentity = $Mailbox.PrimarySmtpAddress.ToString()
    $results = New-Object System.Collections.Generic.List[object]

    # Full Access
    try {
        try {
            $fullAccessPermissions = @(Get-EXOMailboxPermission -Identity $mailboxIdentity -ErrorAction Stop)
            $fullAccessSource = "Get-EXOMailboxPermission"
        }
        catch {
            $fullAccessPermissions = @(Get-MailboxPermission -Identity $mailboxIdentity -ErrorAction Stop)
            $fullAccessSource = "Get-MailboxPermission"
        }

        $fullAccessPermissions = $fullAccessPermissions | Where-Object {
            $_.IsInherited -eq $false -and
            $_.Deny -eq $false -and
            $_.User -notmatch "NT AUTHORITY\\SELF" -and
            $_.User -notmatch "S-1-5-" -and
            $_.AccessRights -contains "FullAccess"
        }

        foreach ($permission in $fullAccessPermissions) {
            $results.Add([PSCustomObject]@{
                PermissionType = "Full Access"
                Delegate       = $permission.User.ToString()
                AccessRights   = ($permission.AccessRights -join ", ")
                Source         = $fullAccessSource
                Notes          = "Can open and read the mailbox. Does not grant Send As by itself."
            })
        }
    }
    catch {
        $results.Add([PSCustomObject]@{
            PermissionType = "Full Access"
            Delegate       = "Error"
            AccessRights   = ""
            Source         = "Mailbox permissions"
            Notes          = $_.Exception.Message
        })
    }

    # Send As
    try {
        try {
            $sendAsPermissions = @(Get-EXORecipientPermission -Identity $mailboxIdentity -ErrorAction Stop)
            $sendAsSource = "Get-EXORecipientPermission"
        }
        catch {
            $sendAsPermissions = @(Get-RecipientPermission -Identity $mailboxIdentity -ErrorAction Stop)
            $sendAsSource = "Get-RecipientPermission"
        }

        $sendAsPermissions = $sendAsPermissions | Where-Object {
            $_.IsInherited -eq $false -and
            $_.Deny -eq $false -and
            $_.Trustee -notmatch "NT AUTHORITY\\SELF" -and
            $_.Trustee -notmatch "S-1-5-" -and
            $_.AccessRights -contains "SendAs"
        }

        foreach ($permission in $sendAsPermissions) {
            $results.Add([PSCustomObject]@{
                PermissionType = "Send As"
                Delegate       = $permission.Trustee.ToString()
                AccessRights   = ($permission.AccessRights -join ", ")
                Source         = $sendAsSource
                Notes          = "Can send as the mailbox."
            })
        }
    }
    catch {
        $results.Add([PSCustomObject]@{
            PermissionType = "Send As"
            Delegate       = "Error"
            AccessRights   = ""
            Source         = "Recipient permissions"
            Notes          = $_.Exception.Message
        })
    }

    # Send on Behalf
    try {
        $sendOnBehalfDelegates = @($Mailbox.GrantSendOnBehalfTo)

        if ($sendOnBehalfDelegates.Count -gt 0 -and $null -ne $sendOnBehalfDelegates[0]) {
            foreach ($delegate in $sendOnBehalfDelegates) {
                $results.Add([PSCustomObject]@{
                    PermissionType = "Send on Behalf"
                    Delegate       = $delegate.ToString()
                    AccessRights   = "GrantSendOnBehalfTo"
                    Source         = "Get-EXOMailbox"
                    Notes          = "Can send on behalf of the mailbox."
                })
            }
        }
    }
    catch {
        $results.Add([PSCustomObject]@{
            PermissionType = "Send on Behalf"
            Delegate       = "Error"
            AccessRights   = ""
            Source         = "Get-EXOMailbox"
            Notes          = $_.Exception.Message
        })
    }

    return $results
}

# =============================
# Permissions Summary (everything in one view)
# =============================

function Get-PermissionsSummaryTool {
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $mailboxIdentity = $Mailbox.PrimarySmtpAddress.ToString()
    $rows = New-Object System.Collections.Generic.List[object]

    # Mailbox-level: Full Access, Send As, Send on Behalf
    $delegation = @(Get-MailboxDelegationTool -Mailbox $Mailbox)

    foreach ($entry in $delegation) {
        $rows.Add([PSCustomObject]@{
            Area         = "Mailbox"
            Type         = $entry.PermissionType
            User         = $entry.Delegate
            AccessRights = $entry.AccessRights
            Details      = $entry.Notes
        })
    }

    if ($delegation.Count -eq 0) {
        $rows.Add([PSCustomObject]@{
            Area         = "Mailbox"
            Type         = "Info"
            User         = "(none)"
            AccessRights = ""
            Details      = "No mailbox-level delegation (Full Access / Send As / Send on Behalf)."
        })
    }

    # Folder-level: Inbox, Calendar, Deleted Items
    foreach ($folderName in @("Inbox", "Calendar", "Deleted Items")) {
        $folderPerms = @(Get-FolderPermissionTool -MailboxIdentity $mailboxIdentity -FolderName $folderName)

        foreach ($perm in $folderPerms) {
            if ($perm.PSObject.Properties.Name -contains "Message") {
                $rows.Add([PSCustomObject]@{
                    Area         = "Folder: $folderName"
                    Type         = "Info"
                    User         = ""
                    AccessRights = ""
                    Details      = $perm.Message
                })
            }
            else {
                $flagNote = ""

                if ($perm.SharingPermissionFlags) {
                    $flagNote = "SharingPermissionFlags: $($perm.SharingPermissionFlags)"
                }

                $rows.Add([PSCustomObject]@{
                    Area         = "Folder: $folderName"
                    Type         = "Folder Permission"
                    User         = $perm.User
                    AccessRights = $perm.AccessRights
                    Details      = $flagNote
                })
            }
        }
    }

    if ($rows.Count -eq 0) {
        return New-InfoRow -Message "No permissions found."
    }

    return $rows
}

# =============================
# Forwarding / Compromise Check
# =============================

function Get-ForwardingCheckTool {
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $mailboxIdentity = $Mailbox.PrimarySmtpAddress.ToString()
    $rows = New-Object System.Collections.Generic.List[object]

    function Add-CheckRow {
        param(
            [string]$Check,
            [string]$Status,
            [string]$Details
        )

        $rows.Add([PSCustomObject]@{
            Check   = $Check
            Status  = $Status
            Details = $Details
        })
    }

    # 1. Mailbox-level external SMTP forwarding
    if ($Mailbox.ForwardingSmtpAddress) {
        Add-CheckRow "Forwarding SMTP Address" "ALERT" "Set to: $($Mailbox.ForwardingSmtpAddress). Common indicator in BEC / compromise cases."
    }
    else {
        Add-CheckRow "Forwarding SMTP Address" "OK" "Not set."
    }

    # 2. Mailbox-level internal forwarding
    if ($Mailbox.ForwardingAddress) {
        Add-CheckRow "Forwarding Address (internal)" "WARN" "Set to: $($Mailbox.ForwardingAddress). Verify this is intentional."
    }
    else {
        Add-CheckRow "Forwarding Address (internal)" "OK" "Not set."
    }

    # 3. Deliver-and-forward flag
    if ($Mailbox.DeliverToMailboxAndForward) {
        Add-CheckRow "Deliver To Mailbox And Forward" "WARN" "Enabled. Mail is delivered AND forwarded."
    }
    else {
        Add-CheckRow "Deliver To Mailbox And Forward" "OK" "Disabled."
    }

    # 4 - 6. Inbox rule checks
    try {
        $visibleRules = @(Get-InboxRule -Mailbox $mailboxIdentity -ErrorAction Stop)
        $allRules = @(Get-InboxRule -Mailbox $mailboxIdentity -IncludeHidden -ErrorAction Stop)

        # Rules that forward or redirect
        $forwardRules = @($allRules | Where-Object {
            $_.ForwardTo -or
            $_.RedirectTo -or
            $_.ForwardAsAttachmentTo
        })

        if ($forwardRules.Count -gt 0) {
            $names = ($forwardRules | ForEach-Object { $_.Name }) -join "; "
            Add-CheckRow "Inbox Rules - Forward/Redirect" "ALERT" "$($forwardRules.Count) rule(s) forward or redirect mail: $names"
        }
        else {
            Add-CheckRow "Inbox Rules - Forward/Redirect" "OK" "No forwarding or redirecting rules."
        }

        # Rules that delete messages (a common way to hide replies)
        $deleteRules = @($allRules | Where-Object { $_.DeleteMessage -eq $true })

        if ($deleteRules.Count -gt 0) {
            $names = ($deleteRules | ForEach-Object { $_.Name }) -join "; "
            Add-CheckRow "Inbox Rules - Delete Message" "WARN" "$($deleteRules.Count) rule(s) delete messages: $names. Verify these are intentional."
        }
        else {
            Add-CheckRow "Inbox Rules - Delete Message" "OK" "No message-deleting rules."
        }

        # Hidden rules: only visible with -IncludeHidden
        $visibleNames = @($visibleRules | ForEach-Object { $_.Name })
        $hiddenOnly = @($allRules | Where-Object { $_.Name -notin $visibleNames })
        $suspiciousHidden = @($hiddenOnly | Where-Object { $_.Name -ne "Junk E-mail Rule" })

        if ($suspiciousHidden.Count -gt 0) {
            $names = ($suspiciousHidden | ForEach-Object { $_.Name }) -join "; "
            Add-CheckRow "Hidden Rules" "ALERT" "$($suspiciousHidden.Count) hidden rule(s) beyond the standard Junk E-mail Rule: $names"
        }
        else {
            Add-CheckRow "Hidden Rules" "OK" "Only the standard hidden Junk E-mail Rule (or none). Note: rules tampered with at the MAPI level do not appear here at all."
        }

        Add-CheckRow "Inbox Rules - Total" "INFO" "$($allRules.Count) rule(s) total including hidden. $(@($allRules | Where-Object { $_.Enabled -eq $true }).Count) enabled."
    }
    catch {
        Add-CheckRow "Inbox Rule Checks" "WARN" "Error pulling inbox rules: $($_.Exception.Message)"
    }

    Add-CheckRow "Transport Rules" "INFO" "Use the 'Check Transport Rules' button above to scan org transport rules for this address (can be slow in large orgs)."

    return $rows
}

function Get-TransportRuleCheckTool {
    param(
        [Parameter(Mandatory)]
        [string]$SmtpAddress
    )

    $rows = New-Object System.Collections.Generic.List[object]

    try {
        $transportRules = @(Get-TransportRule -ResultSize Unlimited -ErrorAction Stop)
        $escaped = [regex]::Escape($SmtpAddress)
        $hits = 0

        foreach ($rule in $transportRules) {
            $reasons = New-Object System.Collections.Generic.List[string]

            if (@($rule.RedirectMessageTo) -match $escaped) { $reasons.Add("RedirectMessageTo") }
            if (@($rule.BlindCopyTo) -match $escaped) { $reasons.Add("BlindCopyTo") }
            if (@($rule.CopyTo) -match $escaped) { $reasons.Add("CopyTo") }
            if (@($rule.AddToRecipients) -match $escaped) { $reasons.Add("AddToRecipients") }

            if ($reasons.Count -gt 0) {
                $hits++

                $rows.Add([PSCustomObject]@{
                    Check   = "Transport Rule: $($rule.Name)"
                    Status  = "ALERT"
                    Details = "Targets this address via: $($reasons -join ', '). State: $($rule.State). Priority: $($rule.Priority)."
                })
            }
        }

        if ($hits -eq 0) {
            $rows.Add([PSCustomObject]@{
                Check   = "Transport Rules"
                Status  = "OK"
                Details = "Scanned $($transportRules.Count) transport rule(s). None redirect, BCC, or add this address as a recipient."
            })
        }
    }
    catch {
        $rows.Add([PSCustomObject]@{
            Check   = "Transport Rules"
            Status  = "WARN"
            Details = "Error scanning transport rules: $($_.Exception.Message)"
        })
    }

    return $rows
}

# =============================
# Send Rights Analyzer (shared mailboxes)
# =============================

function Get-SendRightsTool {
    param(
        [Parameter(Mandatory)]
        [string]$InputIdentity
    )

    $mailbox = Resolve-MailboxIdentity -InputIdentity $InputIdentity
    $rows = New-Object System.Collections.Generic.List[object]

    $rows.Add([PSCustomObject]@{
        Type    = "Mailbox"
        User    = Format-ToolValue $mailbox.PrimarySmtpAddress
        Rights  = Format-ToolValue $mailbox.RecipientTypeDetails
        Details = Format-ToolValue $mailbox.DisplayName
    })

    $delegation = @(Get-MailboxDelegationTool -Mailbox $mailbox)

    foreach ($entry in $delegation) {
        $rows.Add([PSCustomObject]@{
            Type    = $entry.PermissionType
            User    = $entry.Delegate
            Rights  = $entry.AccessRights
            Details = $entry.Notes
        })
    }

    if ($delegation.Count -eq 0) {
        $rows.Add([PSCustomObject]@{
            Type    = "Info"
            User    = "(none)"
            Rights  = ""
            Details = "Nobody has Full Access, Send As, or Send on Behalf on this mailbox."
        })
    }

    # Sent item copy behavior - the usual shared mailbox complaint
    $sentAsCopy = Format-ToolValue $mailbox.MessageCopyForSentAsEnabled
    $sobCopy = Format-ToolValue $mailbox.MessageCopyForSendOnBehalfEnabled

    $rows.Add([PSCustomObject]@{
        Type    = "Setting"
        User    = "MessageCopyForSentAsEnabled"
        Rights  = $sentAsCopy
        Details = if ($sentAsCopy -eq "True") {
                      "Send As messages ARE copied to this mailbox's Sent Items."
                  } else {
                      "Send As messages only land in the SENDER's Sent Items. Fix: Set-Mailbox -MessageCopyForSentAsEnabled `$true"
                  }
    })

    $rows.Add([PSCustomObject]@{
        Type    = "Setting"
        User    = "MessageCopyForSendOnBehalfEnabled"
        Rights  = $sobCopy
        Details = if ($sobCopy -eq "True") {
                      "Send on Behalf messages ARE copied to this mailbox's Sent Items."
                  } else {
                      "Send on Behalf messages only land in the SENDER's Sent Items. Fix: Set-Mailbox -MessageCopyForSendOnBehalfEnabled `$true"
                  }
    })

    return $rows
}

# =============================
# Executive Calendar Troubleshooting
# =============================

function Get-ExecutiveCalendarTool {
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $mailboxIdentity = $Mailbox.PrimarySmtpAddress.ToString()
    $rows = New-Object System.Collections.Generic.List[object]

    function Add-CalRow {
        param(
            [string]$Category,
            [string]$Item,
            [string]$Value,
            [string]$Notes
        )

        $rows.Add([PSCustomObject]@{
            Category = $Category
            Item     = $Item
            Value    = $Value
            Notes    = $Notes
        })
    }

    # Calendar folder permissions, flagging true delegates vs plain Editors
    try {
        $calendarPerms = @(Get-MailboxFolderPermission -Identity "${mailboxIdentity}:\Calendar" -ErrorAction Stop)

        foreach ($perm in $calendarPerms) {
            $flags = ""

            if ($perm.PSObject.Properties.Name -contains "SharingPermissionFlags") {
                $flags = Format-ToolValue $perm.SharingPermissionFlags
            }

            $note = ""

            if ($flags -match "Delegate") {
                $note = "TRUE DELEGATE - receives meeting requests and responses."

                if ($flags -match "CanViewPrivateItems") {
                    $note += " Can view private items."
                }
            }
            elseif (($perm.AccessRights -join ", ") -match "Editor") {
                $note = "Editor rights but NOT a delegate - can edit the calendar, does NOT receive meeting messages. This is the usual cause of 'my assistant isn't getting my invites'."
            }

            $valueText = ($perm.AccessRights -join ", ")

            if ($flags) {
                $valueText = "$valueText [$flags]"
            }

            Add-CalRow "Calendar Permissions" $perm.User.ToString() $valueText $note
        }
    }
    catch {
        Add-CalRow "Calendar Permissions" "Error" "" $_.Exception.Message
    }

    # Send on Behalf (Outlook delegate setup writes here)
    $sendOnBehalfDelegates = @($Mailbox.GrantSendOnBehalfTo)

    if ($sendOnBehalfDelegates.Count -gt 0 -and $null -ne $sendOnBehalfDelegates[0]) {
        foreach ($delegate in $sendOnBehalfDelegates) {
            Add-CalRow "Send on Behalf" $delegate.ToString() "GrantSendOnBehalfTo" "Outlook delegate setup normally adds this automatically."
        }
    }
    else {
        Add-CalRow "Send on Behalf" "(none)" "" "No send-on-behalf access granted."
    }

    # Calendar processing / meeting message handling
    try {
        $calProcessing = Get-CalendarProcessing -Identity $mailboxIdentity -ErrorAction Stop

        Add-CalRow "Meeting Handling" "AutomateProcessing" (Format-ToolValue $calProcessing.AutomateProcessing) "AutoUpdate is normal for user mailboxes. AutoAccept is for rooms/resources."
        Add-CalRow "Meeting Handling" "ResourceDelegates" (Format-ToolValue $calProcessing.ResourceDelegates) ""
        Add-CalRow "Meeting Handling" "ForwardRequestsToDelegates" (Format-ToolValue $calProcessing.ForwardRequestsToDelegates) "If True, delegates receive copies of meeting requests."
        Add-CalRow "Meeting Handling" "DeleteSubject" (Format-ToolValue $calProcessing.DeleteSubject) ""
        Add-CalRow "Meeting Handling" "DeleteComments" (Format-ToolValue $calProcessing.DeleteComments) ""
        Add-CalRow "Meeting Handling" "AddOrganizerToSubject" (Format-ToolValue $calProcessing.AddOrganizerToSubject) ""
        Add-CalRow "Meeting Handling" "RemovePrivateProperty" (Format-ToolValue $calProcessing.RemovePrivateProperty) ""
        Add-CalRow "Meeting Handling" "AllowConflicts" (Format-ToolValue $calProcessing.AllowConflicts) ""
    }
    catch {
        Add-CalRow "Meeting Handling" "Calendar Processing" "Error" $_.Exception.Message
    }

    # Full access holders - relevant because full access includes the calendar
    try {
        $fullAccess = @(Get-MailboxDelegationTool -Mailbox $Mailbox | Where-Object { $_.PermissionType -eq "Full Access" })

        if ($fullAccess.Count -gt 0) {
            foreach ($entry in $fullAccess) {
                Add-CalRow "Full Access" $entry.Delegate $entry.AccessRights "Full mailbox access includes the calendar."
            }
        }
        else {
            Add-CalRow "Full Access" "(none)" "" ""
        }
    }
    catch {
        Add-CalRow "Full Access" "Error" "" $_.Exception.Message
    }

    Add-CalRow "Automapping" "Note" "Not queryable" "EXO does not expose an automapping flag on existing permissions. It is only set when the permission is added (Add-MailboxPermission -AutoMapping `$true/`$false). To change it, remove and re-add the permission."

    return $rows
}

# =============================
# Distribution Group Memberships
# =============================

function Get-DistributionGroupTool {
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $rows = New-Object System.Collections.Generic.List[object]

    try {
        $dn = $Mailbox.DistinguishedName

        if ([string]::IsNullOrWhiteSpace($dn)) {
            $recipient = Get-Recipient -Identity $Mailbox.PrimarySmtpAddress.ToString() -ErrorAction Stop
            $dn = $recipient.DistinguishedName
        }

        # Escape single quotes for the OPath filter
        $dnFiltered = $dn -replace "'", "''"

        $groups = @(Get-DistributionGroup -Filter "Members -eq '$dnFiltered'" -ResultSize Unlimited -ErrorAction Stop)

        foreach ($group in $groups) {
            $groupKind = switch ($group.RecipientTypeDetails.ToString()) {
                "MailUniversalSecurityGroup"     { "Mail-Enabled Security Group" }
                "MailUniversalDistributionGroup" { "Distribution Group" }
                "RoomList"                       { "Room List" }
                default                          { $group.RecipientTypeDetails.ToString() }
            }

            $rows.Add([PSCustomObject]@{
                GroupName          = Format-ToolValue $group.DisplayName
                PrimarySmtpAddress = Format-ToolValue $group.PrimarySmtpAddress
                GroupType          = $groupKind
                ManagedBy          = Format-ToolValue $group.ManagedBy
                Notes              = ""
            })
        }

        if ($groups.Count -eq 0) {
            $rows.Add([PSCustomObject]@{
                GroupName          = "(none found)"
                PrimarySmtpAddress = ""
                GroupType          = ""
                ManagedBy          = ""
                Notes              = "No static distribution or mail-enabled security group memberships."
            })
        }

        $rows.Add([PSCustomObject]@{
            GroupName          = "Note"
            PrimarySmtpAddress = ""
            GroupType          = "Dynamic Groups"
            ManagedBy          = ""
            Notes              = "Dynamic distribution groups and Entra ID dynamic groups resolve membership by query and cannot be listed per-user from EXO. Check the Entra admin center for those."
        })
    }
    catch {
        return New-InfoRow -Message "Error pulling group memberships: $($_.Exception.Message)"
    }

    return $rows
}

# =============================
# Audit Investigation (Search-UnifiedAuditLog)
# =============================

function Search-AuditTool {
    param(
        [Parameter(Mandatory)]
        [string]$UserId,

        [Parameter(Mandatory)]
        [datetime]$StartDate,

        [Parameter(Mandatory)]
        [datetime]$EndDate,

        [Parameter(Mandatory)]
        [ValidateSet("DeletedItems", "CalendarDeletes", "LoginActivity")]
        [string]$Mode
    )

    try {
        $operations = switch ($Mode) {
            "DeletedItems"    { @("SoftDelete", "HardDelete", "MoveToDeletedItems") }
            "CalendarDeletes" { @("SoftDelete", "HardDelete", "MoveToDeletedItems") }
            "LoginActivity"   { @("MailboxLogin", "UserLoggedIn") }
        }

        $records = @(Search-UnifiedAuditLog `
            -StartDate $StartDate `
            -EndDate $EndDate.AddDays(1) `
            -UserIds $UserId `
            -Operations $operations `
            -ResultSize 1000 `
            -ErrorAction Stop)

        if ($records.Count -eq 0) {
            return New-InfoRow -Message "No audit records found for $UserId between $($StartDate.ToShortDateString()) and $($EndDate.ToShortDateString()). Check that the range is within your audit retention period and that mailbox auditing is enabled."
        }

        $rows = New-Object System.Collections.Generic.List[object]

        foreach ($record in $records) {
            $clientIP = ""
            $client = ""
            $items = ""
            $folder = ""
            $result = ""

            try {
                $auditData = $record.AuditData | ConvertFrom-Json

                if ($auditData.ClientIP) { $clientIP = $auditData.ClientIP }
                elseif ($auditData.ClientIPAddress) { $clientIP = $auditData.ClientIPAddress }

                if ($auditData.ClientInfoString) { $client = $auditData.ClientInfoString }
                elseif ($auditData.ClientProcessName) { $client = $auditData.ClientProcessName }
                elseif ($auditData.UserAgent) { $client = $auditData.UserAgent }

                if ($auditData.AffectedItems) {
                    $items = (@($auditData.AffectedItems) | ForEach-Object { $_.Subject }) -join "; "
                    $folder = (@($auditData.AffectedItems) | ForEach-Object { $_.ParentFolder.Path } | Select-Object -Unique) -join "; "
                }
                elseif ($auditData.Item -and $auditData.Item.Subject) {
                    $items = $auditData.Item.Subject
                }

                if ([string]::IsNullOrWhiteSpace($folder) -and $auditData.Folder -and $auditData.Folder.Path) {
                    $folder = $auditData.Folder.Path
                }

                if ($auditData.ResultStatus) { $result = $auditData.ResultStatus }

                # For calendar mode, keep only calendar-related records
                if ($Mode -eq "CalendarDeletes") {
                    $isCalendar = $false

                    if ($folder -match "Calendar") { $isCalendar = $true }

                    if ($auditData.AffectedItems) {
                        foreach ($affected in @($auditData.AffectedItems)) {
                            if ($affected.ItemClass -like "IPM.Appointment*" -or $affected.ItemClass -like "IPM.Schedule.Meeting*") {
                                $isCalendar = $true
                            }
                        }
                    }

                    if (-not $isCalendar) { continue }
                }
            }
            catch {
                $items = "(could not parse AuditData)"
            }

            $rows.Add([PSCustomObject]@{
                Date      = Format-ToolValue $record.CreationDate
                Operation = Format-ToolValue $record.Operations
                User      = Format-ToolValue $record.UserIds
                ClientIP  = $clientIP
                Client    = $client
                Items     = $items
                Folder    = $folder
                Result    = $result
            })
        }

        if ($rows.Count -eq 0) {
            return New-InfoRow -Message "Audit records were returned, but none were calendar items in this range."
        }

        return $rows
    }
    catch {
        return New-InfoRow -Message "Audit search failed: $($_.Exception.Message)  --  Note: Search-UnifiedAuditLog requires the 'View-Only Audit Logs' or 'Audit Logs' role."
    }
}

# =============================
# Message Trace
# =============================

function Search-MessageTraceTool {
    param(
        [string]$Sender,
        [string]$Recipient,
        [string]$MessageId,

        [Parameter(Mandatory)]
        [datetime]$StartDate,

        [Parameter(Mandatory)]
        [datetime]$EndDate
    )

    try {
        $params = @{
            StartDate   = $StartDate
            EndDate     = $EndDate.AddDays(1)
            ErrorAction = "Stop"
        }

        if (-not [string]::IsNullOrWhiteSpace($Sender))    { $params["SenderAddress"] = $Sender.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($Recipient)) { $params["RecipientAddress"] = $Recipient.Trim() }
        if (-not [string]::IsNullOrWhiteSpace($MessageId)) { $params["MessageId"] = $MessageId.Trim() }

        # StartDate + EndDate + ErrorAction = 3. Anything more means a real filter was supplied.
        if ($params.Keys.Count -le 3) {
            return New-InfoRow -Message "Enter at least a sender, recipient, or message ID before searching."
        }

        # Prefer the V2 cmdlet (longer history); fall back to classic Get-MessageTrace
        if (Get-Command Get-MessageTraceV2 -ErrorAction SilentlyContinue) {
            $script:TraceCmdletUsed = "Get-MessageTraceV2"
            $trace = @(Get-MessageTraceV2 @params)
        }
        else {
            $script:TraceCmdletUsed = "Get-MessageTrace"
            $trace = @(Get-MessageTrace @params)
        }

        if ($trace.Count -eq 0) {
            return New-InfoRow -Message "No messages found. Note: $($script:TraceCmdletUsed) only covers recent history. Older mail requires a historical search in the admin center."
        }

        return $trace | Select-Object `
            @{ Name = "Received";         Expression = { Format-ToolValue $_.Received } },
            @{ Name = "SenderAddress";    Expression = { Format-ToolValue $_.SenderAddress } },
            @{ Name = "RecipientAddress"; Expression = { Format-ToolValue $_.RecipientAddress } },
            @{ Name = "Subject";          Expression = { Format-ToolValue $_.Subject } },
            @{ Name = "Status";           Expression = { Format-ToolValue $_.Status } },
            @{ Name = "FromIP";           Expression = { Format-ToolValue $_.FromIP } },
            @{ Name = "Size";             Expression = { Format-ToolValue $_.Size } },
            @{ Name = "MessageTraceId";   Expression = { Format-ToolValue $_.MessageTraceId } }
    }
    catch {
        return New-InfoRow -Message "Message trace failed: $($_.Exception.Message)"
    }
}

function Show-TraceDetailForm {
    param(
        [Parameter(Mandatory)]
        [string]$MessageTraceId,

        [Parameter(Mandatory)]
        [string]$RecipientAddress,

        [Parameter(Mandatory)]
        [datetime]$StartDate,

        [Parameter(Mandatory)]
        [datetime]$EndDate
    )

    try {
        if (Get-Command Get-MessageTraceDetailV2 -ErrorAction SilentlyContinue) {
            $details = @(Get-MessageTraceDetailV2 `
                -MessageTraceId $MessageTraceId `
                -RecipientAddress $RecipientAddress `
                -StartDate $StartDate `
                -EndDate $EndDate.AddDays(1) `
                -ErrorAction Stop)
        }
        else {
            $details = @(Get-MessageTraceDetail `
                -MessageTraceId $MessageTraceId `
                -RecipientAddress $RecipientAddress `
                -ErrorAction Stop)
        }

        if ($details.Count -eq 0) {
            $detailRows = New-InfoRow -Message "No detail events returned for this message."
        }
        else {
            $detailRows = $details | Select-Object `
                @{ Name = "Date";   Expression = { Format-ToolValue $_.Date } },
                @{ Name = "Event";  Expression = { Format-ToolValue $_.Event } },
                @{ Name = "Action"; Expression = { Format-ToolValue $_.Action } },
                @{ Name = "Detail"; Expression = { Format-ToolValue $_.Detail } }
        }
    }
    catch {
        $detailRows = New-InfoRow -Message "Detail lookup failed: $($_.Exception.Message)"
    }

    $detailForm = New-Object System.Windows.Forms.Form
    $detailForm.Text = "Message Trace Detail - $RecipientAddress"
    $detailForm.Size = New-Object System.Drawing.Size(950, 500)
    $detailForm.StartPosition = "CenterParent"

    $detailGrid = New-Object System.Windows.Forms.DataGridView
    $detailGrid.Dock = "Fill"
    $detailGrid.AutoSizeColumnsMode = "Fill"
    $detailGrid.AllowUserToAddRows = $false
    $detailGrid.ReadOnly = $true
    $detailGrid.RowHeadersVisible = $false
    $detailGrid.SelectionMode = "FullRowSelect"

    $detailForm.Controls.Add($detailGrid)
    $detailGrid.DataSource = ConvertTo-DataTable -InputObject @($detailRows)

    [void]$detailForm.ShowDialog()
}

# =============================
# Mailbox Diagnostics
# =============================

function Get-MailboxDiagnosticsTool {
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox
    )

    $mailboxIdentity = $Mailbox.PrimarySmtpAddress.ToString()
    $rows = New-Object System.Collections.Generic.List[object]

    function Add-DiagRow {
        param(
            [string]$Category,
            [string]$Setting,
            [string]$Value,
            [string]$Notes
        )

        $rows.Add([PSCustomObject]@{
            Category = $Category
            Setting  = $Setting
            Value    = $Value
            Notes    = $Notes
        })
    }

    Add-DiagRow "Connection" "Connected Admin" $script:ConnectedAdminUPN ""
    Add-DiagRow "Connection" "ExchangeOnlineManagement Version" $script:LoadedEXOModuleVersion ""
    Add-DiagRow "Connection" "ExchangeOnlineManagement Path" $script:LoadedEXOModulePath ""

    Add-DiagRow "Mailbox" "Display Name" (Format-ToolValue $Mailbox.DisplayName) ""
    Add-DiagRow "Mailbox" "Primary SMTP" (Format-ToolValue $Mailbox.PrimarySmtpAddress) ""
    Add-DiagRow "Mailbox" "Alias" (Format-ToolValue $Mailbox.Alias) ""
    Add-DiagRow "Mailbox" "Recipient Type" (Format-ToolValue $Mailbox.RecipientTypeDetails) ""
    Add-DiagRow "Mailbox" "Hidden From GAL" (Format-ToolValue $Mailbox.HiddenFromAddressListsEnabled) ""
    Add-DiagRow "Mailbox" "Litigation Hold" (Format-ToolValue $Mailbox.LitigationHoldEnabled) ""
    Add-DiagRow "Mailbox" "Archive Status" (Format-ToolValue $Mailbox.ArchiveStatus) ""

    try {
        $stats = Get-EXOMailboxStatistics -Identity $mailboxIdentity -ErrorAction Stop

        Add-DiagRow "Statistics" "Total Item Size" (Format-ToolValue $stats.TotalItemSize) ""
        Add-DiagRow "Statistics" "Item Count" (Format-ToolValue $stats.ItemCount) ""
        Add-DiagRow "Statistics" "Deleted Item Count" (Format-ToolValue $stats.DeletedItemCount) ""
        Add-DiagRow "Statistics" "Last Logon Time" (Format-ToolValue $stats.LastLogonTime) ""
    }
    catch {
        Add-DiagRow "Statistics" "Mailbox Statistics" "Error" $_.Exception.Message
    }

    Add-DiagRow "Forwarding" "Forwarding SMTP Address" (Format-ToolValue $Mailbox.ForwardingSmtpAddress) ""
    Add-DiagRow "Forwarding" "Forwarding Address" (Format-ToolValue $Mailbox.ForwardingAddress) ""
    Add-DiagRow "Forwarding" "Deliver To Mailbox And Forward" (Format-ToolValue $Mailbox.DeliverToMailboxAndForward) ""

    try {
        $rules = @(Get-InboxRule -Mailbox $mailboxIdentity -IncludeHidden -ErrorAction Stop)

        $forwardRules = @($rules | Where-Object {
            $_.ForwardTo -or
            $_.RedirectTo -or
            $_.ForwardAsAttachmentTo
        })

        Add-DiagRow "Inbox Rules" "Total Inbox Rules" ($rules.Count.ToString()) "Includes hidden rules."
        Add-DiagRow "Inbox Rules" "Enabled Rules" (($rules | Where-Object { $_.Enabled -eq $true }).Count.ToString()) ""
        Add-DiagRow "Inbox Rules" "Forward/Redirect Rules" ($forwardRules.Count.ToString()) "Useful for checking suspicious or unexpected forwarding."
    }
    catch {
        Add-DiagRow "Inbox Rules" "Inbox Rule Check" "Error" $_.Exception.Message
    }

    try {
        $delegation = @(Get-MailboxDelegationTool -Mailbox $Mailbox | Where-Object { $_.PermissionType })
        Add-DiagRow "Delegation" "Mailbox-Level Delegation Entries" ($delegation.Count.ToString()) "Includes Full Access, Send As, and Send on Behalf."
    }
    catch {
        Add-DiagRow "Delegation" "Delegation Check" "Error" $_.Exception.Message
    }

    try {
        $inboxPerms = @(Get-MailboxFolderPermission -Identity "${mailboxIdentity}:\Inbox" -ErrorAction Stop)
        $explicitInboxPerms = @($inboxPerms | Where-Object { $_.User.ToString() -notin @("Default", "Anonymous") })

        Add-DiagRow "Folder Permissions" "Explicit Inbox Permission Entries" ($explicitInboxPerms.Count.ToString()) ""
    }
    catch {
        Add-DiagRow "Folder Permissions" "Inbox Permission Check" "Error" $_.Exception.Message
    }

    try {
        $calendarPerms = @(Get-MailboxFolderPermission -Identity "${mailboxIdentity}:\Calendar" -ErrorAction Stop)
        $explicitCalendarPerms = @($calendarPerms | Where-Object { $_.User.ToString() -notin @("Default", "Anonymous") })

        Add-DiagRow "Folder Permissions" "Explicit Calendar Permission Entries" ($explicitCalendarPerms.Count.ToString()) ""
    }
    catch {
        Add-DiagRow "Folder Permissions" "Calendar Permission Check" "Error" $_.Exception.Message
    }

    try {
        $calendarProcessing = Get-CalendarProcessing -Identity $mailboxIdentity -ErrorAction Stop

        Add-DiagRow "Calendar Processing" "Automate Processing" (Format-ToolValue $calendarProcessing.AutomateProcessing) "Most useful for room/resource/shared mailbox scenarios."
        Add-DiagRow "Calendar Processing" "Add Organizer To Subject" (Format-ToolValue $calendarProcessing.AddOrganizerToSubject) ""
        Add-DiagRow "Calendar Processing" "Delete Subject" (Format-ToolValue $calendarProcessing.DeleteSubject) ""
        Add-DiagRow "Calendar Processing" "Delete Comments" (Format-ToolValue $calendarProcessing.DeleteComments) ""
        Add-DiagRow "Calendar Processing" "Remove Private Property" (Format-ToolValue $calendarProcessing.RemovePrivateProperty) ""
    }
    catch {
        Add-DiagRow "Calendar Processing" "Calendar Processing Check" "Not Available or Error" $_.Exception.Message
    }

    return $rows
}

# =============================
# CSV Export
# =============================

function ConvertFrom-GridToObjects {
    # Reads straight off the grid rather than the DataSource so that whatever
    # sorting the user clicked into, and whatever rows they highlighted, is respected.
    param(
        [Parameter(Mandatory)]
        [System.Windows.Forms.DataGridView]$Grid,

        [switch]$SelectedOnly
    )

    $objects = New-Object System.Collections.Generic.List[object]

    if ($Grid.Columns.Count -eq 0 -or $Grid.Rows.Count -eq 0) {
        return ,$objects
    }

    if ($SelectedOnly -and $Grid.SelectedRows.Count -gt 0) {
        $rows = @($Grid.SelectedRows) | Sort-Object Index
    }
    else {
        $rows = @($Grid.Rows)
    }

    foreach ($row in $rows) {
        if ($row.IsNewRow) { continue }

        $ordered = [ordered]@{}

        foreach ($column in $Grid.Columns) {
            $ordered[$column.Name] = "$($row.Cells[$column.Name].Value)"
        }

        $objects.Add([PSCustomObject]$ordered)
    }

    return ,$objects
}

function Get-SafeFileNamePart {
    param(
        [string]$Text
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return "export"
    }

    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    $builder = New-Object System.Text.StringBuilder

    foreach ($char in $Text.ToCharArray()) {
        if ($invalid -contains $char -or $char -eq ' ') {
            [void]$builder.Append('_')
        }
        else {
            [void]$builder.Append($char)
        }
    }

    return $builder.ToString()
}

function Get-ExportContextName {
    if ($null -ne $script:ResolvedMailbox) {
        $alias = "$($script:ResolvedMailbox.Alias)"

        if (-not [string]::IsNullOrWhiteSpace($alias)) {
            return Get-SafeFileNamePart -Text $alias
        }

        return Get-SafeFileNamePart -Text "$($script:ResolvedMailbox.PrimarySmtpAddress)"
    }

    return "MailboxTools"
}

function Export-GridSelection {
    param(
        [Parameter(Mandatory)]
        [object[]]$Targets,        # each item: PSCustomObject with Name + Grid

        [Parameter(Mandatory)]
        [string]$Path,             # folder when separate, full file path when combined

        [switch]$Combined,

        [switch]$SelectedOnly
    )

    # Pull each grid once and cache it
    $harvested = New-Object System.Collections.Generic.List[object]

    foreach ($target in $Targets) {
        $objects = @(ConvertFrom-GridToObjects -Grid $target.Grid -SelectedOnly:$SelectedOnly)

        if ($objects.Count -gt 0) {
            $harvested.Add([PSCustomObject]@{
                Name    = $target.Name
                Objects = $objects
            })
        }
    }

    if ($harvested.Count -eq 0) {
        throw "Nothing to export. The selected tab(s) have no rows$(if ($SelectedOnly) { ', or no rows are highlighted' })."
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $context = Get-ExportContextName
    $written = New-Object System.Collections.Generic.List[string]

    if ($Combined) {
        # Union every column across the chosen tabs, blank-filling gaps,
        # and prefix a Section column so rows stay traceable to their tab.
        $allColumns = New-Object System.Collections.Generic.List[string]

        foreach ($entry in $harvested) {
            foreach ($object in $entry.Objects) {
                foreach ($propertyName in $object.PSObject.Properties.Name) {
                    if (-not $allColumns.Contains($propertyName)) {
                        $allColumns.Add($propertyName)
                    }
                }
            }
        }

        $combinedRows = New-Object System.Collections.Generic.List[object]

        foreach ($entry in $harvested) {
            foreach ($object in $entry.Objects) {
                $ordered = [ordered]@{
                    Section = $entry.Name
                }

                foreach ($columnName in $allColumns) {
                    if ($object.PSObject.Properties.Name -contains $columnName) {
                        $ordered[$columnName] = $object.$columnName
                    }
                    else {
                        $ordered[$columnName] = ""
                    }
                }

                $combinedRows.Add([PSCustomObject]$ordered)
            }
        }

        $combinedRows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        $written.Add($Path)
    }
    else {
        foreach ($entry in $harvested) {
            $namePart = Get-SafeFileNamePart -Text $entry.Name
            $filePath = Join-Path $Path "$($context)_$($namePart)_$timestamp.csv"

            $entry.Objects | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
            $written.Add($filePath)
        }
    }

    return $written
}

function Show-ExportDialog {
    # Returns a settings object, or $null if the user cancelled.
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Export to CSV"
    $dialog.Size = New-Object System.Drawing.Size(460, 480)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $labelPick = New-Object System.Windows.Forms.Label
    $labelPick.Text = "Choose what to export:"
    $labelPick.Location = New-Object System.Drawing.Point(15, 12)
    $labelPick.Size = New-Object System.Drawing.Size(300, 20)
    $dialog.Controls.Add($labelPick)

    $checkedList = New-Object System.Windows.Forms.CheckedListBox
    $checkedList.Location = New-Object System.Drawing.Point(15, 35)
    $checkedList.Size = New-Object System.Drawing.Size(415, 200)
    $checkedList.CheckOnClick = $true
    $dialog.Controls.Add($checkedList)

    # Only offer tabs that actually hold rows; note the row count so it is obvious what is in there.
    $available = New-Object System.Collections.Generic.List[object]

    foreach ($target in $script:ExportTargets) {
        $rowCount = 0

        if ($target.Grid.Columns.Count -gt 0) {
            $rowCount = @($target.Grid.Rows | Where-Object { -not $_.IsNewRow }).Count
        }

        if ($rowCount -gt 0) {
            $available.Add($target)
            [void]$checkedList.Items.Add("$($target.Name)  ($rowCount rows)", $false)
        }
    }

    if ($available.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "No tabs have any data yet. Run a lookup or a search first.",
            "Nothing to Export",
            "OK",
            "Information"
        )

        return $null
    }

    # Pre-check the tab currently in view, since that is usually the intent
    $currentName = $tabControl.SelectedTab.Text

    for ($i = 0; $i -lt $available.Count; $i++) {
        if ($available[$i].Name -eq $currentName) {
            $checkedList.SetItemChecked($i, $true)
        }
    }

    $buttonAll = New-Object System.Windows.Forms.Button
    $buttonAll.Text = "Select All"
    $buttonAll.Location = New-Object System.Drawing.Point(15, 243)
    $buttonAll.Size = New-Object System.Drawing.Size(90, 26)
    $dialog.Controls.Add($buttonAll)

    $buttonNone = New-Object System.Windows.Forms.Button
    $buttonNone.Text = "Select None"
    $buttonNone.Location = New-Object System.Drawing.Point(112, 243)
    $buttonNone.Size = New-Object System.Drawing.Size(90, 26)
    $dialog.Controls.Add($buttonNone)

    $buttonAll.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $true)
        }
    })

    $buttonNone.Add_Click({
        for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
            $checkedList.SetItemChecked($i, $false)
        }
    })

    $groupRows = New-Object System.Windows.Forms.GroupBox
    $groupRows.Text = "Rows"
    $groupRows.Location = New-Object System.Drawing.Point(15, 278)
    $groupRows.Size = New-Object System.Drawing.Size(415, 55)
    $dialog.Controls.Add($groupRows)

    $radioAllRows = New-Object System.Windows.Forms.RadioButton
    $radioAllRows.Text = "All rows"
    $radioAllRows.Location = New-Object System.Drawing.Point(12, 22)
    $radioAllRows.Size = New-Object System.Drawing.Size(90, 22)
    $radioAllRows.Checked = $true
    $groupRows.Controls.Add($radioAllRows)

    $radioSelectedRows = New-Object System.Windows.Forms.RadioButton
    $radioSelectedRows.Text = "Highlighted rows only (falls back to all if none highlighted)"
    $radioSelectedRows.Location = New-Object System.Drawing.Point(110, 22)
    $radioSelectedRows.Size = New-Object System.Drawing.Size(300, 22)
    $groupRows.Controls.Add($radioSelectedRows)

    $groupLayout = New-Object System.Windows.Forms.GroupBox
    $groupLayout.Text = "File layout"
    $groupLayout.Location = New-Object System.Drawing.Point(15, 340)
    $groupLayout.Size = New-Object System.Drawing.Size(415, 55)
    $dialog.Controls.Add($groupLayout)

    $radioSeparate = New-Object System.Windows.Forms.RadioButton
    $radioSeparate.Text = "One CSV per tab"
    $radioSeparate.Location = New-Object System.Drawing.Point(12, 22)
    $radioSeparate.Size = New-Object System.Drawing.Size(130, 22)
    $radioSeparate.Checked = $true
    $groupLayout.Controls.Add($radioSeparate)

    $radioCombined = New-Object System.Windows.Forms.RadioButton
    $radioCombined.Text = "Single combined CSV (adds a Section column)"
    $radioCombined.Location = New-Object System.Drawing.Point(150, 22)
    $radioCombined.Size = New-Object System.Drawing.Size(260, 22)
    $groupLayout.Controls.Add($radioCombined)

    $buttonOk = New-Object System.Windows.Forms.Button
    $buttonOk.Text = "Export"
    $buttonOk.Location = New-Object System.Drawing.Point(240, 405)
    $buttonOk.Size = New-Object System.Drawing.Size(90, 30)
    $dialog.Controls.Add($buttonOk)

    $buttonCancel = New-Object System.Windows.Forms.Button
    $buttonCancel.Text = "Cancel"
    $buttonCancel.Location = New-Object System.Drawing.Point(340, 405)
    $buttonCancel.Size = New-Object System.Drawing.Size(90, 30)
    $buttonCancel.DialogResult = "Cancel"
    $dialog.Controls.Add($buttonCancel)
    $dialog.CancelButton = $buttonCancel

    $script:ExportDialogResult = $null

    $buttonOk.Add_Click({
        if ($checkedList.CheckedIndices.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Tick at least one tab to export.",
                "Nothing Selected",
                "OK",
                "Warning"
            )

            return
        }

        $chosen = New-Object System.Collections.Generic.List[object]

        foreach ($index in $checkedList.CheckedIndices) {
            $chosen.Add($available[$index])
        }

        $script:ExportDialogResult = [PSCustomObject]@{
            Targets      = @($chosen)
            SelectedOnly = $radioSelectedRows.Checked
            Combined     = $radioCombined.Checked
        }

        $dialog.DialogResult = "OK"
        $dialog.Close()
    })

    [void]$dialog.ShowDialog()

    return $script:ExportDialogResult
}

function Invoke-ExportFlow {
    param(
        [Parameter(Mandatory)]
        [object[]]$Targets,

        [switch]$Combined,

        [switch]$SelectedOnly
    )

    $context = Get-ExportContextName
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

    if ($Combined) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "$($context)_MailboxTools_$timestamp.csv"
        $saveDialog.Title = "Save combined CSV"

        if ($saveDialog.ShowDialog() -ne "OK") {
            return
        }

        $destination = $saveDialog.FileName
    }
    else {
        $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDialog.Description = "Choose a folder for the CSV files"

        if ($folderDialog.ShowDialog() -ne "OK") {
            return
        }

        $destination = $folderDialog.SelectedPath
    }

    try {
        Set-StatusText "Exporting..."

        $written = @(Export-GridSelection `
            -Targets $Targets `
            -Path $destination `
            -Combined:$Combined `
            -SelectedOnly:$SelectedOnly)

        Set-StatusText "Exported $($written.Count) file(s)"

        $fileList = ($written | ForEach-Object { Split-Path $_ -Leaf }) -join "`r`n"
        $folder = if ($Combined) { Split-Path $written[0] -Parent } else { $destination }

        $answer = [System.Windows.Forms.MessageBox]::Show(
            "Exported $($written.Count) file(s) to:`r`n$folder`r`n`r`n$fileList`r`n`r`nOpen the folder now?",
            "Export Complete",
            "YesNo",
            "Information"
        )

        if ($answer -eq "Yes") {
            Start-Process explorer.exe -ArgumentList "`"$folder`""
        }
    }
    catch {
        Set-StatusText "Export error"

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Export Error",
            "OK",
            "Error"
        )
    }
}

# =============================
# GUI Setup
# =============================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Exchange Online Mailbox Tools"
$form.Size = New-Object System.Drawing.Size(1300, 800)
$form.StartPosition = "CenterScreen"

$labelUser = New-Object System.Windows.Forms.Label
$labelUser.Text = "LANID or Email:"
$labelUser.Location = New-Object System.Drawing.Point(20, 20)
$labelUser.Size = New-Object System.Drawing.Size(100, 25)
$form.Controls.Add($labelUser)

$textUser = New-Object System.Windows.Forms.TextBox
$textUser.Location = New-Object System.Drawing.Point(125, 18)
$textUser.Size = New-Object System.Drawing.Size(330, 25)
$form.Controls.Add($textUser)

$buttonSearch = New-Object System.Windows.Forms.Button
$buttonSearch.Text = "Lookup Mailbox"
$buttonSearch.Location = New-Object System.Drawing.Point(470, 16)
$buttonSearch.Size = New-Object System.Drawing.Size(130, 30)
$form.Controls.Add($buttonSearch)

$buttonClear = New-Object System.Windows.Forms.Button
$buttonClear.Text = "Clear"
$buttonClear.Location = New-Object System.Drawing.Point(610, 16)
$buttonClear.Size = New-Object System.Drawing.Size(80, 30)
$form.Controls.Add($buttonClear)

$buttonExportTab = New-Object System.Windows.Forms.Button
$buttonExportTab.Text = "Export Tab"
$buttonExportTab.Location = New-Object System.Drawing.Point(700, 16)
$buttonExportTab.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($buttonExportTab)

$buttonExportPick = New-Object System.Windows.Forms.Button
$buttonExportPick.Text = "Export..."
$buttonExportPick.Location = New-Object System.Drawing.Point(810, 16)
$buttonExportPick.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($buttonExportPick)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Not connected"
$statusLabel.Location = New-Object System.Drawing.Point(20, 55)
$statusLabel.Size = New-Object System.Drawing.Size(1230, 25)
$form.Controls.Add($statusLabel)

$connectionLabel = New-Object System.Windows.Forms.Label
$connectionLabel.Text = "Module: Not loaded | Connected Admin: Not connected"
$connectionLabel.Location = New-Object System.Drawing.Point(20, 75)
$connectionLabel.Size = New-Object System.Drawing.Size(1230, 25)
$form.Controls.Add($connectionLabel)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20, 105)
$tabControl.Size = New-Object System.Drawing.Size(1240, 640)
$tabControl.Multiline = $true
$form.Controls.Add($tabControl)

# Status helper - defined here because it touches $statusLabel and $form
function Set-StatusText {
    param(
        [string]$Text
    )

    $statusLabel.Text = "Status: $Text"
    $form.Refresh()
}

function Test-ToolReady {
    param(
        [switch]$RequireMailbox
    )

    if (-not $script:ConnectedToEXO) {
        [System.Windows.Forms.MessageBox]::Show(
            "The tool is not connected to Exchange Online. Close and reopen the tool, then sign in again.",
            "Not Connected",
            "OK",
            "Warning"
        )

        return $false
    }

    if ($RequireMailbox -and $null -eq $script:ResolvedMailbox) {
        [System.Windows.Forms.MessageBox]::Show(
            "Look up a mailbox first, then run this action.",
            "No Mailbox Selected",
            "OK",
            "Warning"
        )

        return $false
    }

    return $true
}

function New-GridTab {
    param(
        [string]$TabName,
        [int]$PanelHeight = 0
    )

    $tab = New-Object System.Windows.Forms.TabPage
    $tab.Text = $TabName
    $tabControl.Controls.Add($tab)

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Dock = "Fill"
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.AllowUserToAddRows = $false
    $grid.ReadOnly = $true
    $grid.RowHeadersVisible = $false
    $grid.SelectionMode = "FullRowSelect"
    $grid.MultiSelect = $true

    $panel = $null

    if ($PanelHeight -gt 0) {
        $panel = New-Object System.Windows.Forms.Panel
        $panel.Dock = "Top"
        $panel.Height = $PanelHeight
    }

    # Add the grid first so Dock=Fill sits below the Dock=Top panel
    $tab.Controls.Add($grid)

    if ($panel) {
        $tab.Controls.Add($panel)
    }

    return [PSCustomObject]@{
        Tab   = $tab
        Grid  = $grid
        Panel = $panel
    }
}

# --- Auto-populating tabs ---
$tabSummary      = New-GridTab -TabName "Mailbox Summary"
$gridSummary     = $tabSummary.Grid

$tabDiagnostics  = New-GridTab -TabName "Mailbox Diagnostics"
$gridDiagnostics = $tabDiagnostics.Grid

$tabRules        = New-GridTab -TabName "Inbox Rules"
$gridRules       = $tabRules.Grid

$tabPermSummary  = New-GridTab -TabName "Permissions Summary"
$gridPermSummary = $tabPermSummary.Grid

$tabForwarding   = New-GridTab -TabName "Forwarding Check" -PanelHeight 45
$gridForwarding  = $tabForwarding.Grid

$tabExecCalendar  = New-GridTab -TabName "Executive Calendar"
$gridExecCalendar = $tabExecCalendar.Grid

# --- On-demand tabs ---
$tabGroups       = New-GridTab -TabName "Distribution Groups" -PanelHeight 45
$gridGroups      = $tabGroups.Grid

$tabSendRights   = New-GridTab -TabName "Send Rights Analyzer" -PanelHeight 45
$gridSendRights  = $tabSendRights.Grid

$tabAudit        = New-GridTab -TabName "Audit Investigation" -PanelHeight 80
$gridAudit       = $tabAudit.Grid

$tabTrace        = New-GridTab -TabName "Message Trace" -PanelHeight 80
$gridTrace       = $tabTrace.Grid

# =============================
# Export wiring
# =============================

# Names here must match the tab text so the export dialog can pre-tick the active tab.
$script:ExportTargets = @(
    [PSCustomObject]@{ Name = "Mailbox Summary";     Grid = $gridSummary }
    [PSCustomObject]@{ Name = "Mailbox Diagnostics"; Grid = $gridDiagnostics }
    [PSCustomObject]@{ Name = "Inbox Rules";         Grid = $gridRules }
    [PSCustomObject]@{ Name = "Permissions Summary"; Grid = $gridPermSummary }
    [PSCustomObject]@{ Name = "Forwarding Check";    Grid = $gridForwarding }
    [PSCustomObject]@{ Name = "Executive Calendar";  Grid = $gridExecCalendar }
    [PSCustomObject]@{ Name = "Distribution Groups"; Grid = $gridGroups }
    [PSCustomObject]@{ Name = "Send Rights Analyzer"; Grid = $gridSendRights }
    [PSCustomObject]@{ Name = "Audit Investigation"; Grid = $gridAudit }
    [PSCustomObject]@{ Name = "Message Trace";       Grid = $gridTrace }
)

$buttonExportTab.Add_Click({
    $currentName = $tabControl.SelectedTab.Text
    $target = $script:ExportTargets | Where-Object { $_.Name -eq $currentName } | Select-Object -First 1

    if ($null -eq $target) {
        [System.Windows.Forms.MessageBox]::Show(
            "This tab has nothing exportable.",
            "Nothing to Export",
            "OK",
            "Information"
        )

        return
    }

    if ($target.Grid.Columns.Count -eq 0 -or $target.Grid.Rows.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "The '$currentName' tab has no data yet.",
            "Nothing to Export",
            "OK",
            "Information"
        )

        return
    }

    $context = Get-ExportContextName
    $namePart = Get-SafeFileNamePart -Text $currentName
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    $saveDialog.FileName = "$($context)_$($namePart)_$timestamp.csv"
    $saveDialog.Title = "Export '$currentName'"

    if ($saveDialog.ShowDialog() -ne "OK") {
        return
    }

    try {
        Set-StatusText "Exporting $currentName..."

        $objects = @(ConvertFrom-GridToObjects -Grid $target.Grid)
        $objects | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8 -ErrorAction Stop

        Set-StatusText "Exported $($objects.Count) row(s) to $(Split-Path $saveDialog.FileName -Leaf)"
    }
    catch {
        Set-StatusText "Export error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Export Error", "OK", "Error")
    }
})

$buttonExportPick.Add_Click({
    $settings = Show-ExportDialog

    if ($null -eq $settings) {
        return
    }

    Invoke-ExportFlow `
        -Targets $settings.Targets `
        -Combined:$settings.Combined `
        -SelectedOnly:$settings.SelectedOnly
})

# =============================
# Forwarding Check tab: controls + status coloring
# =============================

$buttonTransportCheck = New-Object System.Windows.Forms.Button
$buttonTransportCheck.Text = "Check Transport Rules"
$buttonTransportCheck.Location = New-Object System.Drawing.Point(10, 8)
$buttonTransportCheck.Size = New-Object System.Drawing.Size(180, 28)
$tabForwarding.Panel.Controls.Add($buttonTransportCheck)

$labelForwardingHint = New-Object System.Windows.Forms.Label
$labelForwardingHint.Text = "Scans all org transport rules for redirect / BCC / add-recipient targeting this mailbox. Can be slow in large orgs."
$labelForwardingHint.Location = New-Object System.Drawing.Point(205, 13)
$labelForwardingHint.Size = New-Object System.Drawing.Size(900, 20)
$tabForwarding.Panel.Controls.Add($labelForwardingHint)

$gridForwarding.Add_DataBindingComplete({
    if (-not $gridForwarding.Columns.Contains("Status")) {
        return
    }

    foreach ($row in $gridForwarding.Rows) {
        $status = "$($row.Cells['Status'].Value)"

        switch ($status) {
            "OK"    { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(212, 237, 218) }
            "WARN"  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255, 243, 205) }
            "ALERT" { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(248, 215, 218) }
            "INFO"  { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(217, 237, 247) }
        }
    }
})

$buttonTransportCheck.Add_Click({
    if (-not (Test-ToolReady -RequireMailbox)) { return }

    try {
        $smtp = $script:ResolvedMailbox.PrimarySmtpAddress.ToString()

        Set-StatusText "Scanning transport rules for $smtp..."

        $lightChecks = @(Get-ForwardingCheckTool -Mailbox $script:ResolvedMailbox | Where-Object { $_.Check -ne "Transport Rules" })
        $transportChecks = @(Get-TransportRuleCheckTool -SmtpAddress $smtp)

        $gridForwarding.DataSource = ConvertTo-DataTable -InputObject @($lightChecks + $transportChecks)

        Set-StatusText "Forwarding + transport rule check complete for $smtp"
    }
    catch {
        Set-StatusText "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Transport Rule Check Error", "OK", "Error")
    }
})

# =============================
# Distribution Groups tab controls
# =============================

$buttonLoadGroups = New-Object System.Windows.Forms.Button
$buttonLoadGroups.Text = "Load Group Memberships"
$buttonLoadGroups.Location = New-Object System.Drawing.Point(10, 8)
$buttonLoadGroups.Size = New-Object System.Drawing.Size(180, 28)
$tabGroups.Panel.Controls.Add($buttonLoadGroups)

$labelGroupsHint = New-Object System.Windows.Forms.Label
$labelGroupsHint.Text = "Distribution groups and mail-enabled security groups the looked-up mailbox belongs to. On-demand because the filter query can be slow."
$labelGroupsHint.Location = New-Object System.Drawing.Point(205, 13)
$labelGroupsHint.Size = New-Object System.Drawing.Size(950, 20)
$tabGroups.Panel.Controls.Add($labelGroupsHint)

$buttonLoadGroups.Add_Click({
    if (-not (Test-ToolReady -RequireMailbox)) { return }

    try {
        $smtp = $script:ResolvedMailbox.PrimarySmtpAddress.ToString()

        Set-StatusText "Loading group memberships for $smtp (this can take a moment)..."

        $gridGroups.DataSource = ConvertTo-DataTable -InputObject @(Get-DistributionGroupTool -Mailbox $script:ResolvedMailbox)

        Set-StatusText "Group memberships loaded for $smtp"
    }
    catch {
        Set-StatusText "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Group Lookup Error", "OK", "Error")
    }
})

# =============================
# Send Rights Analyzer tab controls
# =============================

$labelSendRights = New-Object System.Windows.Forms.Label
$labelSendRights.Text = "Shared Mailbox:"
$labelSendRights.Location = New-Object System.Drawing.Point(10, 13)
$labelSendRights.Size = New-Object System.Drawing.Size(95, 20)
$tabSendRights.Panel.Controls.Add($labelSendRights)

$textSendRights = New-Object System.Windows.Forms.TextBox
$textSendRights.Location = New-Object System.Drawing.Point(110, 10)
$textSendRights.Size = New-Object System.Drawing.Size(300, 25)
$tabSendRights.Panel.Controls.Add($textSendRights)

$buttonSendRights = New-Object System.Windows.Forms.Button
$buttonSendRights.Text = "Analyze"
$buttonSendRights.Location = New-Object System.Drawing.Point(425, 8)
$buttonSendRights.Size = New-Object System.Drawing.Size(110, 28)
$tabSendRights.Panel.Controls.Add($buttonSendRights)

$labelSendRightsHint = New-Object System.Windows.Forms.Label
$labelSendRightsHint.Text = "Full Access / Send As / Send on Behalf + Sent Items copy settings. Leave blank to use the looked-up mailbox."
$labelSendRightsHint.Location = New-Object System.Drawing.Point(550, 13)
$labelSendRightsHint.Size = New-Object System.Drawing.Size(650, 20)
$tabSendRights.Panel.Controls.Add($labelSendRightsHint)

$buttonSendRights.Add_Click({
    if (-not (Test-ToolReady)) { return }

    $target = $textSendRights.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($target)) {
        if ($null -ne $script:ResolvedMailbox) {
            $target = $script:ResolvedMailbox.PrimarySmtpAddress.ToString()
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Enter a shared mailbox address, or look up a mailbox first.",
                "Missing Input",
                "OK",
                "Warning"
            )

            return
        }
    }

    try {
        Set-StatusText "Analyzing send rights for $target..."

        $gridSendRights.DataSource = ConvertTo-DataTable -InputObject @(Get-SendRightsTool -InputIdentity $target)

        Set-StatusText "Send rights analysis complete for $target"
    }
    catch {
        Set-StatusText "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Send Rights Error", "OK", "Error")
    }
})

$textSendRights.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $buttonSendRights.PerformClick()
    }
})

# =============================
# Audit Investigation tab controls
# =============================

$labelAuditStart = New-Object System.Windows.Forms.Label
$labelAuditStart.Text = "Start Date:"
$labelAuditStart.Location = New-Object System.Drawing.Point(10, 13)
$labelAuditStart.Size = New-Object System.Drawing.Size(70, 20)
$tabAudit.Panel.Controls.Add($labelAuditStart)

$dateAuditStart = New-Object System.Windows.Forms.DateTimePicker
$dateAuditStart.Format = "Short"
$dateAuditStart.Location = New-Object System.Drawing.Point(85, 10)
$dateAuditStart.Size = New-Object System.Drawing.Size(120, 25)
$dateAuditStart.Value = (Get-Date).AddDays(-7)
$tabAudit.Panel.Controls.Add($dateAuditStart)

$labelAuditEnd = New-Object System.Windows.Forms.Label
$labelAuditEnd.Text = "End Date:"
$labelAuditEnd.Location = New-Object System.Drawing.Point(225, 13)
$labelAuditEnd.Size = New-Object System.Drawing.Size(65, 20)
$tabAudit.Panel.Controls.Add($labelAuditEnd)

$dateAuditEnd = New-Object System.Windows.Forms.DateTimePicker
$dateAuditEnd.Format = "Short"
$dateAuditEnd.Location = New-Object System.Drawing.Point(295, 10)
$dateAuditEnd.Size = New-Object System.Drawing.Size(120, 25)
$dateAuditEnd.Value = Get-Date
$tabAudit.Panel.Controls.Add($dateAuditEnd)

$labelAuditHint = New-Object System.Windows.Forms.Label
$labelAuditHint.Text = "Runs Search-UnifiedAuditLog against the looked-up mailbox. Requires the 'View-Only Audit Logs' role. Max 1000 records per search."
$labelAuditHint.Location = New-Object System.Drawing.Point(440, 13)
$labelAuditHint.Size = New-Object System.Drawing.Size(760, 20)
$tabAudit.Panel.Controls.Add($labelAuditHint)

$buttonAuditDeleted = New-Object System.Windows.Forms.Button
$buttonAuditDeleted.Text = "Search Deleted Events"
$buttonAuditDeleted.Location = New-Object System.Drawing.Point(10, 45)
$buttonAuditDeleted.Size = New-Object System.Drawing.Size(170, 28)
$tabAudit.Panel.Controls.Add($buttonAuditDeleted)

$buttonAuditCalendar = New-Object System.Windows.Forms.Button
$buttonAuditCalendar.Text = "Search Calendar Deletes"
$buttonAuditCalendar.Location = New-Object System.Drawing.Point(190, 45)
$buttonAuditCalendar.Size = New-Object System.Drawing.Size(180, 28)
$tabAudit.Panel.Controls.Add($buttonAuditCalendar)

$buttonAuditLogin = New-Object System.Windows.Forms.Button
$buttonAuditLogin.Text = "Search Login Activity"
$buttonAuditLogin.Location = New-Object System.Drawing.Point(380, 45)
$buttonAuditLogin.Size = New-Object System.Drawing.Size(170, 28)
$tabAudit.Panel.Controls.Add($buttonAuditLogin)

function Invoke-AuditSearchFromUI {
    param(
        [string]$Mode,
        [string]$ModeLabel
    )

    if (-not (Test-ToolReady -RequireMailbox)) { return }

    try {
        $upn = $script:ResolvedMailbox.UserPrincipalName.ToString()

        Set-StatusText "Searching audit log ($ModeLabel) for $upn... this can take 30+ seconds."

        $results = @(Search-AuditTool `
            -UserId $upn `
            -StartDate $dateAuditStart.Value.Date `
            -EndDate $dateAuditEnd.Value.Date `
            -Mode $Mode)

        $gridAudit.DataSource = ConvertTo-DataTable -InputObject $results

        Set-StatusText "Audit search ($ModeLabel) complete for $upn - $($results.Count) row(s)"
    }
    catch {
        Set-StatusText "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Audit Search Error", "OK", "Error")
    }
}

$buttonAuditDeleted.Add_Click({ Invoke-AuditSearchFromUI -Mode "DeletedItems" -ModeLabel "deleted items" })
$buttonAuditCalendar.Add_Click({ Invoke-AuditSearchFromUI -Mode "CalendarDeletes" -ModeLabel "calendar deletes" })
$buttonAuditLogin.Add_Click({ Invoke-AuditSearchFromUI -Mode "LoginActivity" -ModeLabel "login activity" })

# =============================
# Message Trace tab controls
# =============================

$labelTraceSender = New-Object System.Windows.Forms.Label
$labelTraceSender.Text = "Sender:"
$labelTraceSender.Location = New-Object System.Drawing.Point(10, 13)
$labelTraceSender.Size = New-Object System.Drawing.Size(50, 20)
$tabTrace.Panel.Controls.Add($labelTraceSender)

$textTraceSender = New-Object System.Windows.Forms.TextBox
$textTraceSender.Location = New-Object System.Drawing.Point(65, 10)
$textTraceSender.Size = New-Object System.Drawing.Size(230, 25)
$tabTrace.Panel.Controls.Add($textTraceSender)

$labelTraceRecipient = New-Object System.Windows.Forms.Label
$labelTraceRecipient.Text = "Recipient:"
$labelTraceRecipient.Location = New-Object System.Drawing.Point(310, 13)
$labelTraceRecipient.Size = New-Object System.Drawing.Size(60, 20)
$tabTrace.Panel.Controls.Add($labelTraceRecipient)

$textTraceRecipient = New-Object System.Windows.Forms.TextBox
$textTraceRecipient.Location = New-Object System.Drawing.Point(375, 10)
$textTraceRecipient.Size = New-Object System.Drawing.Size(230, 25)
$tabTrace.Panel.Controls.Add($textTraceRecipient)

$labelTraceMessageId = New-Object System.Windows.Forms.Label
$labelTraceMessageId.Text = "Message ID:"
$labelTraceMessageId.Location = New-Object System.Drawing.Point(620, 13)
$labelTraceMessageId.Size = New-Object System.Drawing.Size(75, 20)
$tabTrace.Panel.Controls.Add($labelTraceMessageId)

$textTraceMessageId = New-Object System.Windows.Forms.TextBox
$textTraceMessageId.Location = New-Object System.Drawing.Point(700, 10)
$textTraceMessageId.Size = New-Object System.Drawing.Size(400, 25)
$tabTrace.Panel.Controls.Add($textTraceMessageId)

$labelTraceStart = New-Object System.Windows.Forms.Label
$labelTraceStart.Text = "Start:"
$labelTraceStart.Location = New-Object System.Drawing.Point(10, 48)
$labelTraceStart.Size = New-Object System.Drawing.Size(40, 20)
$tabTrace.Panel.Controls.Add($labelTraceStart)

$dateTraceStart = New-Object System.Windows.Forms.DateTimePicker
$dateTraceStart.Format = "Short"
$dateTraceStart.Location = New-Object System.Drawing.Point(55, 45)
$dateTraceStart.Size = New-Object System.Drawing.Size(120, 25)
$dateTraceStart.Value = (Get-Date).AddDays(-7)
$tabTrace.Panel.Controls.Add($dateTraceStart)

$labelTraceEnd = New-Object System.Windows.Forms.Label
$labelTraceEnd.Text = "End:"
$labelTraceEnd.Location = New-Object System.Drawing.Point(190, 48)
$labelTraceEnd.Size = New-Object System.Drawing.Size(35, 20)
$tabTrace.Panel.Controls.Add($labelTraceEnd)

$dateTraceEnd = New-Object System.Windows.Forms.DateTimePicker
$dateTraceEnd.Format = "Short"
$dateTraceEnd.Location = New-Object System.Drawing.Point(230, 45)
$dateTraceEnd.Size = New-Object System.Drawing.Size(120, 25)
$dateTraceEnd.Value = Get-Date
$tabTrace.Panel.Controls.Add($dateTraceEnd)

$buttonTraceSearch = New-Object System.Windows.Forms.Button
$buttonTraceSearch.Text = "Search"
$buttonTraceSearch.Location = New-Object System.Drawing.Point(370, 43)
$buttonTraceSearch.Size = New-Object System.Drawing.Size(110, 28)
$tabTrace.Panel.Controls.Add($buttonTraceSearch)

$buttonTraceUseMailbox = New-Object System.Windows.Forms.Button
$buttonTraceUseMailbox.Text = "Use Looked-Up Mailbox as Recipient"
$buttonTraceUseMailbox.Location = New-Object System.Drawing.Point(490, 43)
$buttonTraceUseMailbox.Size = New-Object System.Drawing.Size(230, 28)
$tabTrace.Panel.Controls.Add($buttonTraceUseMailbox)

$labelTraceHint = New-Object System.Windows.Forms.Label
$labelTraceHint.Text = "Double-click a result row for hop-by-hop delivery detail. Fill in at least one of sender / recipient / message ID."
$labelTraceHint.Location = New-Object System.Drawing.Point(730, 48)
$labelTraceHint.Size = New-Object System.Drawing.Size(480, 20)
$tabTrace.Panel.Controls.Add($labelTraceHint)

$buttonTraceUseMailbox.Add_Click({
    if ($null -eq $script:ResolvedMailbox) {
        [System.Windows.Forms.MessageBox]::Show(
            "Look up a mailbox first.",
            "No Mailbox Selected",
            "OK",
            "Warning"
        )

        return
    }

    $textTraceRecipient.Text = $script:ResolvedMailbox.PrimarySmtpAddress.ToString()
})

$buttonTraceSearch.Add_Click({
    if (-not (Test-ToolReady)) { return }

    try {
        Set-StatusText "Running message trace..."

        $results = @(Search-MessageTraceTool `
            -Sender $textTraceSender.Text `
            -Recipient $textTraceRecipient.Text `
            -MessageId $textTraceMessageId.Text `
            -StartDate $dateTraceStart.Value.Date `
            -EndDate $dateTraceEnd.Value.Date)

        $gridTrace.DataSource = ConvertTo-DataTable -InputObject $results

        Set-StatusText "Message trace complete - $($results.Count) row(s) via $($script:TraceCmdletUsed)"
    }
    catch {
        Set-StatusText "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Message Trace Error", "OK", "Error")
    }
})

$gridTrace.Add_CellDoubleClick({
    param($eventSender, $eventArgs)

    if ($eventArgs.RowIndex -lt 0) { return }
    if (-not $gridTrace.Columns.Contains("MessageTraceId")) { return }

    $row = $gridTrace.Rows[$eventArgs.RowIndex]
    $traceId = "$($row.Cells['MessageTraceId'].Value)"
    $recipient = "$($row.Cells['RecipientAddress'].Value)"

    if ([string]::IsNullOrWhiteSpace($traceId) -or [string]::IsNullOrWhiteSpace($recipient)) {
        return
    }

    try {
        Set-StatusText "Pulling trace detail for $recipient..."

        Show-TraceDetailForm `
            -MessageTraceId $traceId `
            -RecipientAddress $recipient `
            -StartDate $dateTraceStart.Value.Date `
            -EndDate $dateTraceEnd.Value.Date

        Set-StatusText "Trace detail closed"
    }
    catch {
        Set-StatusText "Error"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Trace Detail Error", "OK", "Error")
    }
})

# =============================
# Lookup Action
# =============================

$buttonSearch.Add_Click({
    $inputIdentity = $textUser.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($inputIdentity)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a LANID, alias, UPN, or email address.",
            "Missing Input",
            "OK",
            "Warning"
        )

        return
    }

    if (-not (Test-ToolReady)) { return }

    try {
        Set-StatusText "Resolving mailbox..."

        $mailbox = Resolve-MailboxIdentity -InputIdentity $inputIdentity
        $script:ResolvedMailbox = $mailbox

        $mailboxIdentity = $mailbox.PrimarySmtpAddress.ToString()

        Set-StatusText "Pulling mailbox summary for $mailboxIdentity..."
        $gridSummary.DataSource = ConvertTo-DataTable -InputObject @(Get-MailboxSummaryTool -Mailbox $mailbox)

        Set-StatusText "Running mailbox diagnostics for $mailboxIdentity..."
        $gridDiagnostics.DataSource = ConvertTo-DataTable -InputObject @(Get-MailboxDiagnosticsTool -Mailbox $mailbox)

        Set-StatusText "Pulling inbox rules for $mailboxIdentity..."
        $gridRules.DataSource = ConvertTo-DataTable -InputObject @(Get-InboxRulesTool -MailboxIdentity $mailboxIdentity)

        Set-StatusText "Building permissions summary for $mailboxIdentity..."
        $gridPermSummary.DataSource = ConvertTo-DataTable -InputObject @(Get-PermissionsSummaryTool -Mailbox $mailbox)

        Set-StatusText "Running forwarding checks for $mailboxIdentity..."
        $gridForwarding.DataSource = ConvertTo-DataTable -InputObject @(Get-ForwardingCheckTool -Mailbox $mailbox)

        Set-StatusText "Pulling calendar delegation for $mailboxIdentity..."
        $gridExecCalendar.DataSource = ConvertTo-DataTable -InputObject @(Get-ExecutiveCalendarTool -Mailbox $mailbox)

        # Reset on-demand grids so stale data from a previous user is not left behind
        $gridGroups.DataSource = $null
        $gridAudit.DataSource = $null
        $gridSendRights.DataSource = $null

        Set-StatusText "Completed lookup for $($mailbox.DisplayName) <$mailboxIdentity>"
    }
    catch {
        Set-StatusText "Error"

        [System.Windows.Forms.MessageBox]::Show(
            $_.Exception.Message,
            "Lookup Error",
            "OK",
            "Error"
        )
    }
})

$textUser.Add_KeyDown({
    if ($_.KeyCode -eq "Enter") {
        $buttonSearch.PerformClick()
    }
})

# =============================
# Clear Action
# =============================

$buttonClear.Add_Click({
    $textUser.Clear()
    $textSendRights.Clear()
    $textTraceSender.Clear()
    $textTraceRecipient.Clear()
    $textTraceMessageId.Clear()

    $gridSummary.DataSource = $null
    $gridDiagnostics.DataSource = $null
    $gridRules.DataSource = $null
    $gridPermSummary.DataSource = $null
    $gridForwarding.DataSource = $null
    $gridExecCalendar.DataSource = $null
    $gridGroups.DataSource = $null
    $gridSendRights.DataSource = $null
    $gridAudit.DataSource = $null
    $gridTrace.DataSource = $null

    $script:ResolvedMailbox = $null
    $statusLabel.Text = "Status: Cleared"
})

# =============================
# Connect on Startup
# =============================

$form.Add_Shown({
    $form.Activate()

    Set-StatusText "Connecting to Exchange Online..."

    $connected = Connect-EXOTool

    if ($connected) {
        $script:ConnectedToEXO = $true

        $statusLabel.Text = "Status: Connected to Exchange Online"
        $connectionLabel.Text = "Module: ExchangeOnlineManagement $($script:LoadedEXOModuleVersion) | Connected Admin: $($script:ConnectedAdminUPN)"
    }
    else {
        $script:ConnectedToEXO = $false

        $statusLabel.Text = "Status: Not connected"
        $connectionLabel.Text = "Module: ExchangeOnlineManagement $($script:LoadedEXOModuleVersion) | Connected Admin: Not connected"
    }
})

# =============================
# Disconnect on Close
# =============================

$form.Add_FormClosing({
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }
    catch {
        # Ignore disconnect errors.
    }
})

[void]$form.ShowDialog()
