# MailboxToolsTabbedGUI.ps1
# Exchange Online Mailbox Troubleshooting GUI
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
        return $dataTable
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

    return $dataTable
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
        "EmailAddresses"
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

    if ($results.Count -eq 0) {
        return New-InfoRow -Message "No mailbox-level delegation found. Folder-level Inbox and Calendar permissions are shown in their own tabs."
    }

    return $results
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
# GUI Setup
# =============================

$form = New-Object System.Windows.Forms.Form
$form.Text = "Exchange Online Mailbox Tools"
$form.Size = New-Object System.Drawing.Size(1250, 760)
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

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "Status: Not connected"
$statusLabel.Location = New-Object System.Drawing.Point(20, 55)
$statusLabel.Size = New-Object System.Drawing.Size(1180, 25)
$form.Controls.Add($statusLabel)

$connectionLabel = New-Object System.Windows.Forms.Label
$connectionLabel.Text = "Module: Not loaded | Connected Admin: Not connected"
$connectionLabel.Location = New-Object System.Drawing.Point(20, 75)
$connectionLabel.Size = New-Object System.Drawing.Size(1180, 25)
$form.Controls.Add($connectionLabel)

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20, 105)
$tabControl.Size = New-Object System.Drawing.Size(1190, 585)
$form.Controls.Add($tabControl)

function New-GridTab {
    param(
        [string]$TabName
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

    $tab.Controls.Add($grid)

    return $grid
}

$gridSummary = New-GridTab -TabName "Mailbox Summary"
$gridDiagnostics = New-GridTab -TabName "Mailbox Diagnostics"
$gridRules = New-GridTab -TabName "Inbox Rules"
$gridDelegation = New-GridTab -TabName "Mailbox Delegation"
$gridInboxAccess = New-GridTab -TabName "Inbox Permissions"
$gridCalendarAccess = New-GridTab -TabName "Calendar Permissions"

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

    if (-not $script:ConnectedToEXO) {
        [System.Windows.Forms.MessageBox]::Show(
            "The tool is not connected to Exchange Online. Close and reopen the tool, then sign in again.",
            "Not Connected",
            "OK",
            "Warning"
        )

        return
    }

    try {
        $statusLabel.Text = "Status: Resolving mailbox..."
        $form.Refresh()

        $mailbox = Resolve-MailboxIdentity -InputIdentity $inputIdentity
        $script:ResolvedMailbox = $mailbox

        $mailboxIdentity = $mailbox.PrimarySmtpAddress.ToString()

        $statusLabel.Text = "Status: Pulling mailbox summary for $mailboxIdentity..."
        $form.Refresh()
        $gridSummary.DataSource = ConvertTo-DataTable -InputObject @(Get-MailboxSummaryTool -Mailbox $mailbox)

        $statusLabel.Text = "Status: Running mailbox diagnostics for $mailboxIdentity..."
        $form.Refresh()
        $gridDiagnostics.DataSource = ConvertTo-DataTable -InputObject @(Get-MailboxDiagnosticsTool -Mailbox $mailbox)

        $statusLabel.Text = "Status: Pulling inbox rules for $mailboxIdentity..."
        $form.Refresh()
        $gridRules.DataSource = ConvertTo-DataTable -InputObject @(Get-InboxRulesTool -MailboxIdentity $mailboxIdentity)

        $statusLabel.Text = "Status: Pulling mailbox delegation for $mailboxIdentity..."
        $form.Refresh()
        $gridDelegation.DataSource = ConvertTo-DataTable -InputObject @(Get-MailboxDelegationTool -Mailbox $mailbox)

        $statusLabel.Text = "Status: Pulling inbox permissions for $mailboxIdentity..."
        $form.Refresh()
        $gridInboxAccess.DataSource = ConvertTo-DataTable -InputObject @(Get-FolderPermissionTool -MailboxIdentity $mailboxIdentity -FolderName "Inbox")

        $statusLabel.Text = "Status: Pulling calendar permissions for $mailboxIdentity..."
        $form.Refresh()
        $gridCalendarAccess.DataSource = ConvertTo-DataTable -InputObject @(Get-FolderPermissionTool -MailboxIdentity $mailboxIdentity -FolderName "Calendar")

        $statusLabel.Text = "Status: Completed lookup for $($mailbox.DisplayName) <$mailboxIdentity>"
    }
    catch {
        $statusLabel.Text = "Status: Error"

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

    $gridSummary.DataSource = $null
    $gridDiagnostics.DataSource = $null
    $gridRules.DataSource = $null
    $gridDelegation.DataSource = $null
    $gridInboxAccess.DataSource = $null
    $gridCalendarAccess.DataSource = $null

    $script:ResolvedMailbox = $null
    $statusLabel.Text = "Status: Cleared"
})

# =============================
# Connect on Startup
# =============================

$form.Add_Shown({
    $form.Activate()

    $statusLabel.Text = "Status: Connecting to Exchange Online..."
    $form.Refresh()

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
