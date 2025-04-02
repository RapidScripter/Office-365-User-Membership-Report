<#
=============================================================================================
Name:           Office365 User Membership Report Using PowerShell
Description:    This script exports Office 365 user's group details to CSV
============================================================================================
#>

param
(
    [String] $UsersIdentityFile,
    [Switch] $GuestUsersOnly,
    [Switch] $DisabledUsersOnly,
    [Switch] $UsersNotinAnyGroup,
    [string] $TenantId,
    [string] $ClientId,
    [string] $CertificateThumbprint
)

# Check and install Microsoft Graph module if needed
$MsGraphModule = Get-Module Microsoft.Graph -ListAvailable
if($null -eq $MsGraphModule)
{ 
    Write-Host "Important: Microsoft Graph module is unavailable." -ForegroundColor Yellow
    $confirm = Read-Host "Are you sure you want to install Microsoft Graph module? [Y] Yes [N] No"  
    if($confirm -match "[yY]") 
    { 
        Write-Host "Installing Microsoft Graph module..."
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
        Write-Host "Microsoft Graph module installed successfully" -ForegroundColor Green
        Import-Module Microsoft.Graph
    } 
    else
    { 
        Write-Host "Exiting. Microsoft Graph module is required." -ForegroundColor Red
        Exit 
    } 
}

# Connect to Microsoft Graph
try {
    if (($TenantId -ne "") -and ($ClientId -ne "")) {
        if($CertificateThumbprint -ne "") {
            Connect-MgGraph -TenantId $TenantId -AppId $ClientId -CertificateThumbprint $CertificateThumbprint -ErrorAction Stop | Out-Null
        }
        else {
            Write-Host "Certificate thumbprint is required for app authentication" -ForegroundColor Red
            Exit
        }
    }
    else {
        # Interactive login with required permissions
        Connect-MgGraph -Scopes "Directory.Read.All, User.Read.All" -ErrorAction Stop | Out-Null
    }
    Write-Host "Microsoft Graph PowerShell module connected successfully" -ForegroundColor Green
}
catch {
    Write-Host "Connection failed: $_" -ForegroundColor Red
    Exit
}

Function UserDetails {
    if ([string]$UsersIdentityFile -ne "")
    {
        $IdentityList = Import-Csv -Header "UserIdentityValue" $UsersIdentityFile
        foreach ($IdentityValue in $IdentityList) 
        {
            $CurIdentity = $IdentityValue.UserIdentityValue
            try 
            {
                $LiveUser = Get-MgUser -UserId "$CurIdentity" -ExpandProperty MemberOf -ErrorAction SilentlyContinue
                if($GuestUsersOnly.IsPresent -and $LiveUser.UserType -ne "Guest") 
                {
                    continue
                }
                if($DisabledUsersOnly.IsPresent -and $LiveUser.AccountEnabled -eq $true)
                {
                    continue
                }
                ProcessUser
            }
            catch 
            {
                Write-Host "Given UserIdentity: $CurIdentity is not valid/found."
            }
        }
    }
    else 
    {
        if ($GuestUsersOnly.Ispresent -and $DisabledUsersOnly.Ispresent) 
        {
            Get-MgUser -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf -All | Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
        elseif ($DisabledUsersOnly.Ispresent) 
        {
            Get-MgUser -ExpandProperty MemberOf -All | Where-Object { $_.AccountEnabled -eq $false } | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }  
        }
        elseif ($GuestUsersOnly.Ispresent) 
        {
            Get-MgUser -Filter "UserType eq 'Guest'" -ExpandProperty MemberOf -All | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
        else 
        {
            Get-MgUser -ExpandProperty MemberOf -All | ForEach-Object {
                $LiveUser = $_
                ProcessUser
            }
        }
    }
}

Function ProcessUser {
    $GroupList = @()
    $RolesList = @()
    $Script:ProcessedUsers += 1
    $Name = $LiveUser.DisplayName
    Write-Progress -Activity "Processing $Name" -Status "Processed Users Count: $Script:ProcessedUsers" 
    $UserMembership = Get-MgUserMemberOf -UserId $LiveUser.UserPrincipalName | Select-Object -ExpandProperty AdditionalProperties
    $AllGroupData = $UserMembership | Where-object { $_.'@odata.type' -eq "#microsoft.graph.group" }
    if ($AllGroupData -eq $null) 
    {
        $GroupName = " - "
    }
    else 
    {
        if ($UsersNotinAnyGroup.IsPresent) 
        {
            return
        }
        $GroupName = (@($AllGroupData.displayName) -join ',') 
    }
    $AllRoles = $UserMembership | Where-object { $_.'@odata.type' -eq "#microsoft.graph.directoryRole" }
    if ($AllRoles -eq $null) { 
        $RolesList = " - " 
    }
    else
    {
        $RolesList = @($AllRoles.displayName) -join ','
    }
    if ($LiveUser.AccountEnabled -eq $True) 
    {
        $AccountStatus = "Enabled"
    }
    else 
    {
        $AccountStatus = "Disabled"
    }
    if ($LiveUser.Department -eq $null) 
    {
        $Department = " - " 
    }
    else 
    {
        $Department = $LiveUser.Department
    }
    if ($LiveUser.AssignedLicenses -ne "")
    { 
        $LicenseStatus = "Licensed" 
    }
    else 
    {
        $LicenseStatus = "Unlicensed" 
    }
    ExportResults
}

Function ExportResults {
    $Script:ExportedUsers += 1
    $ExportResult = [PSCustomObject] @{
        'Display Name' = $Name
        'Email Address' = $LiveUser.UserPrincipalName
        'Group Name(s)' = $GroupName
        'License Status' = $LicenseStatus
        'Account Status' = $AccountStatus
        'Department' = $Department
        'Admin Roles' = $RolesList
    }
    $ExportResult | Export-Csv -Path $ExportCSVFileName -NoTypeInformation -Append    
}

$ProcessedUsers = 0
$ExportedUsers = 0
$ExportCSVFileName = ".\UserMembershipReport_$((Get-Date -format 'MMM-dd-hh-mm-ss-tt').ToString()).csv"

UserDetails

if ((Test-Path -Path $ExportCSVFileName) -eq "True") { 
    Write-Progress -Activity "--" -Completed
    Write-Host "`nThe output contains " -NoNewline
    Write-Host "$Script:ExportedUsers Users" -ForegroundColor Magenta -NoNewline
    Write-Host " details"
    Write-Host "`nOutput file: $ExportCSVFileName" -ForegroundColor Green 
    
    $prompt = New-Object -ComObject wscript.shell    
    $userInput = $prompt.popup("Do you want to open output file?", 0, "Open Output File", 4)    
    if ($userInput -eq 6) {    
        Invoke-Item "$ExportCSVFileName"
    }  
}
else {
    Write-Host "`nNo data found matching your criteria" -ForegroundColor Red
}

# Disconnect when done
Disconnect-MgGraph | Out-Null