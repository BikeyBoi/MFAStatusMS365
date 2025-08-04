
# Install required modules if needed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module Microsoft.Graph.Users -Scope CurrentUser -Force
}
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}

# Connect to Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All", "UserAuthenticationMethod.Read.All"

# Get only enabled non-system Azure users with valid UPN
$users = Get-MgUser -All -Property "accountEnabled, userPrincipalName" | Where-Object {
    $_.AccountEnabled -eq $true -and
    ![string]::IsNullOrWhiteSpace($_.UserPrincipalName) -and
    $_.UserPrincipalName -notlike "*.onmicrosoft.com" -and
    $_.UserPrincipalName -notlike "*#EXT#*" -and
    $_.UserPrincipalName -notlike "HealthMailbox*"  # Exclude Exchange health mailboxes
}

$results = @()
Write-Host "`nRetrieved $($users.Count) enabled non-system users with valid UPNs"

# Loop through each user account
foreach ($user in $users) {
    Write-Host "`nProcessing $($user.UserPrincipalName)"
    
    $myObject = [PSCustomObject]@{
        UserPrincipalName = $user.UserPrincipalName
        MFAStatus         = "Disabled"  # Default to disabled
        EmailMethod       = $false
        FIDO2Key          = $false
        AuthenticatorApp  = $false
        PasswordOnly      = $false
        PhoneMethod       = $false
        SoftwareOATH      = $false
        TemporaryPass     = $false
        WindowsHello      = $false
    }

    try {
        $MFAData = Get-MgUserAuthenticationMethod -UserId $user.UserPrincipalName -ErrorAction Stop
        
        # Check authentication methods for each user
        $hasMFA = $false
        foreach ($method in $MFAData) {
            Switch ($method.AdditionalProperties["@odata.type"]) {
                "#microsoft.graph.emailAuthenticationMethod" { 
                    $myObject.EmailMethod = $true 
                    $hasMFA = $true
                } 
                "#microsoft.graph.fido2AuthenticationMethod" { 
                    $myObject.FIDO2Key = $true 
                    $hasMFA = $true
                }    
                "#microsoft.graph.microsoftAuthenticatorAuthenticationMethod" { 
                    $myObject.AuthenticatorApp = $true 
                    $hasMFA = $true
                }    
                "#microsoft.graph.passwordAuthenticationMethod" {              
                    $myObject.PasswordOnly = $true 
                }     
                "#microsoft.graph.phoneAuthenticationMethod" { 
                    $myObject.PhoneMethod = $true 
                    $hasMFA = $true
                }   
                "#microsoft.graph.softwareOathAuthenticationMethod" { 
                    $myObject.SoftwareOATH = $true 
                    $hasMFA = $true
                }           
                "#microsoft.graph.temporaryAccessPassAuthenticationMethod" { 
                    $myObject.TemporaryPass = $true 
                }           
                "#microsoft.graph.windowsHelloForBusinessAuthenticationMethod" { 
                    $myObject.WindowsHello = $true 
                }                   
            }
        }
        
        # Update MFA status if any MFA method found
        if ($hasMFA) {
            $myObject.MFAStatus = "Enabled"
        }
    }
    catch {
        Write-Warning "Error retrieving methods for $($user.UserPrincipalName): $_"
        $myObject.MFAStatus = "Error"
    }

    # Collect objects
    $results += $myObject
}

# Export to Excel with professional formatting
$documentsPath = [Environment]::GetFolderPath("MyDocuments")
$excelPath = Join-Path -Path $documentsPath -ChildPath "MFA_Status_Report.xlsx"

$excelParams = @{
    Path          = $excelPath
    WorksheetName = "MFA Status"
    AutoSize      = $true
    AutoFilter    = $true
    FreezeTopRow  = $true
    BoldTopRow    = $true
    ClearSheet    = $true
}

# Add conditional formatting
$conditionalFormat = @(
    New-ConditionalText -Text "Enabled" -BackgroundColor LightGreen -Range "B:B"
    New-ConditionalText -Text "Disabled" -BackgroundColor LightCoral -Range "B:B"
    New-ConditionalText -Text "Error" -BackgroundColor Gold -Range "B:B"
    New-ConditionalText -Text $true -BackgroundColor LightBlue -Range "C:J"
)

$results | Export-Excel @excelParams -ConditionalText $conditionalFormat

Write-Host "`nReport successfully exported to:" -ForegroundColor Green
Write-Host $excelPath -ForegroundColor Cyan

# Open the Excel file
if ($IsWindows -or $env:OS) {
    Start-Process $excelPath
}
elseif ($IsMacOS) {
    Start-Process "open" -ArgumentList $excelPath
}
elseif ($IsLinux) {
    Start-Process "xdg-open" -ArgumentList $excelPath
}