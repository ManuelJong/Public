<#
.SYNOPSIS
    # Note: The detection.ps1 script will be automatically created and uploaded.
    Intune Win32 App Creation Script for Network Printer Deployment with Automatic Package Creation

.DESCRIPTION
    This script automates the creation of Win32 apps in Microsoft Intune for deploying network printers.
    Enhancements include:
    - Automatic .intunewin package creation if missing
    - Detection script generation
    - Printer installation parameter configuration
    - Icon selection dialog and auto-selection
    - Enhanced error handling and summary output
    - Driver details extraction from INF file
    - Support for DriverWizard.exe and PrnInst.exe for advanced driver installation scenarios
    - Uploads package to Intune with all configuration

.NOTES
    Author:         Manuel Jong
    Version:        1.4.0
    Date Created:   31/12/2021
    Last Modified:  20/08/2025

.VERSION HISTORY
    1.0.0 - Initial release
    1.1.0 - Added Advanced Port Monitor support with PrnInst.exe
    1.2.0 - Removed certificate parameter, added automatic .cat file certificate extraction
    1.3.0 - Added automatic .intunewin package creation using IntuneWinAppUtil.exe
    1.4.0 - Enhanced detection script generation, icon selection dialog, improved error handling, driver details extraction from INF, 
            summary output, and added support for DriverWizard for advanced driver installation.

.REQUIREMENTS
    - IntuneWin32App PowerShell module
    - Microsoft Graph permissions for Intune app management
    - IntuneWinAppUtil.exe (will be downloaded automatically if not present)
    - DriverWizard.exe (optional, for advanced driver installation scenarios)
    - PrnInst.exe (optional, for advanced port monitor scenarios)
    - Administrator rights to install PowerShell modules

.APP REGISTRATION PERMISSIONS NEEDED
    The Azure AD app registration used for authentication must have the following Microsoft Graph API permissions:
    - DeviceManagementApps.ReadWrite.All (Application)
    - DeviceManagementManagedDevices.ReadWrite.All (Application)
    - Directory.Read.All (Application)
    - User.Read (Delegated)
    You must grant admin consent for these permissions.

.EXAMPLE
    .\script.ps1

    This will automatically create the .intunewin package (if needed), create a Win32 app in Intune, and optionally use DriverWizard.exe and PrnInst.exe for advanced driver installation scenarios.

.IMAGE SELECTION FUNCTION
    The script will automatically search for JPEG, JPG, or PNG image files in the script's folder. If only one image is found, it will be used as the app icon. If multiple images are found, you will be prompted to select one via a file dialog. If no images are found, you will be prompted to browse and select an image manually. This icon will be used for the Intune app.

.LINK
    https://github.com/MSEndpointMgr/IntuneWin32App
#>

# ============================================================================== 
# EDITABLE PARAMETERS - MODIFY THESE VALUES FOR YOUR PRINTER CONFIGURATION 
# ============================================================================== 

# Printer Display Name - Use your preferred naming convention
$DisplayName = "Printer-Example-1"

# IP Address Configuration
# IMPORTANT: If you prefix the IP address with "LAN_" (e.g., "LAN_10.5.0.40"), 
# the Install-Printer.ps1 script will automatically use PrnInst.exe and DriverWizard.exe (if present in the same folder)
# to create an Advanced Port Monitor instead of a standard TCP/IP port.
$ipadress = "10.5.0.60"

# Driver Name - Must match exactly the driver name in the INF file
$Drivername = "Example Printer Driver"

# INF File Name - The driver INF file that should be in the same directory
$inffile = "Example.inf"

# Intune Tenant Configuration
$TenantID = "<your-tenant-id>.onmicrosoft.com"
$ClientID = "<your-app-registration-client-id>"

# ============================================================================== 
# END OF EDITABLE PARAMETERS - DO NOT MODIFY BELOW UNLESS YOU KNOW WHAT YOU'RE DOING 
# ============================================================================== 

#region Function to Download and Setup IntuneWinAppUtil
function Get-IntuneWinAppUtil {
    param (
        [string]$TargetPath = "$PSScriptRoot"
    )
    
    $IntuneWinAppUtilPath = Join-Path $TargetPath "IntuneWinAppUtil.exe"
    
    if (Test-Path $IntuneWinAppUtilPath) {
        Write-Host "  ✓ IntuneWinAppUtil.exe found at: $IntuneWinAppUtilPath" -ForegroundColor Green
        return $IntuneWinAppUtilPath
    }
    
    Write-Host "  IntuneWinAppUtil.exe not found. Downloading..." -ForegroundColor Yellow
    
    try {
        # Microsoft's official download URL for IntuneWinAppUtil
        $downloadUrl = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe"
        
        # Use TLS 1.2
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        
        # Download the file
        Write-Host "  Downloading from: $downloadUrl" -ForegroundColor Gray
        Invoke-WebRequest -Uri $downloadUrl -OutFile $IntuneWinAppUtilPath -UseBasicParsing
        
        if (Test-Path $IntuneWinAppUtilPath) {
            Write-Host "  ✓ IntuneWinAppUtil.exe downloaded successfully" -ForegroundColor Green
            return $IntuneWinAppUtilPath
        } else {
            throw "Failed to download IntuneWinAppUtil.exe"
        }
    }
    catch {
        Write-Host "  × Failed to download IntuneWinAppUtil.exe: $_" -ForegroundColor Red
        Write-Host "  Please download manually from: https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool" -ForegroundColor Yellow
        return $null
    }
}
#endregion

#region Function to Create .intunewin Package
function New-IntuneWinPackage {
    param (
        [string]$SourceFolder,
        [string]$SetupFile,
        [string]$OutputFolder
    )
    
    Write-Host "`nCreating .intunewin package..." -ForegroundColor Cyan
    
    # Get or download IntuneWinAppUtil.exe
    $IntuneWinAppUtil = Get-IntuneWinAppUtil -TargetPath $PSScriptRoot
    
    if (-not $IntuneWinAppUtil) {
        Write-Host "× Cannot proceed without IntuneWinAppUtil.exe" -ForegroundColor Red
        return $null
    }
    
    # Verify source files exist
    $SetupFilePath = Join-Path $SourceFolder $SetupFile
    if (-not (Test-Path $SetupFilePath)) {
        Write-Host "× Setup file not found: $SetupFilePath" -ForegroundColor Red
        return $null
    }
    
    Write-Host "  Source folder: $SourceFolder" -ForegroundColor Gray
    Write-Host "  Setup file: $SetupFile" -ForegroundColor Gray
    Write-Host "  Output folder: $OutputFolder" -ForegroundColor Gray
    
    # Build arguments for IntuneWinAppUtil
    $arguments = @(
        "-c", "`"$SourceFolder`"",
        "-s", "`"$SetupFile`"",
        "-o", "`"$OutputFolder`"",
        "-q"  # Quiet mode
    )
    
    try {
        # Run IntuneWinAppUtil.exe
        Write-Host "  Running IntuneWinAppUtil.exe..." -ForegroundColor Yellow
        $process = Start-Process -FilePath $IntuneWinAppUtil -ArgumentList $arguments -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            # Success - find the created .intunewin file
            $packageName = [System.IO.Path]::GetFileNameWithoutExtension($SetupFile) + ".intunewin"
            $packagePath = Join-Path $OutputFolder $packageName
            
            if (Test-Path $packagePath) {
                Write-Host "  ✓ Package created successfully: $packagePath" -ForegroundColor Green
                return $packagePath
            } else {
                Write-Host "  × Package file not found after creation" -ForegroundColor Red
                return $null
            }
        } else {
            Write-Host "  × IntuneWinAppUtil.exe failed with exit code: $($process.ExitCode)" -ForegroundColor Red
            return $null
        }
    }
    catch {
        Write-Host "  × Error running IntuneWinAppUtil.exe: $_" -ForegroundColor Red
        return $null
    }
}
#endregion

#region Check and Create .intunewin Package
Write-Host "================================" -ForegroundColor Cyan
Write-Host "Intune Win32 App Deployment" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan

$IntuneWinFile = "$PSScriptRoot\Install-Printer.intunewin"

# Check if .intunewin package exists
if (-not (Test-Path $IntuneWinFile)) {
    Write-Host "`n.intunewin package not found. Will create it automatically." -ForegroundColor Yellow
    
    # Verify Install-Printer.ps1 exists
    $InstallScript = "$PSScriptRoot\Install-Printer.ps1"
    if (-not (Test-Path $InstallScript)) {
        Write-Host "× Install-Printer.ps1 not found in: $PSScriptRoot" -ForegroundColor Red
        Write-Host "  Please ensure Install-Printer.ps1 is in the same directory as this script." -ForegroundColor Yellow
        exit 1
    }
    
    # Verify driver files exist
    $InfPath = Join-Path $PSScriptRoot $inffile
    if (-not (Test-Path $InfPath)) {
        Write-Host "× INF file not found: $InfPath" -ForegroundColor Red
        Write-Host "  Please ensure all driver files are in the same directory." -ForegroundColor Yellow
        exit 1
    }
    
    # Create the package
    $IntuneWinFile = New-IntuneWinPackage -SourceFolder $PSScriptRoot -SetupFile "Install-Printer.ps1" -OutputFolder $PSScriptRoot
    
    if (-not $IntuneWinFile) {
        Write-Host "`n× Failed to create .intunewin package" -ForegroundColor Red
        exit 1
    }
    
    Write-Host "`n✓ Package created successfully!" -ForegroundColor Green
} else {
    Write-Host "`n✓ .intunewin package found: $IntuneWinFile" -ForegroundColor Green
}
#endregion

#region Initialize Connection
Write-Host "`nConnecting to Microsoft Intune..." -ForegroundColor Yellow
Connect-MSIntuneGraph -TenantID $TenantID -ClientID $ClientID
#endregion

#region Image Selection Dialog
Write-Host "`nLooking for icon images..." -ForegroundColor Cyan

# Find all .jpeg, .jpg, .png files in the root folder
$ImageFiles = Get-ChildItem -Path $PSScriptRoot -File | Where-Object { $_.Extension -match '^\.(jpg|jpeg|png)$' }

if ($ImageFiles.Count -eq 1) {
    # Exactly one image found - use it automatically
    $Imagefile = $ImageFiles[0].FullName
    Write-Host "  ✓ Auto-selected single image found: $($ImageFiles[0].Name)" -ForegroundColor Green
} elseif ($ImageFiles.Count -gt 1) {
    # Multiple images found - prompt user to select
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Image Files (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png"
    $dialog.Title = "Select an icon image for the printer app (optional)"
    $dialog.InitialDirectory = $PSScriptRoot

    # Create a topmost invisible window as owner
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.StartPosition = "CenterScreen"
    $form.ShowInTaskbar = $false
    $form.Size = '0,0'
    $form.Show()
    $form.Hide()

    Write-Host "  Multiple images found. Opening file browser to select an image..." -ForegroundColor Cyan
    if ($dialog.ShowDialog($form) -eq 'OK') {
        $Imagefile = $dialog.FileName
        Write-Host "  ✓ Image selected: $Imagefile" -ForegroundColor Green
    } else {
        Write-Host "  No image file selected. Using default icon." -ForegroundColor Yellow
        $Imagefile = $null
    }
    $form.Dispose()
} else {
    # No images found - prompt user to select
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Image Files (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png"
    $dialog.Title = "No images found - Please select an icon image (optional)"
    $dialog.InitialDirectory = $PSScriptRoot

    # Create a topmost invisible window as owner
    $form = New-Object System.Windows.Forms.Form
    $form.TopMost = $true
    $form.StartPosition = "CenterScreen"
    $form.ShowInTaskbar = $false
    $form.Size = '0,0'
    $form.Show()
    $form.Hide()

    Write-Host "  No image files found. Opening file browser to select an image..." -ForegroundColor Cyan
    if ($dialog.ShowDialog($form) -eq 'OK') {
        $Imagefile = $dialog.FileName
        Write-Host "  ✓ Image selected: $Imagefile" -ForegroundColor Green
    } else {
        Write-Host "  No image file selected. Using default icon." -ForegroundColor Yellow
        $Imagefile = $null
    }
    $form.Dispose()
}
#endregion

#region Create Detection Script
Write-Host "`nCreating detection script..." -ForegroundColor Yellow

# Construct the detection script content that will check if printer is installed
$scriptContent = @"
######################################################################################################################
# Detection.ps1 - Check if a specific printer is installed by looking at the registry
# Auto-generated by Intune Printer Deployment Script v1.3.0
# Author: Manuel Jong
######################################################################################################################

# Define the registry path with the printer name directly in the string
`$RegPath = "HKLM:SYSTEM\CurrentControlSet\Control\Print\Printers\$DisplayName"

# Check if the registry key exists
if (Test-Path `$RegPath) {
    `$RegContent = Get-ItemProperty -Path `$RegPath
    if (`$RegContent.PSChildName -eq `"$DisplayName`") {
        Write-Output "Found it!"
        exit 0  # Success
    }
    else {
        Write-Output "Printer name does not match."
        exit 1  # Name mismatch
    }
} 
else {
    Write-Output "Printer not found in registry."
    exit 1  # Key not found
}
"@

# Write the detection script content to a .ps1 file
$destination = "$PSScriptRoot\Detection.ps1"
Set-Content -Path $destination -Value $scriptContent
Write-Host "  ✓ Detection.ps1 created" -ForegroundColor Green
#endregion

#region Bootstrap PowerShell Package Management
# =========================
# BOOTSTRAP POWERSHELLGET/PACKAGEMANAGEMENT
# =========================
Write-Host "`nBootstrapping PowerShell package management..." -ForegroundColor Yellow
try {
    # Use TLS 1.2 for downloads
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    
    # Ensure PSGallery is trusted
    if (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue) {
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    } else {
        Register-PSRepository -Default -ErrorAction SilentlyContinue
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    }
    
    # Import PackageManagement
    Import-Module PackageManagement -ErrorAction SilentlyContinue
    
    # Ensure NuGet provider exists
    if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -Scope CurrentUser -Force -ErrorAction SilentlyContinue
    }
    
    # Import PowerShellGet
    Import-Module PowerShellGet -ErrorAction SilentlyContinue
    
    # If PowerShellGet still not available, install it
    if (-not (Get-Module PowerShellGet)) {
        Install-Module -Name PowerShellGet -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
        Import-Module PowerShellGet -Force
    }
    
    Write-Host "  ✓ Package management ready" -ForegroundColor Green
} catch {
    Write-Host "  Warning: Bootstrap had issues but continuing: $_" -ForegroundColor Yellow
}

# Now install IntuneWin32App module
if (-not (Get-Module -ListAvailable -Name "IntuneWin32App")) {
    Write-Host "`nIntuneWin32App module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name "IntuneWin32App" -Force
    Write-Host "  ✓ IntuneWin32App module installed" -ForegroundColor Green
}
Import-Module -Name "IntuneWin32App" -Force
#endregion

#region Reconnect to Intune Graph
Write-Host "`nReconnecting to Microsoft Intune Graph..." -ForegroundColor Yellow
Connect-MSIntuneGraph -TenantID $TenantID -ClientID $ClientID
#endregion

#region Create Intune App Configuration
# Create requirement rule for all platforms and Windows 10 20H2
Write-Host "`nCreating app configuration..." -ForegroundColor Yellow
Write-Host "  Creating requirement rule..." -ForegroundColor Gray
$RequirementRule = New-IntuneWin32AppRequirementRule -Architecture "x64" -MinimumSupportedWindowsRelease "W10_21H2"
  
# Create MSI detection rule
Write-Host "  Creating detection rule..." -ForegroundColor Gray
$DetectionRule = New-IntuneWin32AppDetectionRuleScript -ScriptFile $DetectionScriptFile -EnforceSignatureCheck $false -RunAs32Bit $false

# Create custom return code
Write-Host "  Creating return codes..." -ForegroundColor Gray
$ReturnCode = New-IntuneWin32AppReturnCode -ReturnCode 1337 -Type "retry"

# Create icon only if image file was selected
if ($Imagefile -and (Test-Path $Imagefile)) {
    Write-Host "  Creating icon from image..." -ForegroundColor Gray
    try {
        $Icon = New-IntuneWin32AppIcon -FilePath $Imagefile
        Write-Host "  ✓ Icon created successfully" -ForegroundColor Green
    } catch {
        Write-Host "  × Failed to create icon: $_" -ForegroundColor Red
        Write-Host "  Continuing without icon..." -ForegroundColor Yellow
        $Icon = $null
    }
} else {
    Write-Host "  No icon will be added to the app" -ForegroundColor Gray
    $Icon = $null
}
#endregion

#region Extract driver info for description
function Get-DriverDetails {
    param([string]$INFPath)
    $details = @{
        Version = "Unknown"
        Provider = "Unknown"
    }
    if (-not (Test-Path $INFPath)) { return $details }
    $content = Get-Content $INFPath -ErrorAction SilentlyContinue
    $inVersion = $false
    foreach ($line in $content) {
        if ($line -match '^\[Version\]') { $inVersion = $true; continue }
        if ($inVersion) {
            if ($line -match '^\[') { $inVersion = $false; continue }
            if ($line -match 'DriverVer\s*=\s*(.+)') {
                $driverVerLine = $matches[1].Trim()
                if ($driverVerLine -match ',\s*(.+)$') { $details.Version = $matches[1].Trim() }
                elseif ($driverVerLine -match '[\d\.]+') { $details.Version = $matches[0] }
            }
            if ($line -match 'Provider\s*=\s*(.+)') {
                $details.Provider = $matches[1].Trim() -replace '["%]', ''
            }
        }
    }
    return $details
}
#endregion

#region Build Command Lines
$InstallCommandLine = "powershell.exe -executionpolicy bypass -file .\Install-Printer.ps1 -PortName ""$ipadress"" -PrinterIP ""$ipadress"" -PrinterName ""$displayname"" -DriverName ""$Drivername"" -INFFile ""$inffile"""
$UninstallCommandLine = "powershell.exe -executionpolicy bypass -file .\UnInstall-Printer.ps1 -PrinterName ""$DisplayName"""

# Gather driver details for description
$driverDetails = Get-DriverDetails "$PSScriptRoot\$inffile"

# Build enhanced description with summary
$appDescription = @"
Printer Deployment Summary:
- Printer Name: $DisplayName
- IP Address: $ipadress
- Driver Name: $Drivername
- INF File: $inffile
- Provider: $($driverDetails.Provider)
- Driver Version: $($driverDetails.Version)

Other Settings:
- Install Command: $InstallCommandLine
- Uninstall Command: $UninstallCommandLine
"@

Write-Host "`n================================" -ForegroundColor Cyan
Write-Host "App Configuration Summary" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host "  Display Name: $DisplayName" -ForegroundColor White
Write-Host "  IP Address: $ipadress $(if ($ipadress -like 'LAN_*') { '(Advanced Port Monitor)' } else { '(Standard TCP/IP)' })" -ForegroundColor White
Write-Host "  Driver: $Drivername" -ForegroundColor White
Write-Host "  Provider: $($driverDetails.Provider)" -ForegroundColor White
Write-Host "  Driver Version: $($driverDetails.Version)" -ForegroundColor White
Write-Host "  INF File: $inffile" -ForegroundColor White
Write-Host "================================" -ForegroundColor Cyan
#endregion

#region Create Win32 App in Intune
# Create the Win32 app in Intune with better error handling
try {
    Write-Host "`nCreating Win32 app in Intune..." -ForegroundColor Yellow
    
    # Build parameters based on whether icon exists
    $appParams = @{
        FilePath = $IntuneWinFile
        DisplayName = $DisplayName
        Description = $appDescription
        Publisher = "Script Manuel"
        InstallExperience = "system"
        RestartBehavior = "suppress"
        DetectionRule = $DetectionRule
        RequirementRule = $RequirementRule
        ReturnCode = $ReturnCode
        InstallCommandLine = $InstallCommandLine
        UninstallCommandLine = $UninstallCommandLine
        Verbose = $true
    }
    
    # Add icon only if it exists
    if ($Icon) {
        $appParams.Icon = $Icon
    }
    
    $Win32App = Add-IntuneWin32App @appParams
    
    if ($Win32App) {
        Write-Host "`n================================" -ForegroundColor Green
        Write-Host "✓ SUCCESS!" -ForegroundColor Green
        Write-Host "================================" -ForegroundColor Green
        Write-Host "App Name: $($Win32App.displayName)" -ForegroundColor Cyan
        Write-Host "App ID: $($Win32App.id)" -ForegroundColor Cyan
        Write-Host "`nNext Steps:" -ForegroundColor Yellow
        Write-Host "1. Go to the Intune portal" -ForegroundColor White
        Write-Host "2. Navigate to Apps > Windows" -ForegroundColor White
        Write-Host "3. Find '$DisplayName'" -ForegroundColor White
        Write-Host "4. Assign to appropriate groups" -ForegroundColor White
    } else {
        Write-Host "`n× Failed to create Win32 app - no object returned" -ForegroundColor Red
    }
} catch {
    Write-Host "`n× Failed to create Win32 app" -ForegroundColor Red
    Write-Host "Error details: $_" -ForegroundColor Red
    Write-Host "Full error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    
    if ($_.Exception.InnerException) {
        Write-Host "Inner exception: $($_.Exception.InnerException.Message)" -ForegroundColor Red
    }
}
#endregion

Write-Host "`n================================" -ForegroundColor Cyan
Write-Host "Script completed!" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Cyan

