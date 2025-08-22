<#
.SYNOPSIS
    Network Printer Installation Script for Intune Win32 App Deployment
    
.DESCRIPTION
    This script installs network printers with their drivers via Intune Win32 app deployment.
    It automatically extracts and installs certificates from .cat files to prevent driver 
    installation prompts, stages drivers to the Windows driver store, creates printer ports,
    and configures the printer. Supports both standard TCP/IP ports and Advanced Port Monitor
    when using the LAN_ prefix with PrnInst.exe or DriverWizard.exe.
    
.NOTES
    Original Author: Ben Whitmore
    Modified By:     Manuel Jong
    Version:         2.1.0
    Date Created:    31/12/2021
    Last Modified:   21/08/2025
    Company:         BEJO
    
.VERSION HISTORY
    1.0.0 - 31/12/2021 - Ben Whitmore - Initial release
    1.5.0 - 2024       - Added automatic certificate extraction from .cat files
    2.0.0 - 20/08/2025 - Manuel Jong - Added Advanced Port Monitor support with PrnInst.exe,
                         improved certificate handling for both Microsoft and third-party drivers,
                         enhanced logging, added LAN_ prefix detection for advanced ports
    2.1.0 - 21/08/2025 - Manuel Jong - Added DriverWizard.exe support, improved driver pre-existence handling
                         
.PARAMETER PortName
    The name of the printer port to create. 
    Use "LAN_" prefix (e.g., "LAN_10.1.1.1") to enable Advanced Port Monitor with PrnInst.exe or DriverWizard.exe
    
.PARAMETER PrinterIP
    The IP address of the network printer
    
.PARAMETER PrinterName
    The display name for the printer in Windows
    
.PARAMETER DriverName
    The exact driver name as specified in the INF file
    
.PARAMETER INFFile
    The filename of the driver INF file

#### Win32 app Commands ####

Install (Standard TCP/IP):
powershell.exe -executionpolicy bypass -file .\Install-Printer.ps1 -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf"

Install (Advanced Port Monitor):
powershell.exe -executionpolicy bypass -file .\Install-Printer.ps1 -PortName "LAN_10.5.0.40" -PrinterIP "10.5.0.40" -PrinterName "Toshiba Label Printer" -DriverName "Toshiba BV410D TS" -INFFile "Toshiba.inf"

Uninstall:
powershell.exe -executionpolicy bypass -file .\Remove-Printer.ps1 -PrinterName "Canon Printer Upstairs"

Detection:
HKEY_LOCAL_MACHINE\Software\Microsoft\Windows NT\CurrentVersion\Print\Printers\Canon Printer Upstairs
Name = "Canon Printer Upstairs"

.EXAMPLE
    .\Install-Printer.ps1 -PortName "IP_10.10.1.1" -PrinterIP "10.1.1.1" -PrinterName "Canon Printer Upstairs" -DriverName "Canon Generic Plus UFR II" -INFFile "CNLB0MA64.inf"
    
.EXAMPLE
    .\Install-Printer.ps1 -PortName "LAN_10.5.0.40" -PrinterIP "10.5.0.40" -PrinterName "TSTMDM-LAB-1C-V0-SEED-HEALTH-A-1" -DriverName "Toshiba BV410D TS" -INFFile "Toshiba.inf"
#>

[CmdletBinding()]
Param (
    [Parameter(Mandatory = $True)]
    [String]$PortName,
    [Parameter(Mandatory = $True)]
    [String]$PrinterIP,
    [Parameter(Mandatory = $True)]
    [String]$PrinterName,
    [Parameter(Mandatory = $True)]
    [String]$DriverName,
    [Parameter(Mandatory = $True)]
    [String]$INFFile
)

#Reset Error catching variable
$Throwbad = $Null
$UseAdvancedPort = $false
$AdvancedPortCreated = $false
$AdvancedPortTool = $null

#Run script in 64bit PowerShell to enumerate correct path for pnputil
If ($ENV:PROCESSOR_ARCHITEW6432 -eq "AMD64") {
    Try {
        &"$ENV:WINDIR\SysNative\WindowsPowershell\v1.0\PowerShell.exe" -File $PSCOMMANDPATH -PortName $PortName -PrinterIP $PrinterIP -DriverName $DriverName -PrinterName $PrinterName -INFFile $INFFile
        exit $LASTEXITCODE  # EXIT HERE to prevent double execution
    }
    Catch {
        Write-Error "Failed to start $PSCOMMANDPATH"
        Write-Warning "$($_.Exception.Message)"
        exit 1  # EXIT with error
    }
}

function Write-LogEntry {
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Value,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "$($PrinterName).log",
        [switch]$Stamp
    )

    #Build Log File appending System Date/Time to output
    $LogFile = Join-Path -Path $env:SystemRoot -ChildPath $("Temp\$FileName")
    $Time = -join @((Get-Date -Format "HH:mm:ss.fff"), " ", (Get-WmiObject -Class Win32_TimeZone | Select-Object -ExpandProperty Bias))
    $Date = (Get-Date -Format "MM-dd-yyyy")

    If ($Stamp) {
        $LogText = "<$($Value)> <time=""$($Time)"" date=""$($Date)"">"
    }
    else {
        $LogText = "$($Value)"   
    }
    
    Try {
        Out-File -InputObject $LogText -Append -NoClobber -Encoding Default -FilePath $LogFile -ErrorAction Stop
    }
    Catch [System.Exception] {
        Write-Warning -Message "Unable to add log entry to $LogFile.log file. Error message at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
    }
}

Write-LogEntry -Value "##################################"
Write-LogEntry -Stamp -Value "Installation started - Script Version 2.1.0"
Write-LogEntry -Value "##################################"
Write-LogEntry -Value "Install Printer using the following values..."
Write-LogEntry -Value "Port Name: $PortName"
Write-LogEntry -Value "Printer IP: $PrinterIP"
Write-LogEntry -Value "Printer Name: $PrinterName"
Write-LogEntry -Value "Driver Name: $DriverName"
Write-LogEntry -Value "INF File: $INFFile"

# Check for Advanced Port Monitor tools and determine which one to use
$PrnInstPath = Join-Path -Path $PSScriptRoot -ChildPath "PrnInst.exe"
$DriverWizardPath = Join-Path -Path $PSScriptRoot -ChildPath "DriverWizard.exe"

if ($PortName -like "LAN_*") {
    # Check which tool is available
    if (Test-Path $PrnInstPath) {
        Write-LogEntry -Stamp -Value "PrnInst.exe found and LAN_ prefix detected - Advanced Port Monitor will be used"
        $UseAdvancedPort = $true
        $AdvancedPortTool = "PrnInst"
    }
    elseif (Test-Path $DriverWizardPath) {
        Write-LogEntry -Stamp -Value "DriverWizard.exe found and LAN_ prefix detected - Toshiba Advanced Port will be used"
        $UseAdvancedPort = $true
        $AdvancedPortTool = "DriverWizard"
    }
    else {
        Write-LogEntry -Stamp -Value "LAN_ prefix detected but no Advanced Port tool found (PrnInst.exe or DriverWizard.exe)"
        Write-LogEntry -Stamp -Value "Falling back to standard TCP/IP port configuration"
    }
    
    if ($UseAdvancedPort) {
        # Extract the actual IP from the PortName (e.g., "LAN_10.1.1.1" -> "10.1.1.1")
        $ActualIP = $PortName -replace "^LAN_", ""
        
        # Verify PrinterIP matches the extracted IP or use PrinterIP if different
        if ($ActualIP -ne $PrinterIP -and $PrinterIP -notlike "LAN_*") {
            Write-LogEntry -Stamp -Value "Using IP from PrinterIP parameter: $PrinterIP"
            $ActualIP = $PrinterIP
        } elseif ($PrinterIP -like "LAN_*") {
            $ActualIP = $PrinterIP -replace "^LAN_", ""
            Write-LogEntry -Stamp -Value "Extracted IP from PrinterIP parameter: $ActualIP"
        }
        
        Write-LogEntry -Stamp -Value "Advanced Port will use IP: $ActualIP with tool: $AdvancedPortTool"
    }
} else {
    Write-LogEntry -Stamp -Value "Using standard TCP/IP port configuration"
}

# Extract and install certificates from .cat files to prevent driver installation prompts
Write-LogEntry -Stamp -Value "Checking for driver catalog files (.cat) to extract certificates..."
Write-LogEntry -Stamp -Value "Current script directory: $PSScriptRoot"

try {
    # Get all .cat files in the directory
    $CatFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.cat" -File -ErrorAction SilentlyContinue
    
    if ($CatFiles) {
        Write-LogEntry -Stamp -Value "Found $($CatFiles.Count) catalog file(s):"
        foreach ($file in $CatFiles) {
            Write-LogEntry -Stamp -Value "  - $($file.Name) (Size: $($file.Length) bytes)"
        }
        
        # Try opening certificate store
        try {
            $CertStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "LocalMachine")
            $CertStore.Open([System.Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
            Write-LogEntry -Stamp -Value "Successfully opened TrustedPublisher certificate store"
        } catch {
            Write-LogEntry -Stamp -Value "Failed to open certificate store: $_"
        }
        
        $CertsAdded = 0
        $ProcessedCerts = @()
        $MicrosoftCerts = 0
        
        foreach ($CatFile in $CatFiles) {
            Write-LogEntry -Stamp -Value "Processing catalog file: $($CatFile.Name)"
            
            try {
                # Get authenticode signature
                $AuthSig = Get-AuthenticodeSignature -FilePath $CatFile.FullName
                Write-LogEntry -Stamp -Value "  Signature Status: $($AuthSig.Status)"
                
                if ($AuthSig.SignerCertificate) {
                    $Cert = $AuthSig.SignerCertificate
                    Write-LogEntry -Stamp -Value "  Certificate Found:"
                    Write-LogEntry -Stamp -Value "    Subject: $($Cert.Subject)"
                    Write-LogEntry -Stamp -Value "    Issuer: $($Cert.Issuer)"
                    Write-LogEntry -Stamp -Value "    Thumbprint: $($Cert.Thumbprint)"
                    
                    # Check if this is a Microsoft certificate
                    if ($Cert.Subject -match "CN=Microsoft" -or 
                        $Cert.Issuer -match "CN=Microsoft" -or
                        $Cert.Subject -match "O=Microsoft Corporation") {
                        $MicrosoftCerts++
                        Write-LogEntry -Stamp -Value "  Microsoft-signed certificate detected - skipping (already trusted by Windows)"
                        continue  # Skip to next certificate
                    }
                    
                    # Check if we've already processed this certificate
                    if ($ProcessedCerts -contains $Cert.Thumbprint) {
                        Write-LogEntry -Stamp -Value "  Certificate already processed (duplicate)"
                        continue
                    }
                    
                    $ProcessedCerts += $Cert.Thumbprint
                    
                    # Check if certificate is already in store
                    $ExistingCert = Get-ChildItem -Path "cert:\LocalMachine\TrustedPublisher" -ErrorAction SilentlyContinue | 
                                    Where-Object { $_.Thumbprint -eq $Cert.Thumbprint }
                    
                    if (-not $ExistingCert) {
                        Write-LogEntry -Stamp -Value "  Third-party certificate not in TrustedPublisher store, attempting to add..."
                        
                        $certAdded = $false
                        
                        # Try to add the certificate
                        try {
                            $CertStore.Add($Cert)
                            Write-LogEntry -Stamp -Value "  Successfully added certificate via direct store method"
                            $certAdded = $true
                            $CertsAdded++
                        } catch {
                            Write-LogEntry -Stamp -Value "  Failed to add via store method: $_"
                            
                            # Try alternative method using certutil
                            try {
                                $tempCertPath = Join-Path $env:TEMP "temp_cert_$(Get-Random).cer"
                                $certBytes = $Cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Cert)
                                [System.IO.File]::WriteAllBytes($tempCertPath, $certBytes)
                                
                                $certResult = Start-Process -FilePath "certutil.exe" -ArgumentList "-addstore","TrustedPublisher","`"$tempCertPath`"" -Wait -PassThru -NoNewWindow
                                
                                if ($certResult.ExitCode -eq 0) {
                                    Write-LogEntry -Stamp -Value "  Successfully added certificate via certutil"
                                    $certAdded = $true
                                    $CertsAdded++
                                } else {
                                    Write-LogEntry -Stamp -Value "  Certutil failed with exit code: $($certResult.ExitCode)"
                                }
                                
                                # Clean up temp file
                                if (Test-Path $tempCertPath) {
                                    Remove-Item $tempCertPath -Force -ErrorAction SilentlyContinue
                                }
                            } catch {
                                Write-LogEntry -Stamp -Value "  Alternative method also failed: $_"
                            }
                        }
                    } else {
                        Write-LogEntry -Stamp -Value "  Third-party certificate already exists in TrustedPublisher store"
                    }
                    
                } elseif ($AuthSig.Status -eq "NotSigned") {
                    Write-LogEntry -Stamp -Value "  File is not signed (no certificate to extract)"
                } else {
                    Write-LogEntry -Stamp -Value "  No signer certificate found (Status: $($AuthSig.Status))"
                }
            }
            catch {
                Write-LogEntry -Stamp -Value "  ERROR processing catalog file: $_"
            }
        }
        
        # Close the certificate store
        try {
            $CertStore.Close()
        } catch {
            # Store might not have been opened successfully
        }
        
        # Summary with Microsoft cert count
        Write-LogEntry -Stamp -Value "Certificate extraction summary:"
        if ($MicrosoftCerts -gt 0) {
            Write-LogEntry -Stamp -Value "  - Microsoft certificates found (skipped): $MicrosoftCerts"
        }
        Write-LogEntry -Stamp -Value "  - Third-party certificates added: $CertsAdded"
        
        if ($CertsAdded -gt 0) {
            Write-LogEntry -Stamp -Value "Waiting 5 seconds for Windows to process the certificates..."
            Start-Sleep -Seconds 5
        }
    } else {
        Write-LogEntry -Stamp -Value "No catalog files (.cat) found in $PSScriptRoot"
    }
}
catch {
    Write-LogEntry -Stamp -Value "ERROR in certificate extraction section: $_"
    Write-LogEntry -Stamp -Value "Exception details: $($_.Exception.Message)"
}

# Also check for and install any standalone .cer files in the directory
try {
    $CerFiles = Get-ChildItem -Path $PSScriptRoot -Filter "*.cer" -File -ErrorAction SilentlyContinue
    
    if ($CerFiles) {
        Write-LogEntry -Stamp -Value "Found $($CerFiles.Count) certificate file(s) (.cer)"
        
        foreach ($CerFile in $CerFiles) {
            try {
                $certResult = Start-Process -FilePath "certutil.exe" -ArgumentList "-addstore","TrustedPublisher","`"$($CerFile.FullName)`"" -Wait -PassThru -NoNewWindow
                
                if ($certResult.ExitCode -eq 0) {
                    Write-LogEntry -Stamp -Value "Successfully added certificate file: $($CerFile.Name)"
                } else {
                    Write-LogEntry -Stamp -Value "Failed to add certificate file: $($CerFile.Name) (Exit code: $($certResult.ExitCode))"
                }
            }
            catch {
                Write-LogEntry -Stamp -Value "Error installing certificate file $($CerFile.Name): $_"
            }
        }
        
        Start-Sleep -Seconds 2
    }
}
catch {
    Write-LogEntry -Stamp -Value "Error processing .cer files: $_"
}

# Create Advanced Port Monitor if applicable
if ($UseAdvancedPort) {
    Try {
        if ($AdvancedPortTool -eq "PrnInst") {
            Write-LogEntry -Stamp -Value "Creating Advanced Port Monitor port using PrnInst.exe..."
            
            # First, pre-stage the driver
            $PreInstallArgs = "/PREINSTALL=`"$PSScriptRoot\$INFFile`" /q"
            Write-LogEntry -Stamp -Value "Pre-staging driver: PrnInst.exe $PreInstallArgs"
            
            $PreInstallResult = Start-Process -FilePath $PrnInstPath -ArgumentList $PreInstallArgs -Wait -PassThru -NoNewWindow
            
            if ($PreInstallResult.ExitCode -eq 0) {
                Write-LogEntry -Stamp -Value "Driver pre-staged successfully"
                Start-Sleep -Seconds 2
            } else {
                Write-LogEntry -Stamp -Value "Warning: Driver pre-staging returned exit code: $($PreInstallResult.ExitCode)"
            }
            
            # Create Advanced Port Monitor port
            $AdvPortArgs = @(
                "/INSTALLPORTMON",
                "`"-monitor=Advanced Port Monitor`"",
                "/name=`"$PortName`"",
                "/ip=`"$ActualIP`"",
                "/q"
            )
            
            $AdvPortArgsString = $AdvPortArgs -join " "
            Write-LogEntry -Stamp -Value "Creating port: PrnInst.exe $AdvPortArgsString"
            
            $AdvPortResult = Start-Process -FilePath $PrnInstPath -ArgumentList $AdvPortArgsString -Wait -PassThru -NoNewWindow
            
            if ($AdvPortResult.ExitCode -eq 0) {
                Write-LogEntry -Stamp -Value "Advanced Port Monitor port created successfully with PrnInst: $PortName -> $ActualIP"
                $AdvancedPortCreated = $true
                Start-Sleep -Seconds 2
            } else {
                Write-LogEntry -Stamp -Value "PrnInst port creation failed with exit code: $($AdvPortResult.ExitCode)"
                Write-LogEntry -Stamp -Value "Falling back to standard TCP/IP port"
            }
        }
        elseif ($AdvancedPortTool -eq "DriverWizard") {
            Write-LogEntry -Stamp -Value "Creating Toshiba Advanced Port using DriverWizard.exe..."
            
            # DriverWizard.exe command-line for Advanced Port
            # Common syntax for DriverWizard (may vary by version)
            $AdvPortArgs = @(
                "/IA",                    # Install Advanced port
                "/IP=`"$ActualIP`"",      # IP address
                "/PN=`"$PortName`"",      # Port name
                "/PT=9100",               # Port number (9100 for RAW)
                "/Q"                      # Quiet mode
            )
            
            $AdvPortArgsString = $AdvPortArgs -join " "
            Write-LogEntry -Stamp -Value "Creating port: DriverWizard.exe $AdvPortArgsString"
            
            $AdvPortResult = Start-Process -FilePath $DriverWizardPath -ArgumentList $AdvPortArgsString -Wait -PassThru -NoNewWindow
            
            if ($AdvPortResult.ExitCode -eq 0) {
                Write-LogEntry -Stamp -Value "Toshiba Advanced Port created successfully with DriverWizard: $PortName -> $ActualIP"
                $AdvancedPortCreated = $true
                Start-Sleep -Seconds 2
            } else {
                Write-LogEntry -Stamp -Value "DriverWizard port creation failed with exit code: $($AdvPortResult.ExitCode)"
                
                # Try alternative DriverWizard syntax (some versions use different parameters)
                Write-LogEntry -Stamp -Value "Trying alternative DriverWizard syntax..."
                
                $AltPortArgs = @(
                    "/ADDPORT",               # Alternative: Add port
                    "/PORTNAME:`"$PortName`"",
                    "/IPADDRESS:`"$ActualIP`"",
                    "/PORTNUMBER:9100",
                    "/SILENT"
                )
                
                $AltPortArgsString = $AltPortArgs -join " "
                Write-LogEntry -Stamp -Value "Alternative command: DriverWizard.exe $AltPortArgsString"
                
                $AltPortResult = Start-Process -FilePath $DriverWizardPath -ArgumentList $AltPortArgsString -Wait -PassThru -NoNewWindow
                
                if ($AltPortResult.ExitCode -eq 0) {
                    Write-LogEntry -Stamp -Value "Toshiba Advanced Port created successfully with alternative syntax"
                    $AdvancedPortCreated = $true
                    Start-Sleep -Seconds 2
                } else {
                    Write-LogEntry -Stamp -Value "Alternative syntax also failed with exit code: $($AltPortResult.ExitCode)"
                    Write-LogEntry -Stamp -Value "Falling back to standard TCP/IP port"
                }
            }
        }
    }
    Catch {
        Write-Warning "Error creating Advanced Port Monitor"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error creating Advanced Port with $AdvancedPortTool : $($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Falling back to standard TCP/IP port"
    }
}

# Validate INF file exists
$INFFullPath = Join-Path -Path $PSScriptRoot -ChildPath $INFFile
if (-not (Test-Path $INFFullPath)) {
    Write-LogEntry -Stamp -Value "ERROR: INF file not found at $INFFullPath"
    Write-Error "INF file not found: $INFFullPath"
    exit 1
} else {
    Write-LogEntry -Stamp -Value "INF file validated at: $INFFullPath"
}

# Check if driver already exists BEFORE trying to stage it
$DriverAlreadyExists = $false
try {
    $ExistingDriver = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
    if ($ExistingDriver) {
        Write-LogEntry -Stamp -Value "Driver ""$($DriverName)"" is already installed in Windows - skipping driver staging"
        $DriverAlreadyExists = $true
    }
} catch {
    # Driver doesn't exist, will need to install it
}

If (-not $ThrowBad -and -not $DriverAlreadyExists) {

    Try {
        # Build full path to INF file
        $INFFullPath = Join-Path -Path $PSScriptRoot -ChildPath $INFFile
        
        # Verify INF file exists
        if (-not (Test-Path $INFFullPath)) {
            throw "INF file not found at: $INFFullPath"
        }
        
        Write-LogEntry -Stamp -Value "Staging Driver to Windows Driver Store"
        Write-LogEntry -Stamp -Value "INF File Path: $INFFullPath"
        
        # Use full path to pnputil.exe to avoid path issues
        $pnputilPath = "$env:windir\System32\pnputil.exe"
        if (-not (Test-Path $pnputilPath)) {
            # Try sysnative for 32-bit processes on 64-bit systems
            $pnputilPath = "$env:windir\Sysnative\pnputil.exe"
        }
        
        Write-LogEntry -Stamp -Value "Using pnputil at: $pnputilPath"
        
        # Method 1: Try using Start-Process with full paths
        try {
            $pnputilArgs = @("/add-driver", "`"$INFFullPath`"", "/install")
            Write-LogEntry -Stamp -Value "Running: $pnputilPath $($pnputilArgs -join ' ')"
            
            $pnputilResult = Start-Process -FilePath $pnputilPath -ArgumentList $pnputilArgs -Wait -PassThru -NoNewWindow
            
            if ($pnputilResult.ExitCode -eq 0) {
                Write-LogEntry -Stamp -Value "Driver staged successfully (Method 1)"
            } else {
                throw "PnPUtil returned exit code: $($pnputilResult.ExitCode)"
            }
        } catch {
            Write-LogEntry -Stamp -Value "Method 1 failed: $_. Trying Method 2..."
            
            # Method 2: Try using cmd.exe to run pnputil
            try {
                $cmdArgs = "/c `"$pnputilPath`" /add-driver `"$INFFullPath`" /install"
                Write-LogEntry -Stamp -Value "Running via cmd: cmd.exe $cmdArgs"
                
                $cmdResult = Start-Process -FilePath "cmd.exe" -ArgumentList $cmdArgs -Wait -PassThru -NoNewWindow
                
                if ($cmdResult.ExitCode -eq 0) {
                    Write-LogEntry -Stamp -Value "Driver staged successfully (Method 2)"
                } else {
                    throw "CMD/PnPUtil returned exit code: $($cmdResult.ExitCode)"
                }
            } catch {
                Write-LogEntry -Stamp -Value "Method 2 failed: $_. Trying Method 3..."
                
                # Method 3: Use Invoke-Expression
                try {
                    Write-LogEntry -Stamp -Value "Running with Invoke-Expression"
                    
                    $cmd = "& `"$pnputilPath`" /add-driver `"$INFFullPath`" /install"
                    Invoke-Expression $cmd
                    
                    if ($LASTEXITCODE -eq 0) {
                        Write-LogEntry -Stamp -Value "Driver staged successfully (Method 3)"
                    } else {
                        throw "PnPUtil returned exit code: $LASTEXITCODE"
                    }
                } catch {
                    # Check one more time if driver exists now (sometimes pnputil reports error but driver is installed)
                    $DriverCheck = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
                    if ($DriverCheck) {
                        Write-LogEntry -Stamp -Value "Driver installation reported error but driver is now present - continuing"
                        $DriverAlreadyExists = $true
                    } else {
                        throw "All methods failed. Last error: $_"
                    }
                }
            }
        }

    }
    Catch {
        # Check if driver was actually installed despite the error
        $FinalDriverCheck = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
        if ($FinalDriverCheck) {
            Write-LogEntry -Stamp -Value "Driver staging reported error but driver ""$($DriverName)"" is installed - continuing"
            $DriverAlreadyExists = $true
        } else {
            Write-Warning "Error staging driver to Driver Store"
            Write-Warning "$($_.Exception.Message)"
            Write-LogEntry -Stamp -Value "Error staging driver to Driver Store"
            Write-LogEntry -Stamp -Value "Exception: $($_.Exception.Message)"
            Write-LogEntry -Stamp -Value "Exception Type: $($_.Exception.GetType().FullName)"
            
            # Additional debugging information
            Write-LogEntry -Stamp -Value "Current directory: $PSScriptRoot"
            Write-LogEntry -Stamp -Value "INF parameter: $INFFile"
            Write-LogEntry -Stamp -Value "Full INF path: $INFFullPath"
            Write-LogEntry -Stamp -Value "INF exists: $(Test-Path $INFFullPath)"
            Write-LogEntry -Stamp -Value "PnPUtil path: $pnputilPath"
            Write-LogEntry -Stamp -Value "PnPUtil exists: $(Test-Path $pnputilPath)"
            
            $ThrowBad = $True
        }
    }
}

If (-not $ThrowBad) {
    Try {
        # Check if the exact driver is already installed
        $DriverExist = Get-PrinterDriver -Name $DriverName -ErrorAction SilentlyContinue
        
        if (-not $DriverExist) {
            Write-LogEntry -Stamp -Value "Driver ""$($DriverName)"" not found, attempting to install..."
            
            # Try to add the printer driver
            try {
                Add-PrinterDriver -Name $DriverName -Confirm:$false -ErrorAction Stop
                Write-LogEntry -Stamp -Value "Printer Driver ""$($DriverName)"" installed successfully"
            }
            catch {
                # If exact name fails, check if a similar driver was installed by pnputil
                Write-LogEntry -Stamp -Value "Failed with exact name. Checking for similar drivers..."
                
                # Get all drivers and look for partial matches
                $AllDrivers = Get-PrinterDriver -ErrorAction SilentlyContinue
                $FoundDriver = $null
                
                foreach ($driver in $AllDrivers) {
                    # Check if driver name contains key parts of our target
                    if ($driver.Name -like "*Toshiba*" -or $driver.Name -like "*BV410*" -or $driver.Name -like "*BV420*") {
                        $FoundDriver = $driver.Name
                        Write-LogEntry -Stamp -Value "Found similar driver: $($driver.Name)"
                        break
                    }
                }
                
                if ($FoundDriver) {
                    Write-LogEntry -Stamp -Value "Using existing driver: ""$FoundDriver"" instead of ""$DriverName"""
                    $DriverName = $FoundDriver  # Update the driver name to use the found one
                } else {
                    # No similar driver found, throw the original error
                    throw $_
                }
            }
        }
        else {
            Write-LogEntry -Stamp -Value "Print Driver ""$($DriverName)"" already exists. Using existing driver."
        }
    }
    Catch {
        Write-Warning "Error installing Printer Driver"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error installing Printer Driver"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}

# Only create standard TCP/IP port if Advanced Port wasn't created
If (-not $ThrowBad -and -not $AdvancedPortCreated) {
    Try {

        #Create Printer Port
        $PortExist = Get-Printerport -Name $PortName -ErrorAction SilentlyContinue
        if (-not $PortExist) {
            Write-LogEntry -Stamp -Value "Adding Standard TCP/IP Port ""$($PortName)"""
            Add-PrinterPort -name $PortName -PrinterHostAddress $PrinterIP -Confirm:$false
        }
        else {
            Write-LogEntry -Stamp -Value "Port ""$($PortName)"" already exists. Skipping Printer Port installation."
        }
    }
    Catch {
        Write-Warning "Error creating Printer Port"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error creating Printer Port"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}
elseif ($AdvancedPortCreated) {
    Write-LogEntry -Stamp -Value "Using Advanced Port Monitor port: $PortName (Tool: $AdvancedPortTool)"
}

If (-not $ThrowBad) {
    Try {
        #Add Printer
        $PrinterExist = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        
        if ($PrinterExist) {
            # Printer with this name already exists - check configuration
            Write-LogEntry -Stamp -Value "Printer ""$($PrinterName)"" already exists"
            
            # Check if the existing printer has the exact same configuration
            if ($PrinterExist.DriverName -eq $DriverName -and $PrinterExist.PortName -eq $PortName) {
                # Exact match - no changes needed
                Write-LogEntry -Stamp -Value "Existing printer configuration matches desired configuration - no changes needed"
                Write-LogEntry -Stamp -Value "  Current Driver: $($PrinterExist.DriverName)"
                Write-LogEntry -Stamp -Value "  Current Port: $($PrinterExist.PortName)"
            }
            else {
                # Configuration differs - update the existing printer
                Write-LogEntry -Stamp -Value "Existing printer has different configuration - updating"
                Write-LogEntry -Stamp -Value "  Current Driver: $($PrinterExist.DriverName) -> New: $DriverName"
                Write-LogEntry -Stamp -Value "  Current Port: $($PrinterExist.PortName) -> New: $PortName"
                
                # Remove and recreate with new settings
                Write-LogEntry -Stamp -Value "Removing old printer configuration..."
                Remove-Printer -Name $PrinterName -Confirm:$false -ErrorAction Stop
                
                # Small delay to ensure removal completes
                Start-Sleep -Seconds 1
                
                Write-LogEntry -Stamp -Value "Creating printer with new configuration..."
                Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false -ErrorAction Stop
                Write-LogEntry -Stamp -Value "Printer updated successfully"
            }
        }
        else {
            # Printer with this name doesn't exist - create it
            Write-LogEntry -Stamp -Value "Adding new printer ""$($PrinterName)"""
            
            # Check if any printer exists with the same port (to avoid conflicts)
            $PortConflict = Get-Printer | Where-Object { $_.PortName -eq $PortName }
            if ($PortConflict) {
                Write-LogEntry -Stamp -Value "Note: Another printer ($($PortConflict.Name)) is using port $PortName"
                Write-LogEntry -Stamp -Value "Multiple printers can share the same port - continuing with installation"
            }
            
            # Create the new printer
            Add-Printer -Name $PrinterName -DriverName $DriverName -PortName $PortName -Confirm:$false -ErrorAction Stop
            Write-LogEntry -Stamp -Value "Printer added successfully"
        }
        
        # Final verification
        Start-Sleep -Seconds 1
        $PrinterFinal = Get-Printer -Name $PrinterName -ErrorAction SilentlyContinue
        
        if ($PrinterFinal) {
            if ($PrinterFinal.DriverName -eq $DriverName -and $PrinterFinal.PortName -eq $PortName) {
                Write-LogEntry -Stamp -Value "##################################"
                Write-LogEntry -Stamp -Value "Installation completed successfully"
                Write-LogEntry -Stamp -Value "Printer Name: $($PrinterFinal.Name)"
                Write-LogEntry -Stamp -Value "Driver: $($PrinterFinal.DriverName)"
                Write-LogEntry -Stamp -Value "Port: $($PrinterFinal.PortName)"
                Write-LogEntry -Stamp -Value "IP Address: $PrinterIP"
                Write-LogEntry -Stamp -Value "Port Type: $(if ($AdvancedPortCreated) { "Advanced Port ($AdvancedPortTool)" } else { "Standard TCP/IP" })"
                Write-LogEntry -Stamp -Value "Status: $($PrinterFinal.PrinterStatus)"
                Write-LogEntry -Value "##################################"
                
                # Log summary of all printers using this driver
                $PrintersWithSameDriver = Get-Printer | Where-Object { $_.DriverName -eq $DriverName }
                if ($PrintersWithSameDriver.Count -gt 1) {
                    Write-LogEntry -Stamp -Value "Other printers using driver '$DriverName':"
                    foreach ($p in $PrintersWithSameDriver | Where-Object { $_.Name -ne $PrinterName }) {
                        Write-LogEntry -Value "  - $($p.Name) on port $($p.PortName)"
                    }
                }
            }
            else {
                Write-Warning "Printer created but configuration mismatch detected"
                Write-LogEntry -Stamp -Value "WARNING: Printer configuration mismatch"
                Write-LogEntry -Stamp -Value "Expected Driver: $DriverName, Actual: $($PrinterFinal.DriverName)"
                Write-LogEntry -Stamp -Value "Expected Port: $PortName, Actual: $($PrinterFinal.PortName)"
                $ThrowBad = $True
            }
        }
        else {
            Write-Warning "Error creating Printer - verification failed"
            Write-LogEntry -Stamp -Value "ERROR: Printer ""$($PrinterName)"" not found after creation"
            $ThrowBad = $True
        }
    }
    Catch {
        Write-Warning "Error creating Printer"
        Write-Warning "$($_.Exception.Message)"
        Write-LogEntry -Stamp -Value "Error creating Printer"
        Write-LogEntry -Stamp -Value "$($_.Exception)"
        $ThrowBad = $True
    }
}

If ($ThrowBad) {
    Write-Error "An error was thrown during installation. Installation failed. Refer to the log file in %temp% for details"
    Write-LogEntry -Stamp -Value "Installation Failed"
    exit 1
} else {
    exit 0
}