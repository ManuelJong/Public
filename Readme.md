# Printer & Driver Intune Win32 App Automation

## Author
Manuel Jong

## Version
2.2.0 (Public Release)

## Overview
This script automates the process of packaging, uploading, and assigning printer drivers as Win32 apps in Microsoft Intune using the IntuneWin32App PowerShell module.  
It ensures required modules are installed, connects to Intune, creates detection logic, sets up app properties, and assigns the app to Azure AD groups.

**Version 2.2.0 Enhancements (Public Release):**
- All company-specific and confidential information removed
- Documentation and script now suitable for public use
- Automatic creation of `.intunewin` package (no manual packaging needed)
- Automatic certificate extraction and installation from .cat files
- Support for Advanced Port Monitor with PrnInst.exe or DriverWizard.exe
- Intelligent detection of Microsoft vs third-party certificates
- Smart image selection for app icons (auto-selects if 1 image found)
- Intelligent printer management (updates existing, creates new based on name)
- Enhanced error handling with driver pre-existence detection
- No double installation when running in different PowerShell architectures
- Clear instructions for required Azure AD App Registration permissions

---

## Prerequisites

- **Windows PowerShell 5.1 or later**
- **IntuneWin32App PowerShell module** (the script will install if missing)
- **Azure AD App Registration** with permissions for Intune and Microsoft Graph
- **Printer driver files** (INF, CAT, DLL files) and installation scripts
- **Network access** to download IntuneWinAppUtil.exe if not present

---

## Azure AD App Registration Permissions

The Azure AD app registration used for authentication must have the following Microsoft Graph API permissions:
- DeviceManagementApps.ReadWrite.All (Application)
- DeviceManagementManagedDevices.ReadWrite.All (Application)
- Directory.Read.All (Application)
- User.Read (Delegated)
You must grant admin consent for these permissions.

---

## 🌍 Global Printer Naming Convention

Defines a consistent, global standard for printer names.  
See [Microsoft documentation](https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-app-management) for more details.

### General Format
```
<COUNTRY>-<BRAND>-<MODEL>-<BUSINESS_UNIT>-<SEQUENCE>
```

### Example
```
USA-HP-M611-HR-1
FRA-KYO-M8130-FINANCE-1
```

---

## How to Use

### 1. Prepare Your Files

Place all these files in the same folder:
- `script.ps1` (main automation script)
- `Install-Printer.ps1` (installation script)
- `Remove-Printer.ps1` (optional removal script)
- Driver INF file (e.g., `Example.inf`)
- Driver CAT files (for certificates)
- All driver DLL and supporting files
- `PrnInst.exe` or `DriverWizard.exe` (optional - for Advanced Port Monitor)
- Icon image (optional - .jpg, .jpeg, or .png)

**Note:** You do NOT need to create the `.intunewin` package manually - the script will do it automatically.

### 2. Edit Script Parameters

Open `script.ps1` and update the following variables:

```powershell
# Printer Configuration
$DisplayName = "Printer-Example-1"                # Printer name
$Publisher = "ExampleBrand"                        # Manufacturer
$ipadress = "10.5.0.60"                           # Standard TCP/IP port
# OR for Advanced Port Monitor (if PrnInst.exe or DriverWizard.exe is present):
$ipadress = "LAN_10.5.0.60"                       # Advanced port

$Drivername = "Example Printer Driver"             # Exact driver name from INF
$inffile = "Example.inf"                           # INF filename

# Azure Configuration
$TenantID = "<your-tenant-id>.onmicrosoft.com"     # Your tenant
$ClientID = "<your-app-registration-client-id>"    # Your app registration
```

### 3. Run the Script

1. Open PowerShell **as Administrator**
2. Navigate to the folder containing your files:
   ```powershell
   cd "C:\YourPrinterPackageFolder"
   ```
3. Execute the script:
   ```powershell
   .\script.ps1
   ```

### 4. What Happens Automatically

The script will:
1. **Check for `.intunewin` package** - if not found, creates it automatically
2. **Download IntuneWinAppUtil.exe** if needed (from Microsoft's GitHub)
3. **Create the package** from your printer files
4. **Connect to Intune** (you'll be prompted to authenticate)
5. **Handle app icon selection**:
   - Auto-selects if exactly 1 image found
   - Prompts for selection if 0 or multiple images
6. **Generate detection script** automatically
7. **Upload to Intune** with all configurations
8. **Display success message** with app details

---

## New Features in Version 2.2.0

### Public Release
- All company-specific and confidential information removed
- Documentation and script now suitable for public use

### Smart Image Selection
- Automatically uses single image found as app icon
- Prompts for selection if 0 or multiple images
- Supports .jpg, .jpeg, and .png formats

### Enhanced Certificate Handling
- Detects Microsoft-signed certificates and skips them (already trusted)
- Only installs third-party certificates to TrustedPublisher
- Provides detailed logging of certificate operations

### Advanced Port Monitor Support
- Supports PrnInst.exe and DriverWizard.exe for advanced port creation
- Automatically detects which tool is available
- Falls back to standard TCP/IP port if creation fails

### Improved Driver Installation
- Pre-checks if driver exists before attempting installation
- Handles pre-existing drivers gracefully
- Smart driver matching if exact name differs slightly
- No double installation when script redirects to 64-bit PowerShell

### Intelligent Printer Installation
- Creates new printer if name doesn't exist
- Updates existing printer if name exists but configuration differs
- Skips installation if printer already configured correctly
- Supports multiple printers with the same driver

---

## Installation Command Examples

**Standard Port Installation:**
```powershell
powershell.exe -executionpolicy bypass -file .\Install-Printer.ps1 -PortName "IP_10.5.0.60" -PrinterIP "10.5.0.60" -PrinterName "Printer-Example-1" -DriverName "Example Printer Driver" -INFFile "Example.inf"
```

**Advanced Port Installation (with LAN_ prefix):**
```powershell
powershell.exe -executionpolicy bypass -file .\Install-Printer.ps1 -PortName "LAN_10.5.0.60" -PrinterIP "10.5.0.60" -PrinterName "Printer-Example-1" -DriverName "Example Printer Driver" -INFFile "Example.inf"
```

**Uninstallation:**
```powershell
powershell.exe -executionpolicy bypass -file .\Remove-Printer.ps1 -PrinterName "Printer-Example-1"
```

---

## Port Naming Convention

### Standard TCP/IP Ports
Use `IP_` prefix or just the IP address:
- `IP_10.5.0.60`
- `10.5.0.60`

### Advanced Port Monitor
Use `LAN_` prefix (requires PrnInst.exe or DriverWizard.exe):
- `LAN_10.5.0.60`

The script automatically detects the prefix and uses the appropriate port type.

---

## Troubleshooting

### Common Issues

**Package creation fails:**
- Ensure all driver files are in the same folder as `script.ps1`
- Check internet connection (needed to download IntuneWinAppUtil.exe)
- Verify the INF file exists and is valid

**Certificate issues:**
- Check that .cat files are present in the package
- Microsoft certificates are automatically skipped (already trusted)
- Third-party certificates are installed to TrustedPublisher
- Check logs at `%TEMP%\[PrinterName].log` for details

**Driver installation appears to fail but works:**
- Script now checks if driver already exists before installation
- Pre-existing drivers are handled gracefully
- Installation succeeds if printer and port are created successfully

**Advanced Port not working:**
- Ensure `PrnInst.exe` or `DriverWizard.exe` is in the package folder
- Use `LAN_` prefix in the IP address
- Falls back to standard port automatically if it fails
- Check log to see which tool was attempted

**Double installation occurring:**
- Script now properly exits after 64-bit redirection
- Ensures installation only happens once

### Log Files
Installation logs are created at:
- Installation: `%SystemRoot%\Temp\[PrinterName].log`
- Removal: `%SystemRoot%\Temp\[PrinterName]_Removal.log`

Log includes:
- Certificate extraction details
- Driver staging information
- Port creation method (Standard/Advanced)
- Printer configuration steps
- Any errors encountered

### Verification
After deployment, verify on target computers:
```powershell
# Check if printer exists
Get-Printer -Name "YourPrinterName"

# Check driver
Get-PrinterDriver | Where-Object {$_.Name -like "*YourDriver*"}

# Check port (Standard or Advanced)
Get-PrinterPort | Where-Object {$_.Name -like "*YourIP*"}

# Check installed certificates
Get-ChildItem Cert:\LocalMachine\TrustedPublisher | Where-Object {$_.Subject -like "*YourVendor*"}

# View installation log
Get-Content "$env:SystemRoot\Temp\YourPrinterName.log"
```

---

## File Structure Example

```
PrinterPackage/
├── script.ps1                 # Main automation script
├── Install-Printer.ps1        # Installation script
├── Remove-Printer.ps1         # Removal script
├── Example.inf                # Driver INF file
├── example.cat                # Driver catalog file
├── *.dll                      # Driver DLL files
├── PrnInst.exe                # (Optional) Advanced Port tool
├── DriverWizard.exe           # (Optional) Alternative Advanced Port tool
└── icon.png                   # (Optional) App icon - auto-selected if only one
```

---

## Notes

- Run PowerShell as Administrator for all operations
- The script automatically handles 32-bit vs 64-bit compatibility
- Multiple printers can use the same driver (driver installed once, reused)
- Microsoft-signed drivers don't require certificate installation
- Third-party certificates are installed to TrustedPublisher store automatically
- The detection script checks registry: `HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\[PrinterName]`
- Icon selection is automatic when only one image is present

---

## Version History

- **2.2.0** (2025-08-22)
  - Public release: all company-specific information removed
  - Documentation updated for public use
- **2.1.0** (2024-08-21)
  - Added DriverWizard.exe support for Toshiba printers
  - Smart image selection (auto-select if 1 image)
  - Microsoft certificate detection and skipping
  - Fixed double installation issue with 64-bit redirection
  - Improved handling of pre-existing drivers
- **2.0.0** (2024-08-20)
  - Automatic package creation
  - Certificate handling
  - Advanced Port Monitor support
  - Intelligent printer management
- **1.0.0** (2021-12-31)
  - Initial release with manual package creation

---

## References

- [Microsoft Win32 Content Prep Tool](https://learn.microsoft.com/en-us/mem/intune/apps/apps-win32-app-management)
- [IntuneWin32App PowerShell Module](https://github.com/MSEndpointMgr/IntuneWin32App)
- [Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/overview)

---

## Support

For issues or questions:
- Check the log files in `%SystemRoot%\Temp\`
- Verify all prerequisites are met
- Ensure proper permissions in Azure AD
- Test installation locally before Intune deployment

---