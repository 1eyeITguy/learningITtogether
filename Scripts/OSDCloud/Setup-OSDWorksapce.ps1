#Requires -RunAsAdministrator

<#
.SYNOPSIS
    OSD.Workspace Prerequisites Installation Script
    
.DESCRIPTION
    This script installs and configures all prerequisites for OSD.Workspace development.
    Based on the official OSDeploy documentation: https://github.com/OSDeploy/OSD.Workspace/wiki
    
    IMPORTANT: Before running this script, you must set the execution policy manually:
    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
    
    Components installed:
    - PowerShell execution policy and security protocols
    - NuGet package provider
    - Trusted PowerShell Gallery repository
    - PowerShellGet and PackageManagement modules
    - PowerShell 7 with full configuration
    - Git for Windows with user configuration
    - Visual Studio Code Insiders
    - Windows ADK and Windows PE add-on
    - Microsoft Deployment Toolkit
    - OSD PowerShell modules (OSD.Workspace, platyPS, OSD, OSDCloud)
    
.NOTES
    Author: Matthew Miles
    Version: 2.2
    Requires: Administrator privileges
    
.LINK
    https://github.com/OSDeploy/OSD.Workspace/wiki
#>

# ===============================================================================
# SCRIPT INITIALIZATION
# ===============================================================================

Write-Host "+=======================================================================+" -ForegroundColor Cyan
Write-Host "|                 OSD.Workspace Prerequisites Installer                 |" -ForegroundColor Cyan
Write-Host "|                             Version 2.2                               |" -ForegroundColor Cyan
Write-Host "+=======================================================================+" -ForegroundColor Cyan
Write-Host ""

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

# Function to write section headers
function Write-SectionHeader {
    param([string]$Title)
    Write-Host ""
    Write-Host "+- $Title " -ForegroundColor Yellow -NoNewline
    Write-Host ("-" * (70 - $Title.Length)) -ForegroundColor Yellow
    Write-Host ""
}

# Function to write status messages
function Write-Status {
    param(
        [string]$Message,
        [string]$Status = "INFO"
    )
    
    $color = switch ($Status) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        "INFO" { "White" }
        "SKIP" { "Gray" }
    }
    
    $icon = switch ($Status) {
        "SUCCESS" { "[+]" }
        "WARNING" { "[!]" }
        "ERROR" { "[X]" }
        "INFO" { "[i]" }
        "SKIP" { "[-]" }
    }
    
    Write-Host "  $icon " -ForegroundColor $color -NoNewline
    Write-Host $Message -ForegroundColor $color
}

# ===============================================================================
# PREREQUISITE CHECKS
# ===============================================================================

Write-SectionHeader "System Prerequisites Check"

# Check if running as Administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Status "ERROR: This script must be run as Administrator!" "ERROR"
    Write-Status "Please run PowerShell as Administrator and try again." "ERROR"
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
} else {
    Write-Status "Running as Administrator - OK" "SUCCESS"
}

# Check PowerShell version
$psVersion = $PSVersionTable.PSVersion
if ($psVersion.Major -lt 7) {
    Write-Status "PowerShell $($psVersion.ToString()) detected - Installing PowerShell 7..." "WARNING"
    Write-Status "Installing PowerShell 7 with full configuration..." "INFO"
    
    try {
        $ps7InstallCommand = 'winget install -e --id Microsoft.PowerShell --accept-source-agreements --override "/passive ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1 ADD_PATH=1"'
        # Write-Status "Running: $ps7InstallCommand" "INFO"
        $process = Start-Process -FilePath "winget" -ArgumentList "install -e --id Microsoft.PowerShell --accept-source-agreements --override `"/passive ADD_EXPLORER_CONTEXT_MENU_OPENPOWERSHELL=1 ADD_FILE_CONTEXT_MENU_RUNPOWERSHELL=1 ENABLE_PSREMOTING=1 REGISTER_MANIFEST=1 USE_MU=1 ENABLE_MU=1 ADD_PATH=1`"" -Wait -PassThru -NoNewWindow
        
        if ($process.ExitCode -eq 0) {
            Write-Status "PowerShell 7 installed successfully" "SUCCESS"
            Write-Status "Continuing with module installations to PowerShell 7 location..." "INFO"
        } else {
            Write-Status "PowerShell 7 installation failed with exit code: $($process.ExitCode)" "ERROR"
            throw "PowerShell 7 installation failed"
        }
    } catch {
        Write-Status "Failed to install PowerShell 7: $($_.Exception.Message)" "ERROR"
        throw
    }
} else {
    Write-Status "Running PowerShell $($psVersion.ToString()) - OK" "SUCCESS"
}

# Check execution policy
$currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
Write-Status "Current execution policy (CurrentUser scope): $currentPolicy" "INFO"

# Collect Git user information upfront if Git is not installed or not configured
$gitUserEmail = $null
$gitUserName = $null

$gitInstalled = Get-Command git -ErrorAction SilentlyContinue
if ($gitInstalled) {
    $currentEmail = git config --global user.email 2>$null
    $currentName = git config --global user.name 2>$null
    
    # Only ask for email if it's not configured or has default values
    if ([string]::IsNullOrWhiteSpace($currentEmail) -or $currentEmail -eq "you@example.com") {
        Write-Host ""
        Write-Host "Git is installed but email is not configured." -ForegroundColor Yellow
        Write-Host "Please enter your email address: " -ForegroundColor Yellow -NoNewline
        $gitUserEmail = Read-Host
    } else {
        Write-Status "Git email already configured: $currentEmail" "SKIP"
    }
    
    # Only ask for name if it's not configured or has default values
    if ([string]::IsNullOrWhiteSpace($currentName) -or $currentName -eq "Your Name") {
        Write-Host "Please enter your name: " -ForegroundColor Yellow -NoNewline
        $gitUserName = Read-Host
    } else {
        Write-Status "Git user name already configured: $currentName" "SKIP"
    }
} else {
    Write-Host ""
    Write-Host "Git is not installed. Please provide your Git configuration:" -ForegroundColor Yellow
    Write-Host "Please enter your email address: " -ForegroundColor Yellow -NoNewline
    $gitUserEmail = Read-Host
    Write-Host "Please enter your name: " -ForegroundColor Yellow -NoNewline
    $gitUserName = Read-Host
}

# ===============================================================================
# POWERSHELL EXECUTION POLICY AND SECURITY SETUP
# ===============================================================================

Write-SectionHeader "PowerShell Security Configuration"

try {
    # Note: The execution policy should already be set by the user before running this script
    $currentPolicy = Get-ExecutionPolicy -Scope CurrentUser
    Write-Status "Current execution policy (CurrentUser scope): $currentPolicy" "INFO"
    
    # Check and set security protocol
    $currentProtocol = [Net.ServicePointManager]::SecurityProtocol
    if ($currentProtocol -notmatch 'Tls12') {
        Write-Status "Configuring TLS 1.2 security protocol..." "INFO"
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Write-Status "TLS 1.2 protocol configured successfully" "SUCCESS"
    } else {
        Write-Status "TLS 1.2 protocol already configured" "SKIP"
    }
} catch {
    Write-Status "Failed to configure PowerShell security: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# NUGET PACKAGE PROVIDER INSTALLATION
# ===============================================================================

Write-SectionHeader "NuGet Package Provider"

try {
    $nugetProvider = Get-PackageProvider -Name 'NuGet' -ErrorAction SilentlyContinue
    if (-not $nugetProvider) {
        Write-Status "Installing NuGet package provider..." "INFO"
        Install-PackageProvider -Name NuGet -ForceBootstrap -Verbose
        Write-Status "NuGet package provider installed successfully" "SUCCESS"
    } else {
        Write-Status "NuGet package provider already installed (Version: $($nugetProvider.Version))" "SKIP"
    }
} catch {
    Write-Status "Failed to install NuGet package provider: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# POWERSHELL GALLERY REPOSITORY CONFIGURATION
# ===============================================================================

Write-SectionHeader "PowerShell Gallery Repository"

try {
    $psGallery = Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue
    if ($psGallery -and $psGallery.InstallationPolicy -ne 'Trusted') {
        Write-Status "Setting PowerShell Gallery as trusted repository..." "INFO"
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
        Write-Status "PowerShell Gallery (PSGallery) has been set to Trusted" "SUCCESS"
    } elseif ($psGallery) {
        Write-Status "PowerShell Gallery (PSGallery) is already Trusted" "SKIP"
    } else {
        Write-Status "PowerShell Gallery repository not found" "WARNING"
    }
} catch {
    Write-Status "Failed to configure PowerShell Gallery: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# POWERSHELLGET AND PACKAGEMANAGEMENT MODULES
# ===============================================================================

Write-SectionHeader "PowerShell Core Modules"

$ps7ModulesPath = "C:\Program Files\PowerShell\7\Modules"
if (!(Test-Path $ps7ModulesPath)) {
    Write-Status "PowerShell 7 Modules path not found - creating directory" "WARNING"
    New-Item -Path $ps7ModulesPath -ItemType Directory -Force | Out-Null
}

try {
    # Check and install PowerShellGet to PS7
    $psGetPath = Join-Path $ps7ModulesPath "PowerShellGet"
    if (!(Test-Path $psGetPath)) {
        Write-Status "Installing PowerShellGet module to PowerShell 7..." "INFO"
        $tempPath = Join-Path $env:TEMP "PSModule_PowerShellGet"
        if (Test-Path $tempPath) { Remove-Item $tempPath -Recurse -Force }
        Save-Module -Name PowerShellGet -Path $env:TEMP -Repository PSGallery -Force
        $savedPath = Join-Path $env:TEMP "PowerShellGet"
        if (Test-Path $savedPath) {
            Move-Item -Path $savedPath -Destination $ps7ModulesPath -Force
            Write-Status "PowerShellGet module installed to PowerShell 7 successfully" "SUCCESS"
        }
    } else {
        Write-Status "PowerShellGet module already installed in PowerShell 7" "SKIP"
    }
    
    # Check and install PackageManagement to PS7
    $pkgMgmtPath = Join-Path $ps7ModulesPath "PackageManagement"
    if (!(Test-Path $pkgMgmtPath)) {
        Write-Status "Installing PackageManagement module to PowerShell 7..." "INFO"
        $tempPath = Join-Path $env:TEMP "PSModule_PackageManagement"
        if (Test-Path $tempPath) { Remove-Item $tempPath -Recurse -Force }
        Save-Module -Name PackageManagement -Path $env:TEMP -Repository PSGallery -Force
        $savedPath = Join-Path $env:TEMP "PackageManagement"
        if (Test-Path $savedPath) {
            Move-Item -Path $savedPath -Destination $ps7ModulesPath -Force
            Write-Status "PackageManagement module installed to PowerShell 7 successfully" "SUCCESS"
        }
    } else {
        Write-Status "PackageManagement module already installed in PowerShell 7" "SKIP"
    }
} catch {
    Write-Status "Failed to install PowerShell core modules: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# GIT FOR WINDOWS INSTALLATION AND CONFIGURATION
# ===============================================================================

Write-SectionHeader "Git for Windows Installation and Configuration"

try {
    # Check if Git is already installed
    $gitInstalled = Get-Command git -ErrorAction SilentlyContinue
    if ($gitInstalled) {
        $gitVersion = git --version
        Write-Status "Git already installed: $gitVersion" "SKIP"
    } else {
        Write-Status "Installing Git for Windows..." "INFO"
        winget install --id Git.Git -e -h --accept-source-agreements
        Write-Status "Git for Windows installed successfully" "SUCCESS"
        
        # Refresh environment variables
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    }
    
    # Configure Git user identity using collected information
    if ($gitUserEmail) {
        Write-Status "Configuring Git user email: $gitUserEmail" "INFO"
        git config --global user.email "$gitUserEmail"
        Write-Status "Git user email configured successfully" "SUCCESS"
    } else {
        $currentEmail = git config --global user.email 2>$null
        if ($currentEmail) {
            Write-Status "Git user email already configured: $currentEmail" "SKIP"
        }
    }
    
    if ($gitUserName) {
        Write-Status "Configuring Git user name: $gitUserName" "INFO"
        git config --global user.name "$gitUserName"
        Write-Status "Git user name configured successfully" "SUCCESS"
    } else {
        $currentName = git config --global user.name 2>$null
        if ($currentName) {
            Write-Status "Git user name already configured: $currentName" "SKIP"
        }
    }
} catch {
    Write-Status "Failed to install/configure Git: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# VISUAL STUDIO CODE INSIDERS INSTALLATION
# ===============================================================================

Write-SectionHeader "Visual Studio Code Insiders Installation"

try {
    # Check if VS Code Insiders is already installed
    $vscodeInsiders = Get-Command "code-insiders" -ErrorAction SilentlyContinue
    if ($vscodeInsiders) {
        Write-Status "Visual Studio Code Insiders already installed" "SKIP"
    } else {
        Write-Status "Installing Visual Studio Code Insiders..." "INFO"
        $vscodeArgs = '/SILENT /mergetasks="!runcode,addcontextmenufiles,addcontextmenufolders,associatewithfiles,addtopath"'
        winget install -e --id Microsoft.VisualStudioCode.Insiders --accept-source-agreements --override $vscodeArgs
        Write-Status "Visual Studio Code Insiders installed successfully" "SUCCESS"
    }
} catch {
    Write-Status "Failed to install Visual Studio Code Insiders: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# WINDOWS ADK INSTALLATION
# ===============================================================================

Write-SectionHeader "Windows Assessment and Deployment Kit (ADK)"

try {
    # Check if Windows ADK is already installed
    $adkPath = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit"
    if (Test-Path $adkPath) {
        Write-Status "Windows ADK already installed" "SKIP"
    } else {
        Write-Status "Downloading and installing Windows ADK..." "INFO"
        $adkUrl = 'https://go.microsoft.com/fwlink/?linkid=2289980'
        $adkSetupPath = "$env:TEMP\adksetup.exe"
        
        # Download ADK setup
        Invoke-WebRequest -Uri $adkUrl -OutFile $adkSetupPath -UseBasicParsing
        
        # Install ADK with required features
        $adkArgs = @('/features', 'OptionId.DeploymentTools', 'OptionId.ImagingAndConfigurationDesigner', '/quiet', '/ceip', 'off', '/norestart')
        $adkProcess = Start-Process -FilePath $adkSetupPath -ArgumentList $adkArgs -Wait -PassThru
        
        if ($adkProcess.ExitCode -eq 0) {
            Write-Status "Windows ADK installed successfully" "SUCCESS"
        } else {
            Write-Status "Windows ADK installation completed with exit code: $($adkProcess.ExitCode)" "WARNING"
        }
        
        # Clean up
        Remove-Item $adkSetupPath -Force -ErrorAction SilentlyContinue
    }
} catch {
    Write-Status "Failed to install Windows ADK: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# WINDOWS PE ADD-ON INSTALLATION
# ===============================================================================

Write-SectionHeader "Windows PE Add-on for Windows ADK"

try {
    # Check if Windows PE add-on is already installed
    $winpePath = "C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment"
    if (Test-Path $winpePath) {
        Write-Status "Windows PE add-on already installed" "SKIP"
    } else {
        Write-Status "Downloading and installing Windows PE add-on..." "INFO"
        $winpeUrl = 'https://go.microsoft.com/fwlink/?linkid=2289981'
        $winpeSetupPath = "$env:TEMP\adkwinpesetup.exe"
        
        # Download Windows PE add-on setup
        Invoke-WebRequest -Uri $winpeUrl -OutFile $winpeSetupPath -UseBasicParsing
        
        # Install Windows PE add-on
        $winpeArgs = @('/features', 'OptionId.WindowsPreinstallationEnvironment', '/quiet', '/ceip', 'off', '/norestart')
        $winpeProcess = Start-Process -FilePath $winpeSetupPath -ArgumentList $winpeArgs -Wait -PassThru
        
        if ($winpeProcess.ExitCode -eq 0) {
            Write-Status "Windows PE add-on installed successfully" "SUCCESS"
        } else {
            Write-Status "Windows PE add-on installation completed with exit code: $($winpeProcess.ExitCode)" "WARNING"
        }
        
        # Clean up
        Remove-Item $winpeSetupPath -Force -ErrorAction SilentlyContinue
    }
} catch {
    Write-Status "Failed to install Windows PE add-on: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# MDT WINPE X86 BUGFIX
# ===============================================================================

Write-SectionHeader "MDT WinPE x86 Bugfix"

try {
    $winpeOCsPath = 'C:\Program Files (x86)\Windows Kits\10\Assessment and Deployment Kit\Windows Preinstallation Environment\x86\WinPE_OCs'
    if (Test-Path $winpeOCsPath) {
        Write-Status "WinPE x86 OCs directory already exists" "SKIP"
    } else {
        Write-Status "Creating WinPE x86 OCs directory for MDT compatibility..." "INFO"
        New-Item -Path $winpeOCsPath -ItemType Directory -Force | Out-Null
        Write-Status "WinPE x86 OCs directory created successfully" "SUCCESS"
    }
} catch {
    Write-Status "Failed to create WinPE x86 OCs directory: $($_.Exception.Message)" "ERROR"
    throw
}

# ===============================================================================
# MICROSOFT DEPLOYMENT TOOLKIT INSTALLATION
# ===============================================================================

<##
Write-SectionHeader "Microsoft Deployment Toolkit (MDT)"

try {
    # Check if MDT is already installed
    $mdtPath = "${env:ProgramFiles}\Microsoft Deployment Toolkit"
    if (Test-Path $mdtPath) {
        Write-Status "Microsoft Deployment Toolkit already installed" "SKIP"
    } else {
        Write-Status "Installing Microsoft Deployment Toolkit..." "INFO"
        winget install --id Microsoft.DeploymentToolkit -e --accept-source-agreements --accept-package-agreements
        Write-Status "Microsoft Deployment Toolkit installed successfully" "SUCCESS"
    }
} catch {
    Write-Status "Failed to install Microsoft Deployment Toolkit: $($_.Exception.Message)" "ERROR"
    throw
}
#>

# ===============================================================================
# OSD POWERSHELL MODULES INSTALLATION
# ===============================================================================

Write-SectionHeader "OSD PowerShell Modules"

$ps7ModulesPath = "C:\Program Files\PowerShell\7\Modules"
if (!(Test-Path $ps7ModulesPath)) {
    Write-Status "PowerShell 7 Modules path not found - creating directory" "WARNING"
    New-Item -Path $ps7ModulesPath -ItemType Directory -Force | Out-Null
}

$modules = @(
    @{ Name = "OSD.Workspace"; Description = "The main OSDWorkspace PowerShell Module" },
    @{ Name = "platyPS"; Description = "Required for creating OSDWorkspace help files" },
    @{ Name = "OSD"; Description = "Used in some of the OSDWorkspace functions" },
    @{ Name = "OSDCloud"; Description = "Optionally used in some of the OSDWorkspace Gallery functions" }
)

foreach ($module in $modules) {
    try {
        # Check if module is already installed in PS7 location
        $modulePath = Join-Path $ps7ModulesPath $module.Name
        if (Test-Path $modulePath) {
            Write-Status "$($module.Name) already installed in PowerShell 7" "SKIP"
        } else {
            Write-Status "Installing $($module.Name) module to PowerShell 7 - $($module.Description)..." "INFO"
            
            # Save module to temp location
            $tempPath = Join-Path $env:TEMP "PSModule_$($module.Name)"
            if (Test-Path $tempPath) {
                Remove-Item $tempPath -Recurse -Force
            }
            
            Save-Module -Name $module.Name -Path $env:TEMP -Repository PSGallery -Force
            
            # Move to PS7 modules location
            $savedModulePath = Join-Path $env:TEMP $module.Name
            if (Test-Path $savedModulePath) {
                Move-Item -Path $savedModulePath -Destination $ps7ModulesPath -Force
                Write-Status "$($module.Name) module installed to PowerShell 7 successfully" "SUCCESS"
            } else {
                Write-Status "Failed to save $($module.Name) module" "ERROR"
            }
        }
    } catch {
        Write-Status "Failed to install $($module.Name) module: $($_.Exception.Message)" "ERROR"
        throw
    }
}

# ===============================================================================
# SCRIPT COMPLETION
# ===============================================================================

Write-Host ""
Write-Host "+=======================================================================+" -ForegroundColor Green
Write-Host "|                    INSTALLATION COMPLETED SUCCESSFULLY!               |" -ForegroundColor Green
Write-Host "+=======================================================================+" -ForegroundColor Green
Write-Host ""

Write-Status "All OSDCloud Workspace prerequisites have been installed successfully!" "SUCCESS"
Write-Status "You can now proceed with creating your OSD workspace." "INFO"
Write-Host ""

Write-Host "NEXT STEPS:" -ForegroundColor Cyan
Write-Host "----------" -ForegroundColor Cyan
Write-Host ""
Write-Host "1. Open PowerShell 7 and run:" -ForegroundColor Yellow
Write-Host "   Install-OSDWorkspace" -ForegroundColor Cyan
Write-Host "   This will create your OSD workspace directory." -ForegroundColor White
Write-Host ""
Write-Host "2. Open the workspace in VS Code Insiders:" -ForegroundColor Yellow
Write-Host "   cd C:\OSDWorkspace" -ForegroundColor Cyan
Write-Host "   code-insiders ." -ForegroundColor Cyan
Write-Host "   This will open your workspace in VS Code Insiders." -ForegroundColor White
Write-Host ""
Write-Host "3. Your OSD workspace will be located at:" -ForegroundColor Yellow
Write-Host "   C:\OSDWorkspace" -ForegroundColor Cyan
Write-Host ""

Write-Host ""
Write-Host "Script execution completed at $(Get-Date)" -ForegroundColor Gray
