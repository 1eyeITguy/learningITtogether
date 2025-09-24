#Requires -Version 7.0

<#
.SYNOPSIS
    Creates new PowerShell Application Deployment Toolkit (PSADT) v4 deployments with automated setup.

.DESCRIPTION
    This script automates the creation of PSADT v4 deployment packages by:
    - Installing/updating the PSAppDeployToolkit module to the latest version
    - Creating a new deployment template with user-specified application details
    - Downloading and replacing default asset files (AppIcon.png, Banner.Classic.png) if URLs provided
    - Updating deployment script parameters (vendor, name, version, author, date) when values configured
    - Configuring the deployment settings in config.psd1 (company name, log paths) when values configured
    - Optionally opening the deployment folder in Explorer and VS Code
    
    All configuration variables at the top of the script are optional - leave empty to use template defaults.

.PARAMETER None
    This script does not accept command-line parameters. Configuration is done via variables at the top of the script.

.EXAMPLE
    .\New-PSADTDeployment.ps1
    
    Runs the script interactively, prompting for application vendor, name, and version.
    Uses the configured settings for company name, asset URLs, and log location.

.NOTES
    File Name      : New-PSADTDeployment.ps1
    Author         : Matthew Miles
    Prerequisite   : PowerShell 7.0 or later
    Copyright      : Free to use and modify
    
    Configuration Variables:
    - $AppIconUrl: URL for custom application icon (leave empty for template default)
    - $BannerClassicUrl: URL for custom deployment banner (leave empty for template default)
    - $CompanyName: Company name to set in config.psd1 (leave empty for template default)
    - $Author: Script author name (leave empty for template default)
    - $LogPath: Log path to set in config.psd1 (leave empty for template default)
    - $DefaultAppArch: Default architecture - "x64", "x86", or empty for template default

.LINK
    https://github.com/PSAppDeployToolkit/PSAppDeployToolkit

#>


# ==========================
# User-configurable settings
# ==========================

# URLs for asset downloads (leave empty for template default)
# Supports both internet URLs and local file paths (UNC/file server)
# Note: For security, only HTTPS URLs are allowed for internet downloads
$AppIconUrl = ""
$BannerClassicUrl = ""

# Company name to write into config.psd1 (leave empty for template default)
$CompanyName = ""

# Script author name (leave empty for template default)
$Author = ""

# Log path to set in config.psd1 (leave empty for template default)
# For Intune logging use: C:\ProgramData\Microsoft\IntuneManagementExtension\Logs
$LogPath = ""

# Default application architecture (x86, x64, or leave empty for template default)
$DefaultAppArch = ""

# ==========================


# Function to write colored output with icons
function Write-StepOutput {
    param(
        [string]$Message,
        [string]$Status = "Info", # Info, Success, Warning, Error, Progress
        [switch]$NoNewline
    )
    
    $icons = @{
        "Info" = "ℹ️"
        "Success" = "✅"
        "Warning" = "⚠️"
        "Error" = "❌"
        "Progress" = "🔄"
        "Download" = "⬇️"
        "Config" = "⚙️"
        "Check" = "🔍"
    }
    
    $colors = @{
        "Info" = "Cyan"
        "Success" = "Green"
        "Warning" = "Yellow"
        "Error" = "Red"
        "Progress" = "Magenta"
        "Download" = "Blue"
        "Config" = "DarkYellow"
        "Check" = "DarkCyan"
    }
    
    $icon = $icons[$Status]
    $color = $colors[$Status]
    
    if ($NoNewline) {
        Write-Host "$icon $Message" -ForegroundColor $color -NoNewline
    } else {
        Write-Host "$icon $Message" -ForegroundColor $color
    }
}

# Function to create a separator line
function Write-Separator {
    Write-Host "`n" + ("=" * 80) + "`n" -ForegroundColor DarkGray
}

# Function to update script parameters efficiently
function Update-ScriptParameter {
    param(
        [string]$Content,
        [string]$ParameterName,
        [string]$NewValue,
        [string]$DisplayName = $ParameterName,
        [switch]$AllowEmpty
    )
    
    # Don't process if NewValue is empty or null (unless AllowEmpty is specified)
    if ([string]::IsNullOrWhiteSpace($NewValue) -and -not $AllowEmpty) {
        Write-StepOutput "Skipping $DisplayName update - no value provided" -Status "Info"
        return $Content
    }
    
    # Handle special case for empty string parameters (like AppArch)
    if ($AllowEmpty -and [string]::IsNullOrWhiteSpace($NewValue)) {
        Write-StepOutput "$DisplayName variable is empty - leaving template default" -Status "Info"
        return $Content
    }
    
    # Handle both empty strings and regular values
    $regex = if ($ParameterName -eq "AppArch") {
        "AppArch\s*=\s*(['`"]['`"]|[`"'][^`"']*[`"'])"
    } else {
        "$ParameterName\s*=\s*[`"'][^`"']*[`"']"
    }
    
    if ($Content -match $regex) {
        $updatedContent = $Content -replace $regex, "$ParameterName = '$NewValue'"
        Write-StepOutput "Updated $DisplayName to '$NewValue'" -Status "Success"
        return $updatedContent
    } else {
        Write-StepOutput "Warning: $DisplayName parameter not found in file" -Status "Warning"
        return $Content
    }
}

# Function to get validated user input
function Get-ValidatedInput {
    param(
        [string]$Prompt,
        [string]$ErrorMessage = "Input cannot be empty!"
    )
    
    do {
        $userInput = Read-Host $Prompt
        if ([string]::IsNullOrWhiteSpace($userInput)) {
            Write-StepOutput $ErrorMessage -Status "Error"
        }
    } while ([string]::IsNullOrWhiteSpace($userInput))
    
    # Sanitize input: trim and escape quotes for safety
    return $userInput.Trim().Replace("'", "''").Replace('"', '""')
}

# Function to wait for folder structure creation
function Wait-ForFolderStructure {
    param(
        [string]$DeploymentPath,
        [string[]]$RequiredFolders = @("Assets", "Config", "Files"),
        [int]$MaxAttempts = 10
    )
    
    $attempt = 0
    do {
        $attempt++
        Start-Sleep -Milliseconds 500  # Reduced from 1 second for faster checks
        
        $allFoldersExist = $true
        foreach ($folder in $RequiredFolders) {
            $folderPath = Join-Path $DeploymentPath $folder
            if (!(Test-Path $folderPath)) {
                $allFoldersExist = $false
                break
            }
        }
        
        if ($allFoldersExist) {
            Write-StepOutput "Folder structure verification complete" -Status "Success"
            return $true
        } elseif ($attempt -eq $MaxAttempts) {
            Write-StepOutput "Warning: Some folders may not have been created yet" -Status "Warning"
            return $false
        } else {
            Write-StepOutput "Waiting for folder structure... (attempt $attempt/$MaxAttempts)" -Status "Progress"
        }
        
    } while ($attempt -lt $MaxAttempts)
}

# Function to get summary value with fallback
function Get-SummaryValue {
    param(
        [string]$Value,
        [string]$DefaultText = "Template default"
    )
    
    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $DefaultText
    } else {
        return $Value
    }
}

# Function to process asset downloads
function Update-AssetFiles {
    param(
        [string]$AssetsPath,
        [hashtable]$AssetMappings
    )
    
    if ($AssetMappings.Count -eq 0) {
        Write-StepOutput "No custom asset URLs configured - using template defaults" -Status "Info"
        return
    }
    
    Write-StepOutput "Processing custom asset files..." -Status "Download"
    
    # Verify Assets folder exists
    if (!(Test-Path $AssetsPath)) {
        Write-StepOutput "Assets folder not found. Creating directory..." -Status "Warning"
        New-Item -Path $AssetsPath -ItemType Directory -Force | Out-Null
    }

    foreach ($file in $AssetMappings.GetEnumerator()) {
        try {
            $filePath = Join-Path $AssetsPath $file.Key
            $source = $file.Value
            
            # Check if source is a local file path (supports UNC/file server)
            if (Test-Path $source) {
                Write-StepOutput "Copying local file $($file.Key)..." -Status "Download"
                Copy-Item -Path $source -Destination $filePath -Force
                Write-StepOutput "Successfully copied $($file.Key)" -Status "Success"
            } elseif ($source -match '^https?://') {
                # Handle internet URLs with basic security check
                if ($source -notmatch '^https://') {
                    Write-StepOutput "Warning: Skipping insecure URL for $($file.Key)" -Status "Warning"
                    continue
                }
                Write-StepOutput "Downloading $($file.Key)..." -Status "Download"
                Invoke-WebRequest -Uri $source -OutFile $filePath -UseBasicParsing
                Write-StepOutput "Successfully downloaded $($file.Key)" -Status "Success"
            } else {
                Write-StepOutput "Warning: Invalid source for $($file.Key) - must be a valid URL or local path" -Status "Warning"
            }
        } catch {
            Write-StepOutput "Warning: Could not process $($file.Key) - $($_.Exception.Message)" -Status "Warning"
        }
    }
}

# Function to update deployment script parameters
function Update-DeploymentScript {
    param(
        [string]$DeploymentScriptPath,
        [string]$AppVendor,
        [string]$AppName,
        [string]$AppVersion,
        [string]$Author,
        [string]$DefaultAppArch
    )
    
    try {
        if (!(Test-Path $DeploymentScriptPath)) {
            Write-StepOutput "Warning: Deployment script not found at $DeploymentScriptPath" -Status "Warning"
            return
        }
        
        Write-StepOutput "Found deployment script, updating app details..." -Status "Config"
        
        # Read the current deployment script file
        $scriptContent = Get-Content $DeploymentScriptPath -Raw
        
        # Get current date for script metadata
        $currentDate = Get-Date -Format "yyyy-MM-dd"
        
        # Update all parameters using the helper function
        $scriptContent = Update-ScriptParameter $scriptContent "AppVendor" $AppVendor
        $scriptContent = Update-ScriptParameter $scriptContent "AppName" $AppName  
        $scriptContent = Update-ScriptParameter $scriptContent "AppVersion" $AppVersion
        $scriptContent = Update-ScriptParameter $scriptContent "AppScriptDate" $currentDate "AppScriptDate"
        $scriptContent = Update-ScriptParameter $scriptContent "AppScriptAuthor" $Author "AppScriptAuthor"
        $scriptContent = Update-ScriptParameter $scriptContent "AppArch" $DefaultAppArch -AllowEmpty
        
        # Write the updated content back to the file
        Set-Content -Path $DeploymentScriptPath -Value $scriptContent -Encoding UTF8
        Write-StepOutput "Deployment script updated successfully" -Status "Success"
        
    } catch {
        Write-StepOutput "Error updating deployment script: $($_.Exception.Message)" -Status "Error"
    }
}

# Function to update configuration file
function Update-ConfigFile {
    param(
        [string]$ConfigPath,
        [string]$CompanyName,
        [string]$LogPath
    )
    
    # Verify Config folder and file exist before attempting modification
    $configDir = Split-Path $ConfigPath -Parent
    if (!(Test-Path $configDir)) {
        Write-StepOutput "Config folder not found. Creating directory..." -Status "Warning"
        New-Item -Path $configDir -ItemType Directory -Force | Out-Null
    }

    try {
        if (!(Test-Path $ConfigPath)) {
            Write-StepOutput "Warning: Config file not found at $ConfigPath" -Status "Warning"
            Write-StepOutput "This may be normal if the template structure is different than expected" -Status "Info"
            return
        }
        
        Write-StepOutput "Found config file, updating settings..." -Status "Config"
        
        # Read the current config file
        $configContent = Get-Content $ConfigPath -Raw
        
        # Update CompanyName and LogPath using the helper function
        $configContent = Update-ScriptParameter $configContent "CompanyName" $CompanyName
        
        # Update both log path parameters if LogPath is configured
        if (![string]::IsNullOrWhiteSpace($LogPath)) {
            Write-StepOutput "Updating log paths to custom location..." -Status "Config"
            $configContent = Update-ScriptParameter $configContent "LogPath" $LogPath
            $configContent = Update-ScriptParameter $configContent "LogPathnoAdminRights" $LogPath
        } else {
            Write-StepOutput "LogPath variable is empty - leaving template defaults" -Status "Info"
        }
        
        # Write the updated content back to the file
        Set-Content -Path $ConfigPath -Value $configContent -Encoding UTF8
        Write-StepOutput "Configuration file updated successfully" -Status "Success"
        
    } catch {
        Write-StepOutput "Error updating configuration file: $($_.Exception.Message)" -Status "Error"
    }
}

# Clear screen and show header
Clear-Host
Write-Host @"
╔═══════════════════════════════════════════════════════════════════════════╗
║                                                                           ║
║                     PSADT Deployment Creator Script                       ║
║                                                                           ║
╚═══════════════════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Green

Write-Separator

# Step 1: Check and manage PSADT module
Write-StepOutput "Checking PSAppDeployToolkit module installation..." -Status "Check"

try {
    # Check if module is installed and get latest available version
    $installedModule = Get-Module -ListAvailable -Name "PSAppDeployToolkit" | Sort-Object Version -Descending | Select-Object -First 1
    $latestModule = Find-Module -Name "PSAppDeployToolkit" -ErrorAction Stop
    
    if ($installedModule) {
        Write-StepOutput "Found installed version: $($installedModule.Version)" -Status "Info"
        Write-StepOutput "Latest available version: $($latestModule.Version)" -Status "Info"
        
        if ($installedModule.Version -lt $latestModule.Version) {
            Write-StepOutput "Updating PSAppDeployToolkit module to latest version..." -Status "Progress"
            Install-Module -Name "PSAppDeployToolkit" -Force -Scope CurrentUser -AllowClobber
            Write-StepOutput "Successfully updated to version $($latestModule.Version)" -Status "Success"
        } else {
            Write-StepOutput "PSAppDeployToolkit is already up to date" -Status "Success"
        }
    } else {
        Write-StepOutput "PSAppDeployToolkit not found. Installing latest version..." -Status "Progress"
        Install-Module -Name "PSAppDeployToolkit" -Force -Scope CurrentUser
        Write-StepOutput "Successfully installed PSAppDeployToolkit" -Status "Success"
    }
    
    # Import the module
    Import-Module PSAppDeployToolkit -Force
    $currentModule = Get-Module -Name "PSAppDeployToolkit"
    $moduleVersion = $currentModule.Version.ToString()
    Write-StepOutput "Module loaded successfully - Version: $moduleVersion" -Status "Success"
    
} catch {
    Write-StepOutput "Error managing PSAppDeployToolkit module: $($_.Exception.Message)" -Status "Error"
    Write-Host "Press any key to exit..." -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Separator

# Step 2: Get user input for deployment details
Write-StepOutput "Gathering deployment information..." -Status "Info"

# Get application details using the validation function
$appVendor = Get-ValidatedInput "Enter the application vendor (e.g., Microsoft, Adobe, Cisco)"
$appName = Get-ValidatedInput "Enter the application name"
$appVersion = Get-ValidatedInput "Enter the application version"

# Create default destination path (parent directory since New-ADTTemplate creates a subdirectory)
$defaultParentDestination = "C:\PSADT\v$moduleVersion\$appVendor"
Write-StepOutput "Default deployment parent location: $defaultParentDestination" -Status "Info"

$customPath = Read-Host "Press Enter to use default location, or specify custom parent path"
if ([string]::IsNullOrWhiteSpace($customPath)) {
    $parentDestination = $defaultParentDestination
} else {
    $parentDestination = $customPath
}

# The actual deployment will be created in a subdirectory named after the app
$templateName = "$appName $appVersion"
$deploymentPath = Join-Path $parentDestination $templateName

Write-StepOutput "Deployment will be created at: $deploymentPath" -Status "Success"

Write-Separator

# Step 3: Create the PSADT deployment
Write-StepOutput "Creating PSADT deployment template..." -Status "Progress"

try {
    # Ensure the parent directory exists
    if (!(Test-Path $parentDestination)) {
        New-Item -Path $parentDestination -ItemType Directory -Force | Out-Null
    }
    
    # Create the deployment using New-ADTTemplate
    # Note: New-ADTTemplate creates a subdirectory with the name, so we pass the parent path
    $templateName = "$appName $appVersion"
    New-ADTTemplate -Destination $parentDestination -Name $templateName
    
    Write-StepOutput "Template creation command completed. Verifying folder structure..." -Status "Progress"
    
    # Wait for folder structure to be created
    $null = Wait-ForFolderStructure -DeploymentPath $deploymentPath
    
    Write-StepOutput "Successfully created deployment template: '$templateName'" -Status "Success"
    Write-StepOutput "Deployment created at: $deploymentPath" -Status "Info"
    
} catch {
    Write-StepOutput "Error creating deployment template: $($_.Exception.Message)" -Status "Error"
    Write-Host "Press any key to exit..." -ForegroundColor Red
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Separator

# Step 4: Download and replace asset files
Write-StepOutput "Checking asset file configuration..." -Status "Download"

$assetsPath = Join-Path $deploymentPath "Assets"

# Create asset mappings for non-empty URLs
$assetFiles = @{}
if (![string]::IsNullOrWhiteSpace($AppIconUrl)) {
    $assetFiles["AppIcon.png"] = $AppIconUrl
}
if (![string]::IsNullOrWhiteSpace($BannerClassicUrl)) {
    $assetFiles["Banner.Classic.png"] = $BannerClassicUrl
}

# Process asset downloads
Update-AssetFiles -AssetsPath $assetsPath -AssetMappings $assetFiles

# Step 5: Modify deployment script configuration
Write-StepOutput "Updating deployment script configuration..." -Status "Config"

$deploymentScriptPath = Join-Path $deploymentPath "Invoke-AppDeployToolkit.ps1"
Update-DeploymentScript -DeploymentScriptPath $deploymentScriptPath -AppVendor $appVendor -AppName $appName -AppVersion $appVersion -Author $Author -DefaultAppArch $DefaultAppArch

Write-Separator

# Step 6: Modify configuration file
Write-StepOutput "Updating configuration file..." -Status "Config"

$configPath = Join-Path $deploymentPath "Config\config.psd1"
Update-ConfigFile -ConfigPath $configPath -CompanyName $CompanyName -LogPath $LogPath

Write-Separator

# Step 7: Summary
$LogSummary = Get-SummaryValue $LogPath
$CompanySummary = Get-SummaryValue $CompanyName
$AssetSummary = if ([string]::IsNullOrWhiteSpace($AppIconUrl) -and [string]::IsNullOrWhiteSpace($BannerClassicUrl)) { 
    'Template defaults' 
} else { 
    'Custom assets downloaded' 
}

Write-StepOutput "Deployment creation completed!" -Status "Success"
Write-Host @"

📋 Deployment Summary:
   • Vendor: $appVendor
   • Application: $appName
   • Version: $appVersion  
   • Location: $deploymentPath
   • PSADT Module Version: v$moduleVersion
   • Company: $CompanySummary
   • Log Path: $LogSummary
   • Assets: $AssetSummary

🎯 Next Steps:
   1. Navigate to: $deploymentPath
   2. Customize the deployment script as needed
   3. Add your installation files to the Files directory
   4. Test the deployment

"@ -ForegroundColor Green

Write-StepOutput "Opening deployment folder and VS Code..." -Status "Info"

# Ask user if they want to open folder and VS Code
$openChoice = Read-Host "`nWould you like to open the deployment folder in Explorer and VS Code? (Y/N)"

if ($openChoice -match '^[Yy]') {
    try {
        # Open in Explorer
        Write-StepOutput "Opening deployment folder in Explorer..." -Status "Progress"
        Start-Process "explorer.exe" -ArgumentList $deploymentPath
        Write-StepOutput "Explorer opened successfully" -Status "Success"
        
        # Open in VS Code
        Write-StepOutput "Opening deployment folder in VS Code..." -Status "Progress"
        if (Get-Command "code" -ErrorAction SilentlyContinue) {
            # Change to the deployment directory and run 'code .' to open current directory
            Push-Location $deploymentPath
            Start-Process "code" -ArgumentList "." -WorkingDirectory $deploymentPath
            Pop-Location
            Write-StepOutput "VS Code opened successfully" -Status "Success"
        } else {
            Write-StepOutput "VS Code not found in PATH. Opening folder in Explorer only." -Status "Warning"
        }
        
    } catch {
        Write-StepOutput "Error opening applications: $($_.Exception.Message)" -Status "Error"
        Write-StepOutput "Please navigate manually to: $deploymentPath" -Status "Info"
    }
} else {
    Write-StepOutput "Deployment folder location: $deploymentPath" -Status "Info"
}

Write-Host "`n🎉 Script completed successfully! Happy deploying! 🚀" -ForegroundColor Green