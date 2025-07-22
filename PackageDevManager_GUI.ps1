# Package Development Manager GUI
# Version: 2.1.4.7.0.2.4.8.9.4 / 6.5.5.3.5

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Import the refactored functions (for standalone EXE, these are embedded below)
if (Test-Path "$PSScriptRoot\PckgDevMngr_refactored_clean.ps1") {
    . "$PSScriptRoot\PckgDevMngr_refactored_clean.ps1"
}

#region GUI Variables
$global:selectedPackageLocation = ""
$global:selectedVendor = ""
$global:selectedApplication = ""
$global:selectedVersion = ""
$global:selectedDeployment = ""
$global:selectedDRMBuild = ""
$global:exeArguments = ""

# Package file variables (like original code)
$global:sharedSoftwarePath = "\\nasv0718.uhc.com\packagingarchive"
$global:softwareSetup = ""
$global:pathFlag = $false
$global:Arguments = ""
$global:Exectue = ""
$global:Cred = ""
$global:filemsi = ""
$global:filemst = ""
$global:fileexe = ""
$global:fileps1 = ""
$global:pickFile = ""
$global:repPath = ""
$global:msiExist = $false
$global:mstExist = $false
$global:exeExist = $false
$global:ps1Exist = $false

# Package type detection variables
$global:packageType = ""
$global:installCommand = ""

# Package download variables
$global:localPackagePath = ""
$global:downloadCompleted = $false

# GUI Control References
$global:vendorDropdown = $null
$global:appDropdown = $null
$global:versionDropdown = $null
$global:deploymentDropdown = $null
$global:drmBuildDropdown = $null
$global:packageStatusLabel = $null

# Validation report tracking
$global:lastValidationReport = $null
#endregion

#region Modern Professional Form Creation
function New-PackageDevManagerForm {
    # Create main form with modern design
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Package Deployment Manager 2.0 "
    $form.Size = New-Object System.Drawing.Size(1200, 750)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.MinimumSize = New-Object System.Drawing.Size(1000, 650)
    $form.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $form.Icon = [System.Drawing.SystemIcons]::Application

    # Add modern header panel with Optum brand colors
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.Size = New-Object System.Drawing.Size(1200, 80)
    $headerPanel.Location = New-Object System.Drawing.Point(0, 0)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 245, 240)  # Light beige/cream like Optum image
    $headerPanel.Dock = "Top"
    $form.Controls.Add($headerPanel)

    # Add title label with dark blue text
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Package Deployment Manager 2.0"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 55, 109)  # Dark blue text
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $titleLabel.Size = New-Object System.Drawing.Size(450, 25)
    $titleLabel.TextAlign = "MiddleLeft"
    $headerPanel.Controls.Add($titleLabel)

    # Add Optum logo with signature orange color (much larger size)
    $optumLabel = New-Object System.Windows.Forms.Label
    $optumLabel.Text = "Optum"
    $optumLabel.Font = New-Object System.Drawing.Font("Segoe UI", 42, [System.Drawing.FontStyle]::Bold)
    $optumLabel.ForeColor = [System.Drawing.Color]::FromArgb(255, 102, 51)  # Optum signature orange
    $optumLabel.Location = New-Object System.Drawing.Point(950, 10)
    $optumLabel.Size = New-Object System.Drawing.Size(240, 60)
    $optumLabel.TextAlign = "MiddleCenter"
    $optumLabel.BackColor = [System.Drawing.Color]::Transparent
    $headerPanel.Controls.Add($optumLabel)

    # Add subtitle with dark blue text
    $subtitleLabel = New-Object System.Windows.Forms.Label
    $subtitleLabel.Text = "GPS Package Deployment and Validation Suite"
    $subtitleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 55, 109)  # Dark blue text
    $subtitleLabel.Location = New-Object System.Drawing.Point(20, 42)
    $subtitleLabel.Size = New-Object System.Drawing.Size(350, 15)
    $subtitleLabel.TextAlign = "MiddleLeft"
    $headerPanel.Controls.Add($subtitleLabel)

    # Add developer signature at bottom right corner
    $signatureLabel = New-Object System.Windows.Forms.Label
    $signatureLabel.Text = " Developed by Bharadwaj - GPS Packaging Team"
    $signatureLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $signatureLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 55, 109)  # Dark blue
    $signatureLabel.BackColor = [System.Drawing.Color]::FromArgb(248, 245, 240)  # Match header color
    $signatureLabel.Location = New-Object System.Drawing.Point(790, 666)  # Bottom right
    $signatureLabel.Size = New-Object System.Drawing.Size(340, 15)
    $signatureLabel.TextAlign = "MiddleRight"
    $signatureLabel.Anchor = "Bottom,Right"  # Stay in bottom right when resized
    $form.Controls.Add($signatureLabel)

    return $form
}

function Add-ModernPackageLocationPanel {
    param($form)

    # Create main container panel for package location selection with rich styling
    $locationPanel = New-Object System.Windows.Forms.Panel
    $locationPanel.Location = New-Object System.Drawing.Point(20, 100)
    $locationPanel.Size = New-Object System.Drawing.Size(1140, 100)
    $locationPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $locationPanel.BorderStyle = "FixedSingle"
    $form.Controls.Add($locationPanel)

    # Add section title with rich styling
    $sectionTitle = New-Object System.Windows.Forms.Label
    $sectionTitle.Text = "Package Location Selection"
    $sectionTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $sectionTitle.ForeColor = [System.Drawing.Color]::FromArgb(25, 135, 84)
    $sectionTitle.Location = New-Object System.Drawing.Point(15, 8)
    $sectionTitle.Size = New-Object System.Drawing.Size(400, 25)
    $sectionTitle.BackColor = [System.Drawing.Color]::Transparent
    $locationPanel.Controls.Add($sectionTitle)

    # Create toggle buttons for package location selection (more reliable than radio buttons)
    $global:selectedPackageLocation = "PackageArchive"

    # Package Archive Button
    $global:btnPackageArchive = New-Object System.Windows.Forms.Button
    $global:btnPackageArchive.Text = "Package Archive"
    $global:btnPackageArchive.Location = New-Object System.Drawing.Point(50, 40)
    $global:btnPackageArchive.Size = New-Object System.Drawing.Size(200, 40)
    $global:btnPackageArchive.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $global:btnPackageArchive.BackColor = [System.Drawing.Color]::FromArgb(25, 135, 84)
    $global:btnPackageArchive.ForeColor = [System.Drawing.Color]::White
    $global:btnPackageArchive.FlatStyle = "Flat"
    $locationPanel.Controls.Add($global:btnPackageArchive)

    # Completed Package Button
    $global:btnCompletedPackage = New-Object System.Windows.Forms.Button
    $global:btnCompletedPackage.Text = "Completed Package"
    $global:btnCompletedPackage.Location = New-Object System.Drawing.Point(350, 40)
    $global:btnCompletedPackage.Size = New-Object System.Drawing.Size(200, 40)
    $global:btnCompletedPackage.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $global:btnCompletedPackage.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $global:btnCompletedPackage.ForeColor = [System.Drawing.Color]::White
    $global:btnCompletedPackage.FlatStyle = "Flat"
    $locationPanel.Controls.Add($global:btnCompletedPackage)

    # Local Button
    $global:btnLocal = New-Object System.Windows.Forms.Button
    $global:btnLocal.Text = "Local"
    $global:btnLocal.Location = New-Object System.Drawing.Point(650, 40)
    $global:btnLocal.Size = New-Object System.Drawing.Size(200, 40)
    $global:btnLocal.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $global:btnLocal.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $global:btnLocal.ForeColor = [System.Drawing.Color]::White
    $global:btnLocal.FlatStyle = "Flat"
    $locationPanel.Controls.Add($global:btnLocal)

    # Add click event handlers for toggle buttons
    $global:btnPackageArchive.Add_Click({
        # Reset all button colors
        $global:btnPackageArchive.BackColor = [System.Drawing.Color]::FromArgb(25, 135, 84)
        $global:btnCompletedPackage.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        $global:btnLocal.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)

        $global:selectedPackageLocation = "PackageArchive"
        Write-Host "Package Archive selected" -ForegroundColor Green
        try {
            Invoke-VendorScan -archiveLocation $global:sharedSoftwarePath
        } catch {
            Write-Host "Error updating package directory: $($_.Exception.Message)" -ForegroundColor Red
        }
    })

    $global:btnCompletedPackage.Add_Click({
        # Reset all button colors
        $global:btnPackageArchive.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        $global:btnCompletedPackage.BackColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
        $global:btnLocal.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)

        $global:selectedPackageLocation = "CompletedPackage"
        Write-Host "Completed Package selected" -ForegroundColor Blue
        try {
            Invoke-VendorScan -archiveLocation "\\nas00036pn\Cert-Staging\2_Completed Packages"
        } catch {
            Write-Host "Error updating package directory: $($_.Exception.Message)" -ForegroundColor Red
        }
    })

    $global:btnLocal.Add_Click({
        # Reset all button colors
        $global:btnPackageArchive.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        $global:btnCompletedPackage.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        $global:btnLocal.BackColor = [System.Drawing.Color]::FromArgb(255, 193, 7)

        $global:selectedPackageLocation = "Local"
        Write-Host "Local Package selected" -ForegroundColor Yellow
        try {
            Load-LocalPackages
        } catch {
            Write-Host "Error loading local packages: $($_.Exception.Message)" -ForegroundColor Red
        }
    })



    # Return the toggle buttons for reference
    return @{
        PackageArchive = $global:btnPackageArchive
        CompletedPackage = $global:btnCompletedPackage
        Local = $global:btnLocal
    }
}

function Add-ModernDropdownControls {
    param($form)

    # Create main container panel for package selection with rich styling
    $selectionPanel = New-Object System.Windows.Forms.Panel
    $selectionPanel.Location = New-Object System.Drawing.Point(20, 220)
    $selectionPanel.Size = New-Object System.Drawing.Size(1140, 200)
    $selectionPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $selectionPanel.BorderStyle = "FixedSingle"
    $form.Controls.Add($selectionPanel)

    # Add section title with rich styling
    $sectionTitle = New-Object System.Windows.Forms.Label
    $sectionTitle.Text = "Package Selection & Configuration"
    $sectionTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $sectionTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
    $sectionTitle.Location = New-Object System.Drawing.Point(15, 8)
    $sectionTitle.Size = New-Object System.Drawing.Size(400, 25)
    $sectionTitle.BackColor = [System.Drawing.Color]::Transparent
    $selectionPanel.Controls.Add($sectionTitle)

    # Create modern dropdown cards
    $dropdownConfigs = @(
        @{ Label = "Vendor"; Name = "Vendor"; X = 20; Y = 50; Width = 220 },
        @{ Label = "Application"; Name = "Application"; X = 250; Y = 50; Width = 220 },
        @{ Label = "Version"; Name = "Version"; X = 480; Y = 50; Width = 180 },
        @{ Label = "Deployment"; Name = "Deployment"; X = 670; Y = 50; Width = 180 },
        @{ Label = "DRM Build"; Name = "DRMBuild"; X = 860; Y = 50; Width = 180 }
    )

    $dropdowns = @{}

    foreach ($config in $dropdownConfigs) {
        # Create rich card container
        $cardPanel = New-Object System.Windows.Forms.Panel
        $cardPanel.Location = New-Object System.Drawing.Point([int]$config.X, [int]$config.Y)
        $cardPanel.Size = New-Object System.Drawing.Size([int]$config.Width, 80)
        $cardPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
        $cardPanel.BorderStyle = "FixedSingle"
        $selectionPanel.Controls.Add($cardPanel)

        # Add enhanced label
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $config.Label
        $label.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $label.ForeColor = [System.Drawing.Color]::FromArgb(0, 123, 255)
        $label.Location = New-Object System.Drawing.Point(10, 8)
        $label.Size = New-Object System.Drawing.Size(([int]$config.Width - 20), 20)
        $label.TextAlign = "MiddleLeft"
        $cardPanel.Controls.Add($label)

        # Add enhanced dropdown with search capability for Vendor
        $dropdown = New-Object System.Windows.Forms.ComboBox
        $dropdown.Location = New-Object System.Drawing.Point(10, 35)
        $dropdown.Size = New-Object System.Drawing.Size(([int]$config.Width - 20), 25)

        # Enable search for Vendor dropdown, otherwise use DropDownList
        if ($config.Name -eq "Vendor") {
            $dropdown.DropDownStyle = "DropDown"
            $dropdown.AutoCompleteMode = "SuggestAppend"
            $dropdown.AutoCompleteSource = "ListItems"
        } else {
            $dropdown.DropDownStyle = "DropDownList"
        }

        $dropdown.Font = New-Object System.Drawing.Font("Segoe UI", 9)
        $dropdown.BackColor = [System.Drawing.Color]::FromArgb(255, 255, 255)
        $dropdown.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
        $cardPanel.Controls.Add($dropdown)

        $dropdowns[$config.Name] = $dropdown
    }

    # Store global reference for vendor dropdown
    $global:vendorDropdown = $dropdowns["Vendor"]

    # Add package type status display
    $statusPanel = New-Object System.Windows.Forms.Panel
    $statusPanel.Location = New-Object System.Drawing.Point(20, 140)
    $statusPanel.Size = New-Object System.Drawing.Size(1100, 40)
    $statusPanel.BackColor = [System.Drawing.Color]::FromArgb(220, 248, 198)
    $statusPanel.BorderStyle = "FixedSingle"
    $selectionPanel.Controls.Add($statusPanel)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Package Type: Not Detected - Please select a package to analyze"
    $statusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $statusLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 135, 84)
    $statusLabel.Location = New-Object System.Drawing.Point(15, 10)
    $statusLabel.Size = New-Object System.Drawing.Size(1070, 20)
    $statusLabel.TextAlign = "MiddleLeft"
    $statusPanel.Controls.Add($statusLabel)

    # Store global reference for updates
    $global:packageStatusLabel = $statusLabel

    # Store global references for other dropdowns
    $global:appDropdown = $dropdowns["Application"]
    $global:versionDropdown = $dropdowns["Version"]
    $global:deploymentDropdown = $dropdowns["Deployment"]
    $global:drmBuildDropdown = $dropdowns["DRMBuild"]

    return $dropdowns
}

function Add-ModernExeArgumentsControl {
    param($form)

    # Create arguments panel
    $argsPanel = New-Object System.Windows.Forms.Panel
    $argsPanel.Location = New-Object System.Drawing.Point(20, 440)
    $argsPanel.Size = New-Object System.Drawing.Size(1140, 120)
    $argsPanel.BackColor = [System.Drawing.Color]::White
    $argsPanel.BorderStyle = "FixedSingle"
    $form.Controls.Add($argsPanel)

    # Add section title
    $sectionTitle = New-Object System.Windows.Forms.Label
    $sectionTitle.Text = "EXE Arguments Configuration"
    $sectionTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $sectionTitle.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $sectionTitle.Location = New-Object System.Drawing.Point(15, 10)
    $sectionTitle.Size = New-Object System.Drawing.Size(400, 25)
    $sectionTitle.BackColor = [System.Drawing.Color]::Transparent
    $argsPanel.Controls.Add($sectionTitle)

    # Add toggle checkbox for custom arguments
    $customArgsCheckbox = New-Object System.Windows.Forms.CheckBox
    $customArgsCheckbox.Text = "Enable Custom Arguments (Default commands will be populated based on package type)"
    $customArgsCheckbox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $customArgsCheckbox.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $customArgsCheckbox.Location = New-Object System.Drawing.Point(15, 35)
    $customArgsCheckbox.Size = New-Object System.Drawing.Size(600, 20)
    $argsPanel.Controls.Add($customArgsCheckbox)

    # Set checked state after adding to panel
    try {
        $customArgsCheckbox.Checked = $false
    } catch {
        Write-Host "Warning: Could not set checkbox state" -ForegroundColor Yellow
    }

    # Add text box with modern styling (initially disabled)
    $argsTextBox = New-Object System.Windows.Forms.TextBox
    $argsTextBox.Name = "argsTextBox"  # Add name for easier identification
    $argsTextBox.Location = New-Object System.Drawing.Point(15, 65)
    $argsTextBox.Size = New-Object System.Drawing.Size(1110, 45)
    $argsTextBox.Multiline = $true
    $argsTextBox.ScrollBars = "Vertical"
    $argsTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $argsTextBox.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $argsTextBox.BorderStyle = "FixedSingle"
    $argsTextBox.ForeColor = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $argsTextBox.Enabled = $false
    $argsTextBox.Text = "Custom arguments will be populated here based on package type selection..."
    $argsPanel.Controls.Add($argsTextBox)

    # Store references globally for easier access
    $global:argsTextBox = $argsTextBox
    $global:customArgsCheckbox = $customArgsCheckbox

    return @($argsTextBox, $customArgsCheckbox)
}

function Get-DefaultCommand {
    <#
    .SYNOPSIS
        Returns default command line arguments based on detected package type
    #>

    if (-not $global:packageStatusLabel) {
        return "/S /v/qn"  # Default silent install
    }

    $packageType = $global:packageStatusLabel.Text

    switch -Wildcard ($packageType) {
        "*MSI Only*" {
            return "msiexec /i `"[MSI_FILE]`" /qn /l*v `"[LOG_FILE]`""
        }
        "*MSI + MST*" {
            return "msiexec /i `"[MSI_FILE]`" TRANSFORMS=`"[MST_FILE]`" /qn /l*v `"[LOG_FILE]`""
        }
        "*EXE*" {
            return "/S /v/qn"
        }
        "*PSADTK*" {
            return "Application.exe Application.ps1 -DeployMode Silent -DeploymentType Install"
        }
        default {
            return "/S /v/qn"  # Default silent install
        }
    }
}

function Get-ArchiveLocation {
    <#
    .SYNOPSIS
        Returns the archive location path based on selection type
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$SelectionType
    )

    switch ($SelectionType) {
        "PackageArchive" {
            return $global:sharedSoftwarePath
        }
        "CompletedPackage" {
            return "\\nasv0718.uhc.com\packagingarchive\Completed"
        }
        "Local" {
            return "C:\temp\LocalPackages"
        }
        default {
            return $global:sharedSoftwarePath
        }
    }
}

#endregion

#region Modern Action Button Controls
function Add-ModernActionButtons {
    param($form)

    # Create action buttons panel with rich styling
    $buttonsPanel = New-Object System.Windows.Forms.Panel
    $buttonsPanel.Location = New-Object System.Drawing.Point(20, 580)
    $buttonsPanel.Size = New-Object System.Drawing.Size(1140, 120)
    $buttonsPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
    $buttonsPanel.BorderStyle = "FixedSingle"
    $form.Controls.Add($buttonsPanel)

    # Add section title with rich styling
    $sectionTitle = New-Object System.Windows.Forms.Label
    $sectionTitle.Text = "Action Center"
    $sectionTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $sectionTitle.ForeColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $sectionTitle.Location = New-Object System.Drawing.Point(15, 8)
    $sectionTitle.Size = New-Object System.Drawing.Size(200, 25)
    $sectionTitle.BackColor = [System.Drawing.Color]::Transparent
    $buttonsPanel.Controls.Add($sectionTitle)



    # Define modern button configurations - arranged in 2 rows
    $buttonConfigs = @(
        @{ Text = "Install"; Name = "Install"; X = 20; Y = 40; Width = 110; Height = 35; Color = [System.Drawing.Color]::FromArgb(40, 167, 69); TextColor = "White" },
        @{ Text = "Validate Install"; Name = "ValidateInstall"; X = 140; Y = 40; Width = 130; Height = 35; Color = [System.Drawing.Color]::FromArgb(40, 167, 69); TextColor = "White" },
        @{ Text = "Uninstall"; Name = "Uninstall"; X = 280; Y = 40; Width = 110; Height = 35; Color = [System.Drawing.Color]::FromArgb(220, 53, 69); TextColor = "White" },
        @{ Text = "Validate Uninstall"; Name = "ValidateUninstall"; X = 400; Y = 40; Width = 130; Height = 35; Color = [System.Drawing.Color]::FromArgb(220, 53, 69); TextColor = "White" },
        @{ Text = "Clear Media"; Name = "ClearMedia"; X = 540; Y = 40; Width = 110; Height = 35; Color = [System.Drawing.Color]::FromArgb(0, 123, 255); TextColor = "White" },
        @{ Text = "Refresh"; Name = "Refresh"; X = 660; Y = 40; Width = 110; Height = 35; Color = [System.Drawing.Color]::FromArgb(0, 123, 255); TextColor = "White" },
        @{ Text = "Copy Logs"; Name = "CopyLogs"; X = 780; Y = 40; Width = 110; Height = 35; Color = [System.Drawing.Color]::FromArgb(111, 66, 193); TextColor = "White" },
        @{ Text = "Generate TSV"; Name = "GenerateTSV"; X = 900; Y = 40; Width = 110; Height = 35; Color = [System.Drawing.Color]::FromArgb(111, 66, 193); TextColor = "White" },
        @{ Text = "QC Tool"; Name = "QCTool"; X = 1020; Y = 40; Width = 110; Height = 35; Color = [System.Drawing.Color]::FromArgb(111, 66, 193); TextColor = "White" }
    )

    $buttons = @{}

    foreach ($config in $buttonConfigs) {
        # Create enhanced button with modern styling
        $button = New-Object System.Windows.Forms.Button
        $button.Text = $config.Text
        $button.Location = New-Object System.Drawing.Point($config.X, $config.Y)
        $button.Size = New-Object System.Drawing.Size($config.Width, $config.Height)
        $button.BackColor = $config.Color
        $button.ForeColor = if ($config.TextColor -eq "White") { [System.Drawing.Color]::White } else { [System.Drawing.Color]::Black }
        $button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $button.FlatStyle = "Standard"
        $button.Cursor = "Hand"

        $buttonsPanel.Controls.Add($button)
        $buttons[$config.Name] = $button
    }

    # Add event handlers
    $buttons["Install"].Add_Click({ Invoke-InstallPackage })

    $buttons["Uninstall"].Add_Click({ Invoke-UninstallPackage })

    $buttons["ValidateInstall"].Add_Click({ Invoke-ValidateInstall })
    $buttons["ValidateUninstall"].Add_Click({ Invoke-ValidateUninstall })
    $buttons["ClearMedia"].Add_Click({ Clear-Media })
    $buttons["Refresh"].Add_Click({ Refresh-Form })
    $buttons["CopyLogs"].Add_Click({ Copy-Logs })
    $buttons["GenerateTSV"].Add_Click({ Generate-TSV })
    $buttons["QCTool"].Add_Click({ Run-QCTool })

    # Add status information panel at bottom
    $statusInfoPanel = New-Object System.Windows.Forms.Panel
    $statusInfoPanel.Location = New-Object System.Drawing.Point(140, 80)
    $statusInfoPanel.Size = New-Object System.Drawing.Size(990, 25)
    $statusInfoPanel.BackColor = [System.Drawing.Color]::FromArgb(248, 245, 240)  # Match header color
    $statusInfoPanel.BorderStyle = "FixedSingle"
    $buttonsPanel.Controls.Add($statusInfoPanel)

    $statusInfoLabel = New-Object System.Windows.Forms.Label
    $statusInfoLabel.Text = "💡 Ready for package deployment and validation operations"
    $statusInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $statusInfoLabel.ForeColor = [System.Drawing.Color]::FromArgb(25, 55, 109)  # Dark blue to match header
    $statusInfoLabel.Location = New-Object System.Drawing.Point(15, 3)
    $statusInfoLabel.Size = New-Object System.Drawing.Size(970, 20)
    $statusInfoLabel.TextAlign = "MiddleLeft"
    $statusInfoPanel.Controls.Add($statusInfoLabel)

    return $buttons
}

#endregion

#region Helper Functions
function Get-SelectedPackagePath {
    <#
    .SYNOPSIS
        Auto-detects the package path based on current selections
    #>

    try {
        # Check if local package is selected
        if ($global:selectedPackageLocation -eq "Local" -and $global:pickFile) {
            return $global:pickFile
        }

        # For archive/completed packages, try to find based on vendor/app selection
        if ($global:selectedVendor -and $global:selectedApplication) {
            $basePath = switch ($global:selectedPackageLocation) {
                "PackageArchive" { $global:sharedSoftwarePath }
                "CompletedPackage" { "$global:sharedSoftwarePath\Completed" }
                default { $global:sharedSoftwarePath }
            }

            # Look for common package files
            $searchPath = "$basePath\$global:selectedVendor\$global:selectedApplication"
            if (Test-Path $searchPath) {
                # Find MSI or EXE files
                $msiFiles = Get-ChildItem -Path $searchPath -Filter "*.msi" -Recurse | Select-Object -First 1
                $exeFiles = Get-ChildItem -Path $searchPath -Filter "*.exe" -Recurse | Select-Object -First 1

                if ($msiFiles) { return $msiFiles.FullName }
                if ($exeFiles) { return $exeFiles.FullName }
            }
        }

        return $null
    }
    catch {
        Write-Warning "Error detecting package path: $($_.Exception.Message)"
        return $null
    }
}

function Invoke-VendorScan {
    <#
    .SYNOPSIS
        Quickly scans for vendor folder names (like original code)
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$archiveLocation
    )

    try {
        if (Test-Path $archiveLocation) {
            Write-Host "Getting vendor folders..." -ForegroundColor Gray

            # Quick scan - just get directory names (no deep scanning)
            $vendorFolders = Get-ChildItem $archiveLocation -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

            if ($vendorFolders.Count -gt 0) {
                Write-Host "Found $($vendorFolders.Count) vendors: $($vendorFolders -join ', ')" -ForegroundColor Green

                # Update vendor dropdown with actual folder names
                Update-VendorDropdown -vendors $vendorFolders
            }
            else {
                Write-Host "No vendor folders found in $archiveLocation" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Archive location not accessible: $archiveLocation" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error scanning vendors: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Initialize-QuickWorkflow {
    <#
    .SYNOPSIS
        Sets up the quick workflow environment
    #>

    # Initialize global variables for quick access
    $global:sharedSoftwarePath = "\\nasv0718.uhc.com\packagingarchive"
    $global:toolOutputDir = "$env:SystemDrive\temp\AutoTest\Tool_Output"

    # Ensure output directory exists
    if (-not (Test-Path $global:toolOutputDir)) {
        New-Item -ItemType Directory -Path $global:toolOutputDir -Force | Out-Null
    }
}
#endregion

#region Event Handlers
function Load-PackageArchive {
    Write-Host "Loading Package Archive..." -ForegroundColor Yellow
    $global:selectedPackageLocation = "PackageArchive"

    # Quick scan - just get vendor folder names (like original)
    $archiveLocation = "\\nasv0718.uhc.com\packagingarchive"
    Write-Host "Scanning vendor folders from: $archiveLocation" -ForegroundColor Green

    # Fast vendor scan - only get directory names
    Invoke-VendorScan -archiveLocation $archiveLocation
}

function Load-CompleteLocation {
    Write-Host "Loading Completed Packages..." -ForegroundColor Yellow
    $global:selectedPackageLocation = "CompletedPackage"

    # Quick scan - just get vendor folder names (like Package Archive)
    $completedLocation = "\\nas00036pn\Cert-Staging\2_Completed Packages"
    Write-Host "Scanning vendor folders from: $completedLocation" -ForegroundColor Green
    Invoke-VendorScan -archiveLocation $completedLocation
}

function Load-LocalPackages {
    Write-Host "=== LOCAL PACKAGE SELECTION ===" -ForegroundColor Cyan
    $global:selectedPackageLocation = "Local"

    # Open file dialog for local package selection
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "Installer Files (*.msi;*.exe)|*.msi;*.exe|MSI Files (*.msi)|*.msi|EXE Files (*.exe)|*.exe|All Files (*.*)|*.*"
    $openFileDialog.Title = "Select Local Installer File"
    $openFileDialog.InitialDirectory = "C:\"

    if ($openFileDialog.ShowDialog() -eq "OK") {
        $selectedFile = $openFileDialog.FileName
        $fileName = [System.IO.Path]::GetFileName($selectedFile)
        $fileExtension = [System.IO.Path]::GetExtension($selectedFile).ToLower()

        Write-Host "Selected installer: $selectedFile" -ForegroundColor Green

        # Set global variables for local package
        $global:pickFile = $selectedFile
        $global:localPackagePath = [System.IO.Path]::GetDirectoryName($selectedFile)

        # Determine package type and set appropriate global variables
        switch ($fileExtension) {
            ".msi" {
                $global:filemsi = $selectedFile
                $global:msiExist = $true
                $global:mstExist = $false
                $global:exeExist = $false
                $global:ps1Exist = $false
                $global:packageType = "MSI Only"
                Write-Host "Detected: MSI Package" -ForegroundColor Magenta
            }
            ".exe" {
                $global:fileexe = $selectedFile
                $global:exeExist = $true
                $global:msiExist = $false
                $global:mstExist = $false
                $global:ps1Exist = $false
                $global:packageType = "EXE File"
                Write-Host "Detected: EXE Package" -ForegroundColor Magenta
            }
            default {
                Write-Host "Unsupported file type: $fileExtension" -ForegroundColor Red
                return
            }
        }

        # Update EXE Arguments Configuration to show selected installer
        Update-ExeArgumentsForLocal -installerPath $selectedFile

        # Update package status
        if ($global:packageStatusLabel) {
            $global:packageStatusLabel.Text = "Status: Local installer selected - $fileName ($global:packageType)"
            $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Green
        }

        # Enable install functionality
        Enable-LocalInstallMode

        Write-Host "Local installer ready for installation with custom parameters" -ForegroundColor Green
        Write-Host "You can now:" -ForegroundColor Yellow
        Write-Host "  1. Add custom parameters in EXE Arguments Configuration" -ForegroundColor Gray
        Write-Host "  2. Click Install to proceed with installation and scanning" -ForegroundColor Gray
    }
    else {
        Write-Host "No installer selected" -ForegroundColor Yellow
    }
}

function Update-ExeArgumentsForLocal {
    <#
    .SYNOPSIS
        Updates the EXE Arguments Configuration to show the selected local installer
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$installerPath
    )

    try {
        $fileName = [System.IO.Path]::GetFileName($installerPath)

        # Use global references to update the controls
        if ($global:argsTextBox -and $global:customArgsCheckbox) {
            # Enable custom arguments
            $global:customArgsCheckbox.Checked = $true
            $global:argsTextBox.Enabled = $true
            $global:argsTextBox.BackColor = [System.Drawing.Color]::White

            # Set the installer path and default parameters
            $defaultArgs = "`"$installerPath`""
            if ($global:packageType -eq "EXE File") {
                $defaultArgs += " /S /silent"
            } elseif ($global:packageType -eq "MSI Only") {
                $defaultArgs = "msiexec.exe /i `"$installerPath`" /qn /norestart"
            }

            $global:argsTextBox.Text = $defaultArgs
            $global:exeArguments = $defaultArgs

            Write-Host "EXE Arguments Configuration updated with: $defaultArgs" -ForegroundColor Green
        } else {
            Write-Host "Warning: EXE Arguments controls not found" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error updating EXE Arguments Configuration: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Enable-LocalInstallMode {
    <#
    .SYNOPSIS
        Enables install functionality for local packages
    #>

    try {
        # Set flags to indicate local package is ready
        $global:downloadCompleted = $true
        $global:packageAnalysisCompleted = $true

        # Build install command based on package type
        if ($global:packageType -eq "MSI Only") {
            $global:installCommand = Build-InstallCommand -Operation "Install"
        } elseif ($global:packageType -eq "EXE File") {
            if ($global:exeArguments) {
                $global:installCommand = $global:exeArguments
            } else {
                $global:installCommand = "`"$global:fileexe`" /S /silent"
            }
        }

        Write-Host "Install command prepared: $global:installCommand" -ForegroundColor Green
        Write-Host "Local install mode enabled - ready for installation" -ForegroundColor Green
    }
    catch {
        Write-Host "Error enabling local install mode: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-InstallPackage {
    Write-Host "=== INSTALL PACKAGE ===" -ForegroundColor Cyan

    try {
        # Update status bar
        if ($global:packageStatusLabel) {
            $global:packageStatusLabel.Text = "Status: Installing package..."
            $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Orange
            $global:packageStatusLabel.Refresh()
        }

        # Quick base scan like original
        Write-Host "Performing base scan..." -ForegroundColor Yellow
        Invoke-BaseScan
        Write-Host "Base scan completed!" -ForegroundColor Green

        # Check if we have package type detected and install command ready
        if ($global:packageType -and $global:installCommand -and $global:packageType -ne "Unknown") {
            Write-Host "=== INTELLIGENT PACKAGE INSTALLATION ===" -ForegroundColor Cyan
            Write-Host "Package Type: $global:packageType" -ForegroundColor Yellow
            Write-Host "Install Command: $global:installCommand" -ForegroundColor Yellow

            # Step 1: Copy package to local temp (ALWAYS copy to C:\temp first)
            Write-Host "Step 1: Copying package to local temp..." -ForegroundColor Cyan
            $downloadSuccess = $false

            if ($global:selectedPackageLocation -eq "Local" -and $global:pickFile) {
                # Handle local file - copy to C:\temp structure
                Write-Host "Copying local file to C:\temp..." -ForegroundColor Yellow
                $downloadSuccess = Copy-LocalPackageToTemp -localFilePath $global:pickFile
            } else {
                # Handle network packages - download to C:\temp
                Write-Host "Downloading network package to C:\temp..." -ForegroundColor Yellow
                $downloadSuccess = Invoke-PackageDownload
            }

            if (-not $downloadSuccess) {
                # Offer fallback option to work with network files directly
                $result = [System.Windows.Forms.MessageBox]::Show("Package download failed!`n`nWould you like to try installing directly from the network location?`n`nNote: This may require elevated network permissions.", "Download Error - Try Network Install?", "YesNo", "Question")

                if ($result -eq "Yes") {
                    Write-Host "Attempting network-based installation..." -ForegroundColor Yellow
                    # Set up for network-based installation
                    $global:downloadCompleted = $false

                    # Build network path for direct installation
                    $basePath = switch ($global:selectedPackageLocation) {
                        "PackageArchive" { $global:sharedSoftwarePath }
                        "CompletedPackage" { "\\nas00036pn\Cert-Staging\2_Completed Packages" }
                        default { $global:sharedSoftwarePath }
                    }
                    $global:localPackagePath = "$basePath\$global:selectedVendor\$global:selectedApplication\$global:selectedVersion\$global:selectedDeployment\$global:selectedDRMBuild"
                    Write-Host "Using network path: $global:localPackagePath" -ForegroundColor Yellow
                }
                else {
                    [System.Windows.Forms.MessageBox]::Show("Installation cancelled. Please check network connectivity and try again.", "Installation Cancelled", "OK", "Information")
                    return
                }
            }

            # Step 2: Execute via Task Scheduler with high permissions
            Write-Host "Step 2: Executing via Task Scheduler with elevated permissions..." -ForegroundColor Cyan
            $installSuccess = Invoke-TaskSchedulerInstall

            if ($installSuccess) {
                Write-Host "Package installation completed successfully!" -ForegroundColor Green

                # Update status bar
                if ($global:packageStatusLabel) {
                    $global:packageStatusLabel.Text = "Status: Installation completed successfully"
                    $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Green
                    $global:packageStatusLabel.Refresh()
                }

                # Automatically run post-installation validation
                Write-Host "Running automatic post-installation validation..." -ForegroundColor Cyan

                try {
                    $validationResult = Invoke-AutomaticValidation -ValidationType "Install"
                    Write-Host "Validation completed. Result: $validationResult" -ForegroundColor Green

                    if ($validationResult) {
                        [System.Windows.Forms.MessageBox]::Show("Package installation and validation completed successfully!`n`nPackage: $global:selectedApplication $global:selectedVersion`nType: $global:packageType`nLocation: $global:localPackagePath`n`nValidation report generated automatically.`nClick 'Validate Install' to view the report.", "Installation & Validation Complete", "OK", "Information")
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show("Package installation completed but validation failed!`n`nPackage: $global:selectedApplication $global:selectedVersion`nPlease check the logs for details.", "Installation Complete - Validation Error", "OK", "Warning")
                    }
                }
                catch {
                    Write-Host "Error during validation: $($_.Exception.Message)" -ForegroundColor Red
                    [System.Windows.Forms.MessageBox]::Show("Package installation completed but validation encountered an error!`n`nError: $($_.Exception.Message)`n`nPackage: $global:selectedApplication $global:selectedVersion", "Installation Complete - Validation Error", "OK", "Warning")
                }
            }
            else {
                # Update status bar for failure
                if ($global:packageStatusLabel) {
                    $global:packageStatusLabel.Text = "Status: Installation failed"
                    $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Red
                    $global:packageStatusLabel.Refresh()
                }
                [System.Windows.Forms.MessageBox]::Show("Package installation failed!`n`nPlease check the installation log at C:\temp\PackageDevManager_InstallLog.txt", "Installation Error", "OK", "Error")
            }

        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Base scan completed.`n`nPlease install the package manually, then click 'Validate Install'.", "Manual Installation Required", "OK", "Information")
        }
    }
    catch {
        Write-Error "Installation error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Installation error: $($_.Exception.Message)", "Installation Error", "OK", "Error")
    }
}

function Invoke-UninstallPackage {
    Write-Host "=== UNINSTALL PACKAGE ===" -ForegroundColor Cyan

    try {
        # Update status bar
        if ($global:packageStatusLabel) {
            $global:packageStatusLabel.Text = "Status: Uninstalling package..."
            $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Orange
            $global:packageStatusLabel.Refresh()
        }

        # Quick base scan like original
        Write-Host "Performing base scan..." -ForegroundColor Yellow
        Invoke-BaseScan
        Write-Host "Base scan completed!" -ForegroundColor Green

        # Check if we have package type detected and can build uninstall command
        if ($global:packageType -and $global:packageType -ne "Unknown") {
            Write-Host "=== INTELLIGENT PACKAGE UNINSTALLATION ===" -ForegroundColor Cyan
            Write-Host "Package Type: $global:packageType" -ForegroundColor Yellow

            # Step 1: Copy package to local temp (ALWAYS copy to C:\temp first)
            Write-Host "Step 1: Copying package to local temp for uninstall..." -ForegroundColor Cyan
            $downloadSuccess = $false

            if ($global:selectedPackageLocation -eq "Local" -and $global:pickFile) {
                # Handle local file - copy to C:\temp structure
                Write-Host "Copying local file to C:\temp..." -ForegroundColor Yellow
                $downloadSuccess = Copy-LocalPackageToTemp -localFilePath $global:pickFile
            } else {
                # Handle network packages - download to C:\temp
                Write-Host "Downloading network package to C:\temp..." -ForegroundColor Yellow
                $downloadSuccess = Invoke-PackageDownload
            }

            if (-not $downloadSuccess) {
                [System.Windows.Forms.MessageBox]::Show("Package copy/download failed for uninstall!`n`nPackage files needed for proper uninstallation.", "Copy/Download Error", "OK", "Error")
                return
            }

            # Step 2: Build uninstall command based on package type
            $uninstallCommand = Build-UninstallCommand

            if ($uninstallCommand) {
                Write-Host "Uninstall Command: $uninstallCommand" -ForegroundColor Yellow

                # Step 3: Execute uninstall via Task Scheduler
                Write-Host "Executing uninstall via Task Scheduler..." -ForegroundColor Cyan
                $uninstallSuccess = Invoke-TaskSchedulerUninstall -uninstallCommand $uninstallCommand

                if ($uninstallSuccess) {
                    Write-Host "Package uninstallation completed!" -ForegroundColor Green

                    # Update status bar
                    if ($global:packageStatusLabel) {
                        $global:packageStatusLabel.Text = "Status: Uninstallation completed successfully"
                        $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Green
                        $global:packageStatusLabel.Refresh()
                    }

                    # Automatically run post-uninstallation validation
                    Write-Host "Running automatic post-uninstallation validation..." -ForegroundColor Cyan
                    $validationResult = Invoke-AutomaticValidation -ValidationType "Uninstall"

                    if ($validationResult) {
                        [System.Windows.Forms.MessageBox]::Show("Package uninstallation and validation completed successfully!`n`nPackage: $global:selectedApplication $global:selectedVersion`nType: $global:packageType`n`nValidation report generated automatically.`nClick 'Validate Uninstall' to view the report.", "Uninstall & Validation Complete", "OK", "Information")
                    }
                    else {
                        [System.Windows.Forms.MessageBox]::Show("Package uninstallation completed but validation failed!`n`nPackage: $global:selectedApplication $global:selectedVersion`nPlease check the logs for details.", "Uninstall Complete - Validation Error", "OK", "Warning")
                    }
                }
                else {
                    # Update status bar for failure
                    if ($global:packageStatusLabel) {
                        $global:packageStatusLabel.Text = "Status: Uninstallation failed"
                        $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Red
                        $global:packageStatusLabel.Refresh()
                    }
                    [System.Windows.Forms.MessageBox]::Show("Package uninstallation failed!`n`nPlease check the installation log at C:\temp\PackageDevManager_InstallLog.txt", "Uninstall Error", "OK", "Error")
                }
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Cannot build uninstall command for this package type.`n`nPlease uninstall manually, then click 'Validate Uninstall'.", "Manual Uninstall Required", "OK", "Information")
            }
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("Base scan completed.`n`nPackage type not detected. Please uninstall the package manually, then click 'Validate Uninstall'.", "Manual Uninstall Required", "OK", "Information")
        }
    }
    catch {
        Write-Error "Uninstallation error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Uninstallation error: $($_.Exception.Message)", "Uninstall Error", "OK", "Error")
    }
}

function Invoke-AutomaticValidation {
    <#
    .SYNOPSIS
        Runs automatic validation after installation/uninstallation
    .PARAMETER ValidationType
        Type of validation: Install or Uninstall
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Install", "Uninstall")]
        [string]$ValidationType
    )

    try {
        Write-Host "=== AUTOMATIC VALIDATION ($ValidationType) ===" -ForegroundColor Cyan

        # Perform second scan and comparison
        Write-Host "Performing post-$($ValidationType.ToLower()) scan..." -ForegroundColor Yellow
        Invoke-SecondScan -OperationType $ValidationType

        # Run validation tests
        Write-Host "Running validation tests..." -ForegroundColor Yellow
        Test-Shortcuts
        Test-ProgramFiles
        Test-UninstallKeys

        # Generate comprehensive report
        Write-Host "Generating validation report..." -ForegroundColor Yellow
        $reportPath = Generate-CreativeHTMLReport -OperationType $ValidationType

        if ($reportPath) {
            Write-Host "Validation report generated: $reportPath" -ForegroundColor Green
            $global:lastValidationReport = $reportPath
            return $true
        }
        else {
            Write-Host "Failed to generate validation report" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "Automatic validation error: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Invoke-ValidateInstall {
    <#
    .SYNOPSIS
        Opens the installation validation report (generated automatically)
    #>

    try {
        if ($global:lastValidationReport -and (Test-Path $global:lastValidationReport)) {
            Write-Host "Opening installation validation report..." -ForegroundColor Green
            Start-Process $global:lastValidationReport
            [System.Windows.Forms.MessageBox]::Show("Installation validation report opened!`n`nReport: $global:lastValidationReport", "Report Opened", "OK", "Information")
        }
        else {
            Write-Host "No validation report found. Running validation..." -ForegroundColor Yellow
            $validationResult = Invoke-AutomaticValidation -ValidationType "Install"

            if ($validationResult -and $global:lastValidationReport) {
                Start-Process $global:lastValidationReport
                [System.Windows.Forms.MessageBox]::Show("Installation validation completed and report opened!`n`nReport: $global:lastValidationReport", "Validation Complete", "OK", "Information")
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Validation failed or no report generated.`n`nPlease check the logs for details.", "Validation Error", "OK", "Error")
            }
        }
    }
    catch {
        Write-Error "Validation error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Validation failed: $($_.Exception.Message)", "Validation Error", "OK", "Error")
    }
}

function Invoke-ValidateUninstall {
    <#
    .SYNOPSIS
        Opens the uninstallation validation report (generated automatically)
    #>

    try {
        if ($global:lastValidationReport -and (Test-Path $global:lastValidationReport)) {
            Write-Host "Opening uninstallation validation report..." -ForegroundColor Green
            Start-Process $global:lastValidationReport
            [System.Windows.Forms.MessageBox]::Show("Uninstallation validation report opened!`n`nReport: $global:lastValidationReport", "Report Opened", "OK", "Information")
        }
        else {
            Write-Host "No validation report found. Running validation..." -ForegroundColor Yellow
            $validationResult = Invoke-AutomaticValidation -ValidationType "Uninstall"

            if ($validationResult -and $global:lastValidationReport) {
                Start-Process $global:lastValidationReport
                [System.Windows.Forms.MessageBox]::Show("Uninstallation validation completed and report opened!`n`nReport: $global:lastValidationReport", "Validation Complete", "OK", "Information")
            }
            else {
                [System.Windows.Forms.MessageBox]::Show("Validation failed or no report generated.`n`nPlease check the logs for details.", "Validation Error", "OK", "Error")
            }
        }
    }
    catch {
        Write-Error "Validation error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Validation failed: $($_.Exception.Message)", "Validation Error", "OK", "Error")
    }
}

function Open-UHGInstallLog {
    Write-Host "Opening UHG Install Log..."

    $logPath = "$env:windir\UHGLogs"
    if (Test-Path $logPath) {
        Start-Process "explorer.exe" -ArgumentList $logPath
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("UHG Logs directory not found: $logPath", "Log Directory Not Found", "OK", "Warning")
    }
}

function Clear-Media {
    Write-Host "Clearing media and temporary files..." -ForegroundColor Cyan
    try {
        # Clear global variables
        $global:selectedVendor = ""
        $global:selectedApplication = ""
        $global:selectedVersion = ""
        $global:selectedDeployment = ""
        $global:selectedDRMBuild = ""
        $global:pickFile = ""
        $global:packageType = ""
        $global:installCommand = ""
        $global:exeArguments = ""
        $global:localPackagePath = ""
        $global:downloadCompleted = $false

        # Remove copied package media folders from C:\temp\<Application>\<Version>\...
        Write-Host "Removing copied package media folders..." -ForegroundColor Yellow
        $tempPath = "C:\temp"
        $foldersRemoved = 0
        $totalSizeRemoved = 0

        if (Test-Path $tempPath) {
            # Get all directories in C:\temp that are application folders
            $applicationFolders = Get-ChildItem $tempPath -Directory -ErrorAction SilentlyContinue

            foreach ($appFolder in $applicationFolders) {
                try {
                    # Check if this folder has the package structure: C:\temp\<App>\<Version>\...
                    $versionFolders = Get-ChildItem $appFolder.FullName -Directory -ErrorAction SilentlyContinue

                    foreach ($versionFolder in $versionFolders) {
                        try {
                            # Check if version folder contains DRMBUILD or package files
                            $drmBuildFolders = Get-ChildItem $versionFolder.FullName -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "DRMBUILD" }
                            $hasPackageFiles = (Get-ChildItem $versionFolder.FullName -File -Recurse -ErrorAction SilentlyContinue | Where-Object {
                                $_.Extension -in @('.msi', '.exe', '.ps1', '.mst')
                            }).Count -gt 0

                            if ($drmBuildFolders.Count -gt 0 -or $hasPackageFiles) {
                                # Calculate folder size before removal
                                $folderSize = (Get-ChildItem $versionFolder.FullName -Recurse -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
                                $folderSizeMB = [math]::Round($folderSize / 1MB, 2)

                                Write-Host "Removing package media: $($appFolder.Name)\$($versionFolder.Name) ($folderSizeMB MB)" -ForegroundColor Yellow
                                Remove-Item $versionFolder.FullName -Recurse -Force -ErrorAction SilentlyContinue
                                $foldersRemoved++
                                $totalSizeRemoved += $folderSizeMB
                            }
                        } catch {
                            Write-Host "Could not remove version folder $($appFolder.Name)\$($versionFolder.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                        }
                    }

                    # Remove empty application folder if no version folders remain
                    $remainingVersions = Get-ChildItem $appFolder.FullName -Directory -ErrorAction SilentlyContinue
                    if ($remainingVersions.Count -eq 0) {
                        Write-Host "Removing empty application folder: $($appFolder.Name)" -ForegroundColor Yellow
                        Remove-Item $appFolder.FullName -Force -ErrorAction SilentlyContinue
                    }
                } catch {
                    Write-Host "Could not process application folder $($appFolder.Name): $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }

        # Also clean up specific known temp locations
        $knownTempPaths = @(
            "C:\temp\AutoTest",
            "C:\temp\AutoTestScan",
            "C:\temp\LocalPackages"
        )

        foreach ($path in $knownTempPaths) {
            if (Test-Path $path) {
                try {
                    Write-Host "Cleaning up: $path" -ForegroundColor Yellow
                    Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
                } catch {
                    Write-Host "Could not remove $path`: $($_.Exception.Message)" -ForegroundColor Yellow
                }
            }
        }

        Write-Host "Media clearing completed!" -ForegroundColor Green
        Write-Host "- Form fields and selections cleared" -ForegroundColor Green
        Write-Host "- $foldersRemoved package media folders removed from C:\temp" -ForegroundColor Green
        Write-Host "- Total space freed: $([math]::Round($totalSizeRemoved, 2)) MB" -ForegroundColor Green
        Write-Host "- Temporary directories cleaned up" -ForegroundColor Green

        $sizeText = if ($totalSizeRemoved -gt 0) { "`n• Total space freed: $([math]::Round($totalSizeRemoved, 2)) MB" } else { "" }
        [System.Windows.Forms.MessageBox]::Show("Media cleared successfully!`n`n• Form fields and selections cleared`n• $foldersRemoved package media folders removed from C:\temp$sizeText`n• Temporary directories cleaned up", "Clear Media Complete", "OK", "Information")
    } catch {
        Write-Host "Error clearing media: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error clearing media: $($_.Exception.Message)", "Clear Media Error", "OK", "Error")
    }
}

function Refresh-Form {
    Write-Host "Refreshing form..." -ForegroundColor Cyan
    try {
        # Clear current selections
        Write-Host "Clearing current selections..." -ForegroundColor Yellow
        $global:selectedVendor = ""
        $global:selectedApplication = ""
        $global:selectedVersion = ""
        $global:selectedDeployment = ""
        $global:selectedDRMBuild = ""
        $global:packageType = ""
        $global:installCommand = ""

        # Clear all dropdown selections and items
        Write-Host "Clearing dropdown controls..." -ForegroundColor Yellow
        if ($global:vendorDropdown) {
            $global:vendorDropdown.SelectedIndex = -1
            $global:vendorDropdown.Items.Clear()
        }
        if ($global:appDropdown) {
            $global:appDropdown.SelectedIndex = -1
            $global:appDropdown.Items.Clear()
        }
        if ($global:versionDropdown) {
            $global:versionDropdown.SelectedIndex = -1
            $global:versionDropdown.Items.Clear()
        }
        if ($global:deploymentDropdown) {
            $global:deploymentDropdown.SelectedIndex = -1
            $global:deploymentDropdown.Items.Clear()
        }
        if ($global:drmBuildDropdown) {
            $global:drmBuildDropdown.SelectedIndex = -1
            $global:drmBuildDropdown.Items.Clear()
        }

        # Clear package status label
        if ($global:packageStatusLabel) {
            $global:packageStatusLabel.Text = "Status: Ready for package selection"
            $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
        }

        # Refresh based on current package location selection
        switch ($global:selectedPackageLocation) {
            "PackageArchive" {
                Write-Host "Refreshing Package Archive..." -ForegroundColor Green
                Invoke-VendorScan -archiveLocation $global:sharedSoftwarePath
            }
            "CompletedPackage" {
                Write-Host "Refreshing Completed Packages..." -ForegroundColor Blue
                Invoke-VendorScan -archiveLocation "\\nasv0718.uhc.com\packagingarchive\Completed"
            }
            "Local" {
                Write-Host "Refreshing Local Packages..." -ForegroundColor Yellow
                Load-LocalPackages
            }
            default {
                Write-Host "Refreshing Package Archive (default)..." -ForegroundColor Green
                Invoke-VendorScan -archiveLocation $global:sharedSoftwarePath
            }
        }

        Write-Host "Form refresh completed successfully!" -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show("Form refreshed successfully. All selections cleared and vendor list updated.", "Refresh Complete", "OK", "Information")
    } catch {
        Write-Host "Error refreshing form: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error refreshing form: $($_.Exception.Message)", "Refresh Error", "OK", "Error")
    }
}



function Copy-Logs {
    <#
    .SYNOPSIS
        Copies logs to the Install_Uninstall_Reports folder
    #>

    Write-Host "=== COPYING LOGS ===" -ForegroundColor Cyan

    try {
        # Set up destination path as requested
        $AppLogsPath = "C:\Temp\AutoTestScan\AppLogs"

        Write-Host "Destination path: $AppLogsPath" -ForegroundColor Yellow

        # Clean up existing AppLogs directory to avoid copying old logs
        if (Test-Path -Path $AppLogsPath) {
            Write-Host "Cleaning up existing AppLogs directory..." -ForegroundColor Yellow
            Remove-Item -Path $AppLogsPath -Recurse -Force -ErrorAction SilentlyContinue
        }

        # Ensure the destination directory exists
        if (-not (Test-Path -Path $AppLogsPath)) {
            New-Item -ItemType Directory -Path $AppLogsPath -Force | Out-Null
            Write-Host "Created destination directory: $AppLogsPath" -ForegroundColor Green
        }

        # Copy logs based on difflogFiles.txt
        $tempPath = "$env:SystemDrive\temp\AutoTestScan"
        $diffLogFilesPath = "$tempPath\difflogFiles.txt"

        $logsCopied = 0

        if (Test-Path -Path $diffLogFilesPath) {
            Write-Host "Reading log files list from: $diffLogFilesPath" -ForegroundColor Yellow
            $filesToCopy = Get-Content -Path $diffLogFilesPath -ErrorAction SilentlyContinue

            foreach ($file in $filesToCopy) {
                if ($file -and (Test-Path -Path $file)) {
                    try {
                        Copy-Item -Path $file -Destination $AppLogsPath -Force
                        Write-Host "Copied: $file" -ForegroundColor Green
                        $logsCopied++
                    }
                    catch {
                        Write-Host "Failed to copy: $file - $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "File does not exist: $file" -ForegroundColor Yellow
                }
            }
        }

        # Copy NEW UHG logs based on diffUHGLogs.txt (only new items from installation)
        $diffUHGLogsPath = "$tempPath\diffUHGLogs.txt"
        if (Test-Path $diffUHGLogsPath) {
            Write-Host "Copying NEW UHG logs based on installation differences..." -ForegroundColor Yellow

            try {
                $newUHGLogs = Get-Content $diffUHGLogsPath -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
                if ($newUHGLogs) {
                    foreach ($log in $newUHGLogs) {
                        # Skip Health Check files (ignore as requested - performance optimization)
                        if ($log.Name -like "Health Check*") {
                            Write-Host "Skipped Health Check file: $($log.Name) (ignored - performance optimization)" -ForegroundColor DarkYellow
                            continue
                        }

                        # Skip BrowserUpdateManager folder (ignore as requested - performance optimization)
                        if (($log.PSIsContainer -or $log.ItemType -eq "Folder") -and $log.Name -eq "BrowserUpdateManager") {
                            Write-Host "Skipped BrowserUpdateManager folder: $($log.Name) (ignored - performance optimization)" -ForegroundColor DarkYellow
                            continue
                        }

                        try {
                            if ($log.PSIsContainer -or $log.ItemType -eq "Folder") {
                                # It's a directory - copy recursively
                                $destinationFolder = Join-Path $AppLogsPath $log.Name
                                if (Test-Path $log.FullName) {
                                    Copy-Item -Path $log.FullName -Destination $destinationFolder -Recurse -Force
                                    Write-Host "✓ Copied NEW UHG folder: $($log.Name)" -ForegroundColor Green
                                    $logsCopied++
                                }
                            }
                            else {
                                # It's a file - copy directly
                                if (Test-Path $log.FullName) {
                                    Copy-Item -Path $log.FullName -Destination $AppLogsPath -Force
                                    Write-Host "✓ Copied NEW UHG file: $($log.Name)" -ForegroundColor Green
                                    $logsCopied++
                                }
                            }
                        }
                        catch {
                            Write-Host "✗ Failed to copy NEW UHG item: $($log.Name) - $($_.Exception.Message)" -ForegroundColor Red
                        }
                    }
                }
                else {
                    Write-Host "No new UHG logs found to copy" -ForegroundColor Yellow
                }
            }
            catch {
                Write-Host "Error reading diffUHGLogs.txt: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        else {
            Write-Host "No new UHG logs found (diffUHGLogs.txt not found)" -ForegroundColor Yellow
        }

        if ($logsCopied -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("Successfully copied $logsCopied log files to:`n$AppLogsPath", "Copy Logs Complete", "OK", "Information")
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("No log files found to copy.`n`nDestination: $AppLogsPath", "Copy Logs - No Files", "OK", "Warning")
        }
    }
    catch {
        Write-Host "Error copying logs: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error copying logs: $($_.Exception.Message)", "Copy Logs Error", "OK", "Error")
    }
}

function Generate-TSV {
    <#
    .SYNOPSIS
        Generates a TSV file from the registry uninstall keys detected during installation
    #>

    Write-Host "=== GENERATING TSV FILE ===" -ForegroundColor Cyan

    try {
        # Set up paths
        $tempPath = "$env:SystemDrive\temp\AutoTestScan"
        $toolOutputDir = "$env:SystemDrive\temp\AutoTest\Tool_Output"
        $diffUninstallKeyPath = "$tempPath\diffUninstallKey.txt"

        # Ensure output directory exists
        if (-not (Test-Path $toolOutputDir)) {
            New-Item -ItemType Directory -Path $toolOutputDir -Force | Out-Null
            Write-Host "Created output directory: $toolOutputDir" -ForegroundColor Yellow
        }

        # Check if diff file exists
        if (-not (Test-Path $diffUninstallKeyPath)) {
            Write-Host "No registry changes detected. Please run an installation first." -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show("No registry changes detected.`n`nPlease run an installation first to generate registry data.", "TSV Generation - No Data", "OK", "Warning")
            return
        }

        # Read and parse the uninstall keys JSON
        Write-Host "Reading registry changes from: $diffUninstallKeyPath" -ForegroundColor Yellow
        $uninstallKeysJson = Get-Content $diffUninstallKeyPath -Raw -ErrorAction SilentlyContinue

        if (-not $uninstallKeysJson) {
            Write-Host "Registry changes file is empty." -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show("Registry changes file is empty.`n`nNo data available for TSV generation.", "TSV Generation - No Data", "OK", "Warning")
            return
        }

        $uninstallKeys = $uninstallKeysJson | ConvertFrom-Json
        if (-not $uninstallKeys -or $uninstallKeys.Count -eq 0) {
            Write-Host "No registry entries found in changes file." -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show("No registry entries found.`n`nNo data available for TSV generation.", "TSV Generation - No Data", "OK", "Warning")
            return
        }

        Write-Host "Found $($uninstallKeys.Count) registry entries to process" -ForegroundColor Green

        # Get system information
        Write-Host "Gathering system information..." -ForegroundColor Yellow
        $os = Get-WmiObject Win32_OperatingSystem
        $osName = $os.Caption
        $osVersion = $os.Version
        $processorArchitecture = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" -Name PROCESSOR_ARCHITECTURE -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "PROCESSOR_ARCHITECTURE"
        $processorIdentifier = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" -Name PROCESSOR_IDENTIFIER -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "PROCESSOR_IDENTIFIER"

        # Create TSV header
        $data = @(
           "DisplayName`tDisplayVersion`tPublisher`tInstallDate`tHidden`tOS`tVersion`tProcessor_Architecture`tProcessor_Identifier`tWow6432Node`tComments`tGUID`tDRMBUILD"
        )

        $processedCount = 0
        $tsvFileName = ""

        # Process each uninstall key
        foreach ($registryEntry in $uninstallKeys) {
            try {
                $registryPath = $registryEntry.PSPath

                if (-not $registryPath) {
                    Write-Host "Skipping entry with missing PSPath" -ForegroundColor Yellow
                    continue
                }

                Write-Host "Processing: $registryPath" -ForegroundColor Gray

                # Extract values from the registry entry object (already loaded from JSON)
                $displayName = $registryEntry.DisplayName
                $displayVersion = $registryEntry.DisplayVersion
                $publisher = $registryEntry.Publisher
                $installDate = $registryEntry.InstallDate

                # Try to get additional properties from live registry if available
                $systemComponent = ""
                $comments = ""
                $drmValue = ""

                if (Test-Path "Registry::$registryPath") {
                    try {
                        $systemComponentValue = Get-ItemProperty -Path "Registry::$registryPath" -Name SystemComponent -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "SystemComponent"
                        if ($systemComponentValue -eq 1) {
                            $systemComponent = "Yes"
                        }

                        $commentsValue = Get-ItemProperty -Path "Registry::$registryPath" -Name Comments -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "Comments"
                        if ($commentsValue -and $commentsValue.Substring(0,3) -ieq "DRM") {
                            $comments = $commentsValue
                        }

                        $drmValue = Get-ItemProperty -Path "Registry::$registryPath" -Name DRMBUILD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty "DRMBUILD"
                    }
                    catch {
                        Write-Host "Warning: Could not read additional properties from $registryPath" -ForegroundColor Yellow
                    }
                }

                # Clean display name if needed
                if ($displayName) {
                    $firstChar = [char]$displayName[0]
                    $asciiValue = [int]$firstChar

                    if (($asciiValue -lt 48) -or ($asciiValue -gt 126) -or (($asciiValue -gt 57) -and ($asciiValue -lt 64))) {
                        $displayName = "***Contains non displayable characters***"
                    }
                }

                # Extract GUID from registry path
                $guid = ($registryPath -split "\\")[-1]

                # Determine if it's WOW6432Node
                $isWow64 = if ($registryPath -like "*WOW6432Node*") { "Yes" } else { "No" }

                # Add data row
                $data += "$displayName`t$displayVersion`t$publisher`t$installDate`t$systemComponent`t$osName`t$osVersion`t$processorArchitecture`t$processorIdentifier`t$isWow64`t$comments`t$guid`t$drmValue"

                # Use first entry for filename
                if ($processedCount -eq 0 -and $displayName -and $displayVersion) {
                    $tsvFileName = "$displayName" + "_" + "$displayVersion"
                    # Clean filename of invalid characters
                    $tsvFileName = $tsvFileName -replace '[\\/:*?"<>|]', '_'
                }

                $processedCount++
            }
            catch {
                Write-Host "Error processing registry entry: $($_.Exception.Message)" -ForegroundColor Red
            }
        }

        if ($processedCount -eq 0) {
            Write-Host "No valid registry entries were processed." -ForegroundColor Yellow
            [System.Windows.Forms.MessageBox]::Show("No valid registry entries were processed.`n`nTSV generation failed.", "TSV Generation Error", "OK", "Warning")
            return
        }

        # Generate filename
        if (-not $tsvFileName) {
            $tsvFileName = "PackageInstall_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        }

        $tsvFilePath = "$toolOutputDir\$tsvFileName.tsv"

        # Write TSV file
        Write-Host "Writing TSV file: $tsvFilePath" -ForegroundColor Yellow
        $data | Out-File -FilePath $tsvFilePath -Encoding UTF8

        # Copy to package folder if available
        if ($global:localPackagePath -and (Test-Path $global:localPackagePath)) {
            try {
                Copy-Item -Path $tsvFilePath -Destination $global:localPackagePath -Force
                Write-Host "TSV file copied to package folder: $global:localPackagePath" -ForegroundColor Green
            }
            catch {
                Write-Host "Warning: Could not copy TSV to package folder: $($_.Exception.Message)" -ForegroundColor Yellow
            }
        }

        Write-Host "TSV file generated successfully!" -ForegroundColor Green
        Write-Host "Processed $processedCount registry entries" -ForegroundColor Green
        Write-Host "File location: $tsvFilePath" -ForegroundColor Green

        [System.Windows.Forms.MessageBox]::Show("TSV file generated successfully!`n`nProcessed: $processedCount registry entries`nLocation: $tsvFilePath", "TSV Generation Complete", "OK", "Information")
    }
    catch {
        Write-Host "Error generating TSV file: $($_.Exception.Message)" -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error generating TSV file:`n`n$($_.Exception.Message)", "TSV Generation Error", "OK", "Error")
    }
}

function Run-QCTool {
    <#
    .SYNOPSIS
        This will run the QC tool when the QC Button clicked.
    #>

    Write-Host "Launching QC Tool..." -ForegroundColor Cyan
    Start-Process -FilePath "$Dirpath\QC_Tool.EXE"
}

#endregion

#region Main Execution
function Show-PackageDevManagerGUI {
    <#
    .SYNOPSIS
        Main function to display the Package Development Manager GUI
    #>

    try {
        # Initialize working directories
        Initialize-WorkingDirectories

        # Create the main form with modern design
        $form = New-PackageDevManagerForm

        # Add modern package location toggle buttons
        $locationButtons = Add-ModernPackageLocationPanel -form $form

        # Add modern dropdown controls
        $dropdowns = Add-ModernDropdownControls -form $form

        # Add modern exe arguments control
        $argsControls = Add-ModernExeArgumentsControl -form $form
        $argsTextBox = $argsControls[0]
        $customArgsCheckbox = $argsControls[1]

        # Add event handler for checkbox after initialization
        try {
            $customArgsCheckbox.Add_CheckedChanged({
                try {
                    if ($this.Checked) {
                        $argsTextBox.Enabled = $true
                        $argsTextBox.BackColor = [System.Drawing.Color]::White
                        # Populate default command based on current package type
                        $defaultCommand = Get-DefaultCommand
                        $argsTextBox.Text = $defaultCommand
                        $global:exeArguments = $argsTextBox.Text
                        Write-Host "Custom arguments enabled: $global:exeArguments" -ForegroundColor Green
                    } else {
                        $argsTextBox.Enabled = $false
                        $argsTextBox.BackColor = [System.Drawing.Color]::FromArgb(248, 249, 250)
                        $argsTextBox.Text = "Custom arguments will be populated here based on package type selection..."
                        $global:exeArguments = ""
                        Write-Host "Custom arguments disabled" -ForegroundColor Yellow
                    }
                } catch {
                    Write-Host "Error in checkbox event: $($_.Exception.Message)" -ForegroundColor Red
                }
            })
        } catch {
            Write-Host "Warning: Could not add checkbox event handler" -ForegroundColor Yellow
        }

        # Add modern action buttons
        $buttons = Add-ModernActionButtons -form $form

        # Set up event handlers for dropdowns
        $dropdowns.Vendor.Add_SelectedIndexChanged({
            $global:selectedVendor = $dropdowns.Vendor.SelectedItem
            Write-Host "Selected Vendor: $($global:selectedVendor)" -ForegroundColor Cyan

            # Clear dependent dropdowns
            $dropdowns.Application.Items.Clear()
            $dropdowns.Version.Items.Clear()
            $dropdowns.Deployment.Items.Clear()
            $dropdowns.DRMBuild.Items.Clear()

            # Update application dropdown based on vendor selection
            Update-ApplicationDropdown -vendorCombo $dropdowns.Vendor -appCombo $dropdowns.Application
        })

        $dropdowns.Application.Add_SelectedIndexChanged({
            $global:selectedApplication = $dropdowns.Application.SelectedItem
            Write-Host "Selected Application: $($global:selectedApplication)" -ForegroundColor Cyan

            # Clear dependent dropdowns
            $dropdowns.Version.Items.Clear()
            $dropdowns.Deployment.Items.Clear()
            $dropdowns.DRMBuild.Items.Clear()

            # Update version dropdown based on application selection
            Update-VersionDropdown -appCombo $dropdowns.Application -versionCombo $dropdowns.Version
        })

        $dropdowns.Version.Add_SelectedIndexChanged({
            $global:selectedVersion = $dropdowns.Version.SelectedItem
            Write-Host "Selected Version: $($global:selectedVersion)" -ForegroundColor Cyan

            # Clear dependent dropdowns
            $dropdowns.Deployment.Items.Clear()
            $dropdowns.DRMBuild.Items.Clear()

            # Update deployment dropdown when version is selected
            Update-DeploymentDropdown -versionCombo $dropdowns.Version -deploymentCombo $dropdowns.Deployment
        })

        $dropdowns.Deployment.Add_SelectedIndexChanged({
            $global:selectedDeployment = $dropdowns.Deployment.SelectedItem
            Write-Host "Selected Deployment: $($global:selectedDeployment)" -ForegroundColor Cyan

            # Clear dependent dropdown
            $dropdowns.DRMBuild.Items.Clear()

            # Update DRM build dropdown when deployment is selected
            Update-DRMBuildDropdown -deploymentCombo $dropdowns.Deployment -drmBuildCombo $dropdowns.DRMBuild
        })

        $dropdowns.DRMBuild.Add_SelectedIndexChanged({
            $global:selectedDRMBuild = $dropdowns.DRMBuild.SelectedItem
            Write-Host "Selected DRMBuild: $($global:selectedDRMBuild)" -ForegroundColor Cyan

            # Scan for package files when DRM build is selected (final level)
            if ($global:selectedVendor -and $global:selectedApplication -and $global:selectedVersion -and $global:selectedDeployment -and $global:selectedDRMBuild) {
                $basePath = switch ($global:selectedPackageLocation) {
                    "PackageArchive" { $global:sharedSoftwarePath }
                    "CompletedPackage" { "\\nas00036pn\Cert-Staging\2_Completed Packages" }
                    default { $global:sharedSoftwarePath }
                }

                $packagePath = "$basePath\$global:selectedVendor\$global:selectedApplication\$global:selectedVersion\$global:selectedDeployment\$global:selectedDRMBuild"
                Write-Host "Final package path: $packagePath" -ForegroundColor Yellow
                Invoke-PackageFileScan -packagePath $packagePath
            }
        })

        # Set up event handler for exe arguments
        $argsTextBox.Add_TextChanged({
            # Only update global arguments if custom arguments are enabled
            if ($customArgsCheckbox.Checked) {
                $global:exeArguments = $argsTextBox.Text
                Write-Host "Custom arguments updated: $global:exeArguments" -ForegroundColor Cyan
            } else {
                $global:exeArguments = ""
            }
        })

        # Initialize quick workflow
        Initialize-QuickWorkflow

        # Auto-load Package Archive like original code
        Write-Host "Auto-loading Package Archive..." -ForegroundColor Cyan
        # Package Archive is already selected by default
        try {
            Invoke-VendorScan -archiveLocation $global:sharedSoftwarePath
        } catch {
            Write-Host "Error auto-loading package archive: $($_.Exception.Message)" -ForegroundColor Red
        }

        # Show the form
        Write-Host "Launching Package Development Manager GUI..." -ForegroundColor Green
        $form.Add_Shown({
            $form.Activate()
            Write-Host "GUI Ready! Quick workflow enabled." -ForegroundColor Green
        })
        [System.Windows.Forms.Application]::Run($form)
    }
    catch {
        Write-Error "Error launching GUI: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Error launching GUI: $($_.Exception.Message)", "GUI Error", "OK", "Error")
    }
}

function Update-VendorDropdown {
    <#
    .SYNOPSIS
        Updates vendor dropdown with actual folder names from network scan
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$vendors
    )

    if ($global:vendorDropdown) {
        $global:vendorDropdown.Items.Clear()
        $global:vendorDropdown.Items.AddRange($vendors)
        Write-Host "Vendor dropdown updated with $($vendors.Count) vendors" -ForegroundColor Green
    }
}

function Update-ApplicationDropdown {
    <#
    .SYNOPSIS
        Dynamically scans for applications under selected vendor
    #>
    param($vendorCombo, $appCombo)

    $appCombo.Items.Clear()

    if ($vendorCombo.SelectedItem -and $global:selectedPackageLocation) {
        $vendorName = $vendorCombo.SelectedItem

        # Build path based on current location
        $basePath = switch ($global:selectedPackageLocation) {
            "PackageArchive" { $global:sharedSoftwarePath }
            "CompletedPackage" { "\\nas00036pn\Cert-Staging\2_Completed Packages" }
            default { $global:sharedSoftwarePath }
        }

        $vendorPath = "$basePath\$vendorName"

        if (Test-Path $vendorPath) {
            Write-Host "Scanning applications for vendor: $vendorName" -ForegroundColor Gray
            $appFolders = Get-ChildItem $vendorPath -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

            if ($appFolders.Count -gt 0) {
                $appCombo.Items.AddRange($appFolders)
                Write-Host "Found $($appFolders.Count) applications for $vendorName" -ForegroundColor Green
            }
            else {
                $appCombo.Items.Add("No applications found")
            }
        }
        else {
            $appCombo.Items.Add("Vendor path not accessible")
        }
    }
}

function Update-VersionDropdown {
    <#
    .SYNOPSIS
        Dynamically scans for versions under selected vendor/application
    #>
    param($appCombo, $versionCombo)

    $versionCombo.Items.Clear()

    if ($appCombo.SelectedItem -and $global:selectedVendor -and $global:selectedPackageLocation) {
        $vendorName = $global:selectedVendor
        $appName = $appCombo.SelectedItem

        # Build path based on current location
        $basePath = switch ($global:selectedPackageLocation) {
            "PackageArchive" { $global:sharedSoftwarePath }
            "CompletedPackage" { "\\nas00036pn\Cert-Staging\2_Completed Packages" }
            default { $global:sharedSoftwarePath }
        }

        $appPath = "$basePath\$vendorName\$appName"

        if (Test-Path $appPath) {
            Write-Host "Scanning versions for: $vendorName\$appName" -ForegroundColor Gray
            $versionFolders = Get-ChildItem $appPath -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

            if ($versionFolders.Count -gt 0) {
                $versionCombo.Items.AddRange($versionFolders)
                Write-Host "Found $($versionFolders.Count) versions for $appName" -ForegroundColor Green
            }
            else {
                $versionCombo.Items.Add("No versions found")
            }
        }
        else {
            $versionCombo.Items.Add("Application path not accessible")
        }
    }
}

function Update-DeploymentDropdown {
    <#
    .SYNOPSIS
        Dynamically scans for deployment folders under selected vendor/application/version
    #>
    param($versionCombo, $deploymentCombo)

    $deploymentCombo.Items.Clear()

    if ($versionCombo.SelectedItem -and $global:selectedVendor -and $global:selectedApplication -and $global:selectedPackageLocation) {
        $vendorName = $global:selectedVendor
        $appName = $global:selectedApplication
        $versionName = $versionCombo.SelectedItem

        # Build path based on current location
        $basePath = switch ($global:selectedPackageLocation) {
            "PackageArchive" { $global:sharedSoftwarePath }
            "CompletedPackage" { "\\nas00036pn\Cert-Staging\2_Completed Packages" }
            default { $global:sharedSoftwarePath }
        }

        $versionPath = "$basePath\$vendorName\$appName\$versionName"

        if (Test-Path $versionPath) {
            Write-Host "Scanning deployments for: $vendorName\$appName\$versionName" -ForegroundColor Gray
            $deploymentFolders = Get-ChildItem $versionPath -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

            if ($deploymentFolders.Count -gt 0) {
                $deploymentCombo.Items.AddRange($deploymentFolders)
                Write-Host "Found $($deploymentFolders.Count) deployment types for $versionName" -ForegroundColor Green
            }
            else {
                $deploymentCombo.Items.Add("No deployments found")
            }
        }
        else {
            $deploymentCombo.Items.Add("Version path not accessible")
        }
    }
}

function Update-DRMBuildDropdown {
    <#
    .SYNOPSIS
        Dynamically scans for DRM Build folders under selected vendor/application/version/deployment
    #>
    param($deploymentCombo, $drmBuildCombo)

    $drmBuildCombo.Items.Clear()

    if ($deploymentCombo.SelectedItem -and $global:selectedVendor -and $global:selectedApplication -and $global:selectedVersion -and $global:selectedPackageLocation) {
        $vendorName = $global:selectedVendor
        $appName = $global:selectedApplication
        $versionName = $global:selectedVersion
        $deploymentName = $deploymentCombo.SelectedItem

        # Build path based on current location
        $basePath = switch ($global:selectedPackageLocation) {
            "PackageArchive" { $global:sharedSoftwarePath }
            "CompletedPackage" { "\\nas00036pn\Cert-Staging\2_Completed Packages" }
            default { $global:sharedSoftwarePath }
        }

        $deploymentPath = "$basePath\$vendorName\$appName\$versionName\$deploymentName"

        if (Test-Path $deploymentPath) {
            Write-Host "Scanning DRM builds for: $vendorName\$appName\$versionName\$deploymentName" -ForegroundColor Gray
            $drmBuildFolders = Get-ChildItem $deploymentPath -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

            if ($drmBuildFolders.Count -gt 0) {
                $drmBuildCombo.Items.AddRange($drmBuildFolders)
                Write-Host "Found $($drmBuildFolders.Count) DRM builds for $deploymentName" -ForegroundColor Green
            }
            else {
                $drmBuildCombo.Items.Add("No DRM builds found")
            }
        }
        else {
            $drmBuildCombo.Items.Add("Deployment path not accessible")
        }
    }
}

function Invoke-PackageFileScan {
    <#
    .SYNOPSIS
        Scans for actual package files (MSI, EXE, PS1) in the selected path and determines package type
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$packagePath
    )

    try {
        if (Test-Path $packagePath) {
            Write-Host "Scanning package files in: $packagePath" -ForegroundColor Gray

            # Reset global package variables
            $global:filemsi = ""
            $global:fileexe = ""
            $global:fileps1 = ""
            $global:filemst = ""
            $global:msiExist = $false
            $global:exeExist = $false
            $global:ps1Exist = $false
            $global:mstExist = $false
            $global:packageType = ""
            $global:installCommand = ""

            # Scan for package files
            $msiFiles = Get-ChildItem $packagePath -Filter "*.msi" -ErrorAction SilentlyContinue
            $exeFiles = Get-ChildItem $packagePath -Filter "*.exe" -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "ServiceUI.exe" }
            $ps1Files = Get-ChildItem $packagePath -Filter "*.ps1" -ErrorAction SilentlyContinue
            $mstFiles = Get-ChildItem $packagePath -Filter "*.mst" -ErrorAction SilentlyContinue

            # Update global variables
            if ($msiFiles.Count -gt 0) {
                $global:filemsi = $msiFiles[0].FullName
                $global:msiExist = $true
                Write-Host "Found MSI: $($msiFiles[0].Name)" -ForegroundColor Green
            }

            if ($exeFiles.Count -gt 0) {
                $global:fileexe = $exeFiles[0].FullName
                $global:exeExist = $true
                Write-Host "Found EXE: $($exeFiles[0].Name)" -ForegroundColor Green
            }

            if ($ps1Files.Count -gt 0) {
                $global:fileps1 = $ps1Files[0].FullName
                $global:ps1Exist = $true
                Write-Host "Found PS1: $($ps1Files[0].Name)" -ForegroundColor Green
            }

            if ($mstFiles.Count -gt 0) {
                $global:filemst = $mstFiles[0].FullName
                $global:mstExist = $true
                Write-Host "Found MST: $($mstFiles[0].Name)" -ForegroundColor Green
            }

            # Determine package type and build install command
            Determine-PackageType

            $totalFiles = $msiFiles.Count + $exeFiles.Count + $ps1Files.Count + $mstFiles.Count
            Write-Host "Package scan complete - found $totalFiles package files" -ForegroundColor Green
            Write-Host "Package Type: $global:packageType" -ForegroundColor Yellow
            Write-Host "Install Command: $global:installCommand" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error scanning package files: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Determine-PackageType {
    <#
    .SYNOPSIS
        Determines the package type and builds appropriate installation command
    #>

    try {
        # Check for PSADTK Wrapper (PS1 file present)
        if ($global:ps1Exist) {
            $global:packageType = "PSADTK Wrapper"

            # Build PSADTK install command using the new Build-InstallCommand function
            $global:installCommand = Build-InstallCommand -Operation "Install"

            Write-Host "Detected: PSADTK Wrapper - PowerShell App Deployment Toolkit" -ForegroundColor Magenta
            Write-Host "  PS1: $([System.IO.Path]::GetFileName($global:fileps1))" -ForegroundColor Gray
            if ($global:exeExist) {
                Write-Host "  EXE: $([System.IO.Path]::GetFileName($global:fileexe))" -ForegroundColor Gray
            }
            Update-PackageStatusLabel
            return
        }

        # Check for MSI + MST combination
        if ($global:msiExist -and $global:mstExist) {
            $global:packageType = "MSI + MST"
            $msiName = [System.IO.Path]::GetFileName($global:filemsi)
            $mstName = [System.IO.Path]::GetFileName($global:filemst)

            # Build MSI command with MST transform using new function
            $global:installCommand = Build-InstallCommand -Operation "Install"

            Write-Host "Detected: MSI with MST Transform" -ForegroundColor Magenta
            Write-Host "  MSI: $msiName" -ForegroundColor Gray
            Write-Host "  MST: $mstName" -ForegroundColor Gray
            Update-PackageStatusLabel
            return
        }

        # Check for MSI only
        if ($global:msiExist -and -not $global:mstExist) {
            $global:packageType = "MSI Only"
            $msiName = [System.IO.Path]::GetFileName($global:filemsi)

            # Build MSI command with logging using Build-InstallCommand function
            $global:installCommand = Build-InstallCommand -Operation "Install"

            Write-Host "Detected: MSI Package Only" -ForegroundColor Magenta
            Write-Host "  MSI: $msiName" -ForegroundColor Gray
            Update-PackageStatusLabel
            return
        }

        # Check for EXE file (excluding ServiceUI.exe)
        if ($global:exeExist) {
            $global:packageType = "EXE File"
            $exeName = [System.IO.Path]::GetFileName($global:fileexe)

            # Build EXE command with custom arguments
            if ($global:exeArguments) {
                $global:installCommand = "`"$global:fileexe`" $global:exeArguments"
            } else {
                # Common silent switches for EXE files
                $global:installCommand = "`"$global:fileexe`" /S /silent"
            }

            Write-Host "Detected: EXE Package" -ForegroundColor Magenta
            Write-Host "  EXE: $exeName" -ForegroundColor Gray
            Update-PackageStatusLabel
            return
        }

        # No recognized package type
        $global:packageType = "Unknown"
        $global:installCommand = ""
        Write-Host "Warning: No recognized package type found" -ForegroundColor Yellow

    }
    catch {
        Write-Host "Error determining package type: $($_.Exception.Message)" -ForegroundColor Red
        $global:packageType = "Error"
        $global:installCommand = ""
    }
}

function Update-PackageStatusLabel {
    <#
    .SYNOPSIS
        Updates the GUI status label with detected package type
    #>

    if ($global:packageStatusLabel) {
        $statusText = "Package Type: $global:packageType"

        # Set color based on package type
        $color = switch ($global:packageType) {
            "PSADTK Wrapper" { [System.Drawing.Color]::Purple }
            "MSI + MST" { [System.Drawing.Color]::DarkGreen }
            "MSI Only" { [System.Drawing.Color]::Blue }
            "EXE File" { [System.Drawing.Color]::DarkOrange }
            "Unknown" { [System.Drawing.Color]::Red }
            default { [System.Drawing.Color]::Gray }
        }

        $global:packageStatusLabel.Text = $statusText
        $global:packageStatusLabel.ForeColor = $color

        Write-Host "GUI Status Updated: $statusText" -ForegroundColor Green
    }
}

function Copy-LocalPackageToTemp {
    <#
    .SYNOPSIS
        Copies a local package file to C:\temp\<PackageName>\<Version>\DRMBUILD\
    .PARAMETER localFilePath
        Path to the local package file
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$localFilePath
    )

    try {
        if (-not (Test-Path $localFilePath)) {
            Write-Host "Error: Local file not found: $localFilePath" -ForegroundColor Red
            return $false
        }

        # Extract file info
        $fileInfo = Get-Item $localFilePath
        $fileName = $fileInfo.Name
        $fileBaseName = $fileInfo.BaseName
        $fileExtension = $fileInfo.Extension

        # Create a package structure in C:\temp
        # Use filename as application name if not available from dropdowns
        $appName = if ($global:selectedApplication) { $global:selectedApplication } else { $fileBaseName }
        $version = if ($global:selectedVersion) { $global:selectedVersion } else { "1.0.0" }

        $destinationPath = "C:\temp\$appName\$version\DRMBUILD"

        Write-Host "=== LOCAL PACKAGE COPY ===" -ForegroundColor Cyan
        Write-Host "Source: $localFilePath" -ForegroundColor Yellow
        Write-Host "Destination: $destinationPath" -ForegroundColor Yellow

        # Create destination directory
        if (-not (Test-Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
            Write-Host "Created destination directory: $destinationPath" -ForegroundColor Green
        }

        # Copy the file
        $destinationFile = Join-Path $destinationPath $fileName
        Copy-Item $localFilePath $destinationFile -Force
        Write-Host "Local package copied successfully!" -ForegroundColor Green

        # Update global variables
        $global:localPackagePath = $destinationPath
        $global:downloadCompleted = $true

        # Update package file paths to point to local copy
        switch ($fileExtension.ToLower()) {
            ".msi" { $global:filemsi = $destinationFile }
            ".exe" { $global:fileexe = $destinationFile }
            ".ps1" { $global:fileps1 = $destinationFile }
        }

        Write-Host "Package file paths updated to local copy" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Error copying local package: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Invoke-PackageDownload {
    <#
    .SYNOPSIS
        Downloads the entire package to C:\temp\<PackageName>\<Version>\<DRMBuild>\
    #>

    try {
        # Debug: Check package selection variables
        Write-Host "=== DOWNLOAD DEBUG INFO ===" -ForegroundColor Magenta
        Write-Host "Vendor: '$global:selectedVendor'" -ForegroundColor Gray
        Write-Host "Application: '$global:selectedApplication'" -ForegroundColor Gray
        Write-Host "Version: '$global:selectedVersion'" -ForegroundColor Gray
        Write-Host "Deployment: '$global:selectedDeployment'" -ForegroundColor Gray
        Write-Host "DRM Build: '$global:selectedDRMBuild'" -ForegroundColor Gray
        Write-Host "Package Location: '$global:selectedPackageLocation'" -ForegroundColor Gray

        if (-not ($global:selectedVendor -and $global:selectedApplication -and $global:selectedVersion -and $global:selectedDeployment -and $global:selectedDRMBuild)) {
            Write-Host "Error: Package selection incomplete" -ForegroundColor Red
            Write-Host "Missing variables detected - please ensure all dropdowns are selected" -ForegroundColor Red
            return $false
        }

        # Build source path
        $basePath = switch ($global:selectedPackageLocation) {
            "PackageArchive" { $global:sharedSoftwarePath }
            "CompletedPackage" { "\\nas00036pn\Cert-Staging\2_Completed Packages" }
            default { $global:sharedSoftwarePath }
        }

        $sourcePath = "$basePath\$global:selectedVendor\$global:selectedApplication\$global:selectedVersion\$global:selectedDeployment\$global:selectedDRMBuild"

        # Check if source path exists
        Write-Host "Checking source path accessibility..." -ForegroundColor Yellow
        if (-not (Test-Path $sourcePath)) {
            Write-Host "Error: Source path not accessible: $sourcePath" -ForegroundColor Red
            Write-Host "Please verify the package path exists and is accessible" -ForegroundColor Red
            return $false
        }
        Write-Host "Source path verified: $sourcePath" -ForegroundColor Green

        # Build destination path: C:\temp\<PackageName>\<Version>\<DRMBuild>\
        $destinationPath = "C:\temp\$global:selectedApplication\$global:selectedVersion\$global:selectedDRMBuild"

        Write-Host "=== PACKAGE DOWNLOAD ===" -ForegroundColor Cyan
        Write-Host "Source: $sourcePath" -ForegroundColor Yellow
        Write-Host "Destination: $destinationPath" -ForegroundColor Yellow

        # Update GUI status
        if ($global:packageStatusLabel) {
            $global:packageStatusLabel.Text = "Status: Downloading package..."
            $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Orange
        }

        # Create destination directory
        if (-not (Test-Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
            Write-Host "Created destination directory: $destinationPath" -ForegroundColor Green
        }

        # Copy package files using Robocopy for reliability
        Write-Host "Copying package files..." -ForegroundColor Yellow

        # Ensure paths are properly quoted for ROBOCOPY
        $quotedSourcePath = "`"$sourcePath`""
        $quotedDestinationPath = "`"$destinationPath`""

        Write-Host "ROBOCOPY Debug:" -ForegroundColor Magenta
        Write-Host "  Quoted Source: $quotedSourcePath" -ForegroundColor Gray
        Write-Host "  Quoted Destination: $quotedDestinationPath" -ForegroundColor Gray

        $robocopyArgs = @(
            $quotedSourcePath,
            $quotedDestinationPath,
            "/E",           # Copy subdirectories including empty ones
            "/R:3",         # Retry 3 times on failed copies
            "/W:5",         # Wait 5 seconds between retries
            "/MT:8",        # Multi-threaded copy (8 threads)
            "/LOG:C:\temp\robocopy.log"
        )

        Write-Host "ROBOCOPY Command: robocopy $($robocopyArgs -join ' ')" -ForegroundColor Gray
        $robocopyResult = Start-Process "robocopy" -ArgumentList $robocopyArgs -Wait -PassThru -NoNewWindow

        # Check robocopy result (exit codes 0-7 are success, 8+ are errors)
        if ($robocopyResult.ExitCode -le 7) {
            Write-Host "Package download completed successfully!" -ForegroundColor Green
            $global:localPackagePath = $destinationPath
            $global:downloadCompleted = $true

            # Update package file paths to local copies
            Update-LocalPackagePaths -localPath $destinationPath

            # Update GUI status
            if ($global:packageStatusLabel) {
                $global:packageStatusLabel.Text = "Status: Package downloaded - $global:packageType"
                $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Green
            }

            return $true
        }
        else {
            Write-Host "ROBOCOPY failed with exit code: $($robocopyResult.ExitCode)" -ForegroundColor Red
            Write-Host "Attempting fallback copy method..." -ForegroundColor Yellow

            try {
                # Fallback: Use PowerShell Copy-Item with force
                Write-Host "Using PowerShell Copy-Item as fallback..." -ForegroundColor Yellow
                Copy-Item -Path "$sourcePath\*" -Destination $destinationPath -Recurse -Force -ErrorAction Stop

                Write-Host "Fallback copy completed successfully!" -ForegroundColor Green
                $global:localPackagePath = $destinationPath
                $global:downloadCompleted = $true

                # Update package file paths to local copies
                Update-LocalPackagePaths -localPath $destinationPath

                # Update GUI status
                if ($global:packageStatusLabel) {
                    $global:packageStatusLabel.Text = "Status: Package downloaded - $global:packageType"
                    $global:packageStatusLabel.ForeColor = [System.Drawing.Color]::Green
                }

                return $true
            }
            catch {
                Write-Host "Fallback copy also failed: $($_.Exception.Message)" -ForegroundColor Red
                return $false
            }
        }
    }
    catch {
        Write-Host "Error during package download: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Update-LocalPackagePaths {
    <#
    .SYNOPSIS
        Updates global package file paths to point to local downloaded copies
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$localPath
    )

    try {
        # Scan local directory for package files
        $localMsiFiles = Get-ChildItem $localPath -Filter "*.msi" -ErrorAction SilentlyContinue
        $localExeFiles = Get-ChildItem $localPath -Filter "*.exe" -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "ServiceUI.exe" }
        $localPs1Files = Get-ChildItem $localPath -Filter "*.ps1" -ErrorAction SilentlyContinue
        $localMstFiles = Get-ChildItem $localPath -Filter "*.mst" -ErrorAction SilentlyContinue

        # Update global paths to local copies
        if ($localMsiFiles.Count -gt 0) {
            $global:filemsi = $localMsiFiles[0].FullName
            Write-Host "Updated MSI path: $global:filemsi" -ForegroundColor Green
        }

        if ($localExeFiles.Count -gt 0) {
            $global:fileexe = $localExeFiles[0].FullName
            Write-Host "Updated EXE path: $global:fileexe" -ForegroundColor Green
        }

        if ($localPs1Files.Count -gt 0) {
            $global:fileps1 = $localPs1Files[0].FullName
            Write-Host "Updated PS1 path: $global:fileps1" -ForegroundColor Green
        }

        if ($localMstFiles.Count -gt 0) {
            $global:filemst = $localMstFiles[0].FullName
            Write-Host "Updated MST path: $global:filemst" -ForegroundColor Green
        }

        # Rebuild install command with local paths
        Determine-PackageType

        Write-Host "Local package paths updated successfully" -ForegroundColor Green
    }
    catch {
        Write-Host "Error updating local package paths: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Invoke-TaskSchedulerInstall {
    <#
    .SYNOPSIS
        Executes package installation using SCHTASKS command (like original implementation)
    #>

    try {
        if (-not $global:downloadCompleted -and -not $global:localPackagePath) {
            Write-Host "Error: Package not downloaded and no package path available" -ForegroundColor Red
            return $false
        }

        # Handle both downloaded and network-based installations
        if ($global:downloadCompleted) {
            Write-Host "Using downloaded package from: $global:localPackagePath" -ForegroundColor Green
        }
        else {
            Write-Host "Using network package from: $global:localPackagePath" -ForegroundColor Yellow
            Write-Host "Note: Installing directly from network location" -ForegroundColor Yellow
        }

        $taskName = "PDM_Install_$([System.Guid]::NewGuid().ToString().Substring(0,8))"

        Write-Host "=== TASK SCHEDULER EXECUTION (SCHTASKS) ===" -ForegroundColor Cyan
        Write-Host "Task Name: $taskName" -ForegroundColor Yellow
        Write-Host "Package Type: $global:packageType" -ForegroundColor Yellow
        Write-Host "Install Command: $global:installCommand" -ForegroundColor Yellow

        # Create batch file for execution (simpler approach like original)
        $batchPath = "C:\temp\install_$([System.Guid]::NewGuid().ToString().Substring(0,8)).bat"

        # Create batch content based on package type
        $batchContent = @"
@echo off
echo === Package Development Manager - Installation ===
echo Package: $global:selectedApplication $global:selectedVersion
echo Type: $global:packageType
echo Time: %DATE% %TIME%
echo.

cd /d "$global:localPackagePath"

echo Executing installation command...
echo Command: $global:installCommand
echo.

"@

        # Use the complete install command (preserves MSI logging parameters)
        Write-Host "Using install command: $global:installCommand" -ForegroundColor Yellow
        $batchContent += "$global:installCommand`n"

        $batchContent += @"

echo.
echo Installation completed at %DATE% %TIME%
echo Logging result...
echo %DATE% %TIME%: Installation completed - $global:selectedApplication $global:selectedVersion ($global:packageType) >> C:\temp\PackageDevManager_InstallLog.txt

timeout /t 3 /nobreak >nul
del "$batchPath" >nul 2>&1
"@

        # Write batch file
        $batchContent | Out-File -FilePath $batchPath -Encoding ASCII
        Write-Host "Created installation batch: $batchPath" -ForegroundColor Green

        # Use SCHTASKS command (like original implementation)
        $schtasksCmd = "schtasks /create /tn `"$taskName`" /tr `"$batchPath`" /sc once /st 00:00 /ru SYSTEM /rl HIGHEST /f"
        Write-Host "Creating task with SCHTASKS..." -ForegroundColor Gray

        $createResult = cmd /c $schtasksCmd 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Task created successfully" -ForegroundColor Green

            # Run task immediately
            $runCmd = "schtasks /run /tn `"$taskName`""
            Write-Host "Starting task..." -ForegroundColor Gray
            $runResult = cmd /c $runCmd 2>&1

            if ($LASTEXITCODE -eq 0) {
                Write-Host "Task started with SYSTEM privileges" -ForegroundColor Green

                # Wait for completion (simpler approach)
                Write-Host "Waiting for installation to complete..." -ForegroundColor Yellow
                Start-Sleep -Seconds 10  # Give it time to start

                # Check if batch file still exists (it deletes itself when done)
                $timeout = 300  # 5 minutes
                $elapsed = 0
                while ((Test-Path $batchPath) -and $elapsed -lt $timeout) {
                    Start-Sleep -Seconds 5
                    $elapsed += 5
                    Write-Host "Installation in progress... ($elapsed seconds)" -ForegroundColor Gray
                }

                # Clean up task
                $deleteCmd = "schtasks /delete /tn `"$taskName`" /f"
                cmd /c $deleteCmd 2>&1 | Out-Null
                Write-Host "Task cleaned up" -ForegroundColor Green

                return $true
            }
            else {
                Write-Host "Failed to start task: $runResult" -ForegroundColor Red
                return $false
            }
        }
        else {
            Write-Host "Failed to create task: $createResult" -ForegroundColor Red
            return $false
        }

        return $true
    }
    catch {
        Write-Host "Error in task scheduler execution: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Build-InstallCommand {
    <#
    .SYNOPSIS
        Builds appropriate install command based on package type and operation
    .PARAMETER Operation
        Install or Uninstall operation
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Install", "Uninstall")]
        [string]$Operation
    )

    try {
        switch ($global:packageType) {
            "PSADTK Wrapper" {
                # PSADTK with ServiceUI.exe wrapper or PowerShell direct
                if ($global:exeExist) {
                    # Use ServiceUI.exe wrapper for PSADTK
                    $baseCommand = "`"$global:fileexe`" `"$global:fileps1`" -DeployMode Silent -DeploymentType $Operation"
                    if ($global:exeArguments) {
                        return "$baseCommand $global:exeArguments"
                    } else {
                        return $baseCommand
                    }
                }
                else {
                    # Fallback to PowerShell direct execution
                    $baseCommand = "PowerShell.exe -ExecutionPolicy Bypass -File `"$global:fileps1`" -DeployMode Silent -DeploymentType $Operation"
                    if ($global:exeArguments) {
                        return "$baseCommand $global:exeArguments"
                    } else {
                        return $baseCommand
                    }
                }
            }

            "MSI + MST" {
                # Create UHG logs directory
                $uhgLogsDir = "C:\Windows\UHGLOGS"
                if (-not (Test-Path $uhgLogsDir)) {
                    New-Item -ItemType Directory -Path $uhgLogsDir -Force | Out-Null
                }

                $logFile = "$uhgLogsDir\$($global:selectedApplication)_$($global:selectedVersion)_$Operation.log"

                if ($Operation -eq "Install") {
                    $baseArgs = "/i `"$global:filemsi`" TRANSFORMS=`"$global:filemst`" /qn /norestart /l*v `"$logFile`""
                    if ($global:exeArguments) {
                        return "msiexec.exe $baseArgs $global:exeArguments"
                    } else {
                        return "msiexec.exe $baseArgs"
                    }
                }
                else {
                    return "msiexec.exe /x `"$global:filemsi`" TRANSFORMS=`"$global:filemst`" /qn /norestart /l*v `"$logFile`""
                }
            }

            "MSI Only" {
                # Create UHG logs directory
                $uhgLogsDir = "C:\Windows\UHGLOGS"
                if (-not (Test-Path $uhgLogsDir)) {
                    New-Item -ItemType Directory -Path $uhgLogsDir -Force | Out-Null
                }

                $logFile = "$uhgLogsDir\$($global:selectedApplication)_$($global:selectedVersion)_$Operation.log"

                if ($Operation -eq "Install") {
                    $baseArgs = "/i `"$global:filemsi`" /qn /norestart /l*v `"$logFile`""
                    if ($global:exeArguments) {
                        return "msiexec.exe $baseArgs $global:exeArguments"
                    } else {
                        return "msiexec.exe $baseArgs"
                    }
                }
                else {
                    return "msiexec.exe /x `"$global:filemsi`" /qn /norestart /l*v `"$logFile`""
                }
            }

            "EXE File" {
                if ($Operation -eq "Install") {
                    if ($global:exeArguments) {
                        return "`"$global:fileexe`" $global:exeArguments"
                    } else {
                        return "`"$global:fileexe`" /S"
                    }
                }
                else {
                    return "`"$global:fileexe`" /uninstall "
                }
            }

            default {
                Write-Host "Unknown package type for $Operation`: $global:packageType" -ForegroundColor Yellow
                return $null
            }
        }
    }
    catch {
        Write-Host "Error building $Operation command: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Build-UninstallCommand {
    <#
    .SYNOPSIS
        Builds appropriate uninstall command based on package type
    #>

    try {
        switch ($global:packageType) {
            "PSADTK Wrapper" {
                # PSADTK uninstall using ServiceUI.exe wrapper or PowerShell direct
                if ($global:exeExist) {
                    # Use ServiceUI.exe wrapper for PSADTK uninstall
                    return "`"$global:fileexe`" `"$global:fileps1`" -DeployMode Silent -DeploymentType Uninstall"
                }
                else {
                    # Fallback to PowerShell direct execution
                    return "PowerShell.exe -ExecutionPolicy Bypass -File `"$global:fileps1`" -DeployMode Silent -DeploymentType Uninstall"
                }
            }

            "MSI + MST" {
                # Create UHG logs directory
                $uhgLogsDir = "C:\Windows\UHGLOGS"
                if (-not (Test-Path $uhgLogsDir)) {
                    New-Item -ItemType Directory -Path $uhgLogsDir -Force | Out-Null
                }

                $logFile = "$uhgLogsDir\$($global:selectedApplication)_$($global:selectedVersion)_Uninstall.log"
                return "msiexec.exe /x `"$global:filemsi`" TRANSFORMS=`"$global:filemst`" /qn /norestart /l*v `"$logFile`""
            }

            "MSI Only" {
                # Create UHG logs directory
                $uhgLogsDir = "C:\Windows\UHGLOGS"
                if (-not (Test-Path $uhgLogsDir)) {
                    New-Item -ItemType Directory -Path $uhgLogsDir -Force | Out-Null
                }

                $logFile = "$uhgLogsDir\$($global:selectedApplication)_$($global:selectedVersion)_Uninstall.log"
                return "msiexec.exe /x `"$global:filemsi`" /qn /norestart /l*v `"$logFile`""
            }

            "EXE File" {
                # EXE uninstall - try common switches
                $exeName = [System.IO.Path]::GetFileName($global:fileexe)
                return "`"$global:fileexe`" /uninstall /S /silent"
            }

            default {
                Write-Host "Unknown package type for uninstall: $global:packageType" -ForegroundColor Yellow
                return $null
            }
        }
    }
    catch {
        Write-Host "Error building uninstall command: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Invoke-TaskSchedulerUninstall {
    <#
    .SYNOPSIS
        Executes package uninstallation using SCHTASKS command (like original implementation)
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$uninstallCommand
    )

    try {
        $taskName = "PDM_Uninstall_$([System.Guid]::NewGuid().ToString().Substring(0,8))"

        Write-Host "=== TASK SCHEDULER UNINSTALLATION (SCHTASKS) ===" -ForegroundColor Cyan
        Write-Host "Task Name: $taskName" -ForegroundColor Yellow
        Write-Host "Uninstall Command: $uninstallCommand" -ForegroundColor Yellow

        # Create batch file for uninstall execution
        $batchPath = "C:\temp\uninstall_$([System.Guid]::NewGuid().ToString().Substring(0,8)).bat"

        # Create batch content for uninstall
        $batchContent = @"
@echo off
echo === Package Development Manager - Uninstallation ===
echo Package: $global:selectedApplication $global:selectedVersion
echo Type: $global:packageType
echo Time: %DATE% %TIME%
echo.

"@

        # Add working directory if available
        if ($global:localPackagePath -and (Test-Path $global:localPackagePath)) {
            $batchContent += "cd /d `"$global:localPackagePath`"`n"
        }

        $batchContent += @"
echo Executing uninstallation command...
echo Command: $uninstallCommand
echo.

$uninstallCommand

echo.
echo Uninstallation completed at %DATE% %TIME%
echo Logging result...
echo %DATE% %TIME%: Uninstallation completed - $global:selectedApplication $global:selectedVersion ($global:packageType) >> C:\temp\PackageDevManager_InstallLog.txt

timeout /t 3 /nobreak >nul
del "$batchPath" >nul 2>&1
"@

        # Write batch file
        $batchContent | Out-File -FilePath $batchPath -Encoding ASCII
        Write-Host "Created uninstallation batch: $batchPath" -ForegroundColor Green

        # Use SCHTASKS command (same as install)
        $schtasksCmd = "schtasks /create /tn `"$taskName`" /tr `"$batchPath`" /sc once /st 00:00 /ru SYSTEM /rl HIGHEST /f"
        Write-Host "Creating uninstall task with SCHTASKS..." -ForegroundColor Gray

        $createResult = cmd /c $schtasksCmd 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Uninstall task created successfully" -ForegroundColor Green

            # Run task immediately
            $runCmd = "schtasks /run /tn `"$taskName`""
            Write-Host "Starting uninstall task..." -ForegroundColor Gray
            $runResult = cmd /c $runCmd 2>&1

            if ($LASTEXITCODE -eq 0) {
                Write-Host "Uninstall task started with SYSTEM privileges" -ForegroundColor Green

                # Wait for completion
                Write-Host "Waiting for uninstallation to complete..." -ForegroundColor Yellow
                Start-Sleep -Seconds 10  # Give it time to start

                # Check if batch file still exists (it deletes itself when done)
                $timeout = 300  # 5 minutes
                $elapsed = 0
                while ((Test-Path $batchPath) -and $elapsed -lt $timeout) {
                    Start-Sleep -Seconds 5
                    $elapsed += 5
                    Write-Host "Uninstallation in progress... ($elapsed seconds)" -ForegroundColor Gray
                }

                # Clean up task
                $deleteCmd = "schtasks /delete /tn `"$taskName`" /f"
                cmd /c $deleteCmd 2>&1 | Out-Null
                Write-Host "Uninstall task cleaned up" -ForegroundColor Green

                return $true
            }
            else {
                Write-Host "Failed to start uninstall task: $runResult" -ForegroundColor Red
                return $false
            }
        }
        else {
            Write-Host "Failed to create uninstall task: $createResult" -ForegroundColor Red
            return $false
        }

        return $true
    }
    catch {
        Write-Host "Error in task scheduler uninstallation: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Generate-CreativeHTMLReport {
    <#
    .SYNOPSIS
        Generates a comprehensive, creative HTML report for Package Development Manager
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$OperationType = "Validation"
    )

    try {
        # Determine report location - save in package folder with proper naming
        $reportFolder = if ($global:localPackagePath) {
            $global:localPackagePath
        } else {
            "$env:SystemDrive\temp\AutoTest\Tool_Output"
        }

        # Create report name using same standard as logs with operation type
        $reportName = if ($global:selectedApplication -and $global:selectedVersion) {
            "$($global:selectedApplication)_$($global:selectedVersion)_$($OperationType)Report.html"
        } else {
            "PackageDevManager_$($OperationType)Report.html"
        }

        $reportPath = Join-Path $reportFolder $reportName
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        Write-Host "Generating comprehensive HTML report..." -ForegroundColor Cyan
        Write-Host "Report location: $reportPath" -ForegroundColor Yellow

        # Create comprehensive HTML report
        $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Package Deployment Manager 2.0 - Validation Report</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }

        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }

        .container {
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }

        .header {
            background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
            color: white;
            padding: 30px;
            text-align: center;
            position: relative;
        }

        .header::before {
            content: '';
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: url('data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><circle cx="20" cy="20" r="2" fill="rgba(255,255,255,0.1)"/><circle cx="80" cy="40" r="3" fill="rgba(255,255,255,0.1)"/><circle cx="40" cy="80" r="2" fill="rgba(255,255,255,0.1)"/></svg>');
        }

        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            position: relative;
            z-index: 1;
        }

        .header .subtitle {
            font-size: 1.2em;
            opacity: 0.9;
            position: relative;
            z-index: 1;
        }

        .optum-logo {
            position: absolute;
            top: 20px;
            right: 30px;
            background: rgba(255,255,255,0.2);
            padding: 10px 20px;
            border-radius: 25px;
            font-weight: bold;
            font-size: 1.1em;
        }

        .content {
            padding: 40px;
        }

        .info-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }

        .info-card {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
            padding: 25px;
            border-radius: 12px;
            box-shadow: 0 8px 16px rgba(0,0,0,0.1);
            transform: translateY(0);
            transition: transform 0.3s ease;
        }

        .info-card:hover {
            transform: translateY(-5px);
        }

        .info-card h3 {
            font-size: 1.3em;
            margin-bottom: 15px;
            display: flex;
            align-items: center;
        }

        .info-card .icon {
            width: 24px;
            height: 24px;
            margin-right: 10px;
            background: rgba(255,255,255,0.3);
            border-radius: 50%;
            display: inline-flex;
            align-items: center;
            justify-content: center;
        }

        .section {
            margin-bottom: 40px;
            background: #f8f9fa;
            border-radius: 12px;
            padding: 30px;
            border-left: 5px solid #007bff;
        }

        .section h2 {
            color: #2c3e50;
            margin-bottom: 20px;
            font-size: 1.8em;
            display: flex;
            align-items: center;
        }

        .section h2::before {
            content: '🔍';
            margin-right: 10px;
            font-size: 1.2em;
        }

        .validation-item {
            background: white;
            margin: 10px 0;
            padding: 15px;
            border-radius: 8px;
            border-left: 4px solid #28a745;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            display: flex;
            justify-content: space-between;
            align-items: center;
        }

        .validation-item.invalid {
            border-left-color: #dc3545;
        }

        .validation-item.warning {
            border-left-color: #ffc107;
        }

        .status-badge {
            padding: 5px 12px;
            border-radius: 20px;
            font-size: 0.9em;
            font-weight: bold;
            text-transform: uppercase;
        }

        .status-valid {
            background: #d4edda;
            color: #155724;
        }

        .status-invalid {
            background: #f8d7da;
            color: #721c24;
        }

        .status-warning {
            background: #fff3cd;
            color: #856404;
        }

        .timeline {
            position: relative;
            padding-left: 30px;
        }

        .timeline::before {
            content: '';
            position: absolute;
            left: 15px;
            top: 0;
            bottom: 0;
            width: 2px;
            background: #007bff;
        }

        .timeline-item {
            position: relative;
            margin-bottom: 20px;
            background: white;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }

        .timeline-item::before {
            content: '';
            position: absolute;
            left: -37px;
            top: 25px;
            width: 12px;
            height: 12px;
            background: #007bff;
            border-radius: 50%;
            border: 3px solid white;
        }

        .footer {
            background: #2c3e50;
            color: white;
            padding: 20px;
            text-align: center;
        }

        .no-data {
            text-align: center;
            color: #6c757d;
            font-style: italic;
            padding: 40px;
        }

        @media (max-width: 768px) {
            .info-grid {
                grid-template-columns: 1fr;
            }

            .header h1 {
                font-size: 2em;
            }

            .optum-logo {
                position: static;
                margin-top: 20px;
                display: inline-block;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="optum-logo">Optum</div>
            <h1>📦 Package Deployment Manager 2.0</h1>
            <div class="subtitle">Comprehensive Installation & Validation Report</div>
        </div>

        <div class="content">
            <div class="info-grid">
                <div class="info-card">
                    <h3><span class="icon">📋</span>Package Information</h3>
                    <p><strong>Application:</strong> $global:selectedApplication</p>
                    <p><strong>Version:</strong> $global:selectedVersion</p>
                    <p><strong>Vendor:</strong> $global:selectedVendor</p>
                    <p><strong>Package Type:</strong> $global:packageType</p>
                </div>

                <div class="info-card">
                    <h3><span class="icon">⚙️</span>Deployment Details</h3>
                    <p><strong>Deployment:</strong> $global:selectedDeployment</p>
                    <p><strong>DRM Build:</strong> $global:selectedDRMBuild</p>
                    <p><strong>Location:</strong> $global:selectedPackageLocation</p>
                    <p><strong>Local Path:</strong> $global:localPackagePath</p>
                </div>

                <div class="info-card">
                    <h3><span class="icon">📊</span>Report Details</h3>
                    <p><strong>Generated:</strong> $timestamp</p>
                    <p><strong>Machine:</strong> $env:COMPUTERNAME</p>
                    <p><strong>User:</strong> $env:USERNAME</p>
                    <p><strong>Domain:</strong> $env:USERDOMAIN</p>
                </div>
            </div>
"@



        # Validation Results section removed as requested - it was giving no results

        # Add package files section
        if ($global:packageType -and $global:packageType -ne "Unknown") {
            $htmlContent += @"
            <div class="section">
                <h2>Package Files Detected</h2>
"@

            if ($global:filemsi) {
                $msiName = [System.IO.Path]::GetFileName($global:filemsi)
                $htmlContent += @"
                <div class="validation-item">
                    <span>📦 MSI Package: $msiName</span>
                    <span class="status-badge status-valid">Found</span>
                </div>
"@
            }

            if ($global:filemst) {
                $mstName = [System.IO.Path]::GetFileName($global:filemst)
                $htmlContent += @"
                <div class="validation-item">
                    <span>🔧 MST Transform: $mstName</span>
                    <span class="status-badge status-valid">Found</span>
                </div>
"@
            }

            if ($global:fileexe) {
                $exeName = [System.IO.Path]::GetFileName($global:fileexe)
                $htmlContent += @"
                <div class="validation-item">
                    <span>⚙️ EXE Installer: $exeName</span>
                    <span class="status-badge status-valid">Found</span>
                </div>
"@
            }

            if ($global:fileps1) {
                $ps1Name = [System.IO.Path]::GetFileName($global:fileps1)
                $htmlContent += @"
                <div class="validation-item">
                    <span>📜 PowerShell Script: $ps1Name</span>
                    <span class="status-badge status-valid">Found</span>
                </div>
"@
            }

            $htmlContent += @"
            </div>
"@
        }

        # Add command section based on operation type
        $commandToShow = ""
        $commandTitle = ""

        if ($OperationType -eq "Install" -and $global:installCommand) {
            $commandToShow = $global:installCommand
            $commandTitle = "Installation Command"
        }
        elseif ($OperationType -eq "Uninstall" -and $global:packageType) {
            # Get the uninstall command before using it in the here-string
            $uninstallCmd = Build-UninstallCommand
            $commandToShow = $uninstallCmd
            $commandTitle = "Uninstallation Command"
        }

        if ($commandToShow) {
            # Simple HTML escaping for special characters
            $escapedCommand = $commandToShow -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;'

            $htmlContent += @"
            <div class="section">
                <h2>$commandTitle</h2>
                <div class="validation-item">
                    <span style="font-family: monospace; background: #f8f9fa; padding: 10px; border-radius: 4px; display: block;">$escapedCommand</span>
                </div>
            </div>
"@
        }

        # Add comprehensive validation sections
        $htmlContent += Generate-ValidationSections -OperationType $OperationType

        # Close HTML
        $htmlContent += @"
        </div>

        <div class="footer">
            <p>&copy; 2025 Developed by Bharadwaj @ GPS Packaging Team | Generated on $timestamp</p>
            <p>🏢 Optum - Automated Package Validation System</p>
        </div>
    </div>

    <script>
        // Add some interactive effects
        document.addEventListener('DOMContentLoaded', function() {
            const cards = document.querySelectorAll('.info-card');
            cards.forEach(card => {
                card.addEventListener('mouseenter', function() {
                    this.style.transform = 'translateY(-5px) scale(1.02)';
                });
                card.addEventListener('mouseleave', function() {
                    this.style.transform = 'translateY(0) scale(1)';
                });
            });
        });
    </script>
</body>
</html>
"@

        # Write the HTML report
        $htmlContent | Out-File -FilePath $reportPath -Encoding UTF8
        Write-Host "Creative HTML report generated: $reportPath" -ForegroundColor Green

        return $reportPath
    }
    catch {
        Write-Host "Error generating HTML report: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Generate-ValidationSections {
    <#
    .SYNOPSIS
        Generates comprehensive validation sections for HTML report
    .PARAMETER OperationType
        Type of operation: Install or Uninstall
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$OperationType = "Install"
    )

    $sectionsHtml = ""

    # Determine status labels based on operation type
    $fileSystemStatus = if ($OperationType -eq "Uninstall") { "Removed" } else { "Added" }
    $registryStatus = if ($OperationType -eq "Uninstall") { "Removed" } else { "Added" }
    $shortcutStatus = if ($OperationType -eq "Uninstall") { "Deleted" } else { "Created" }
    $uhgLogStatus = if ($OperationType -eq "Uninstall") { "Removed" } else { "Created" }

    # FileSystem Section
    $sectionsHtml += @"
            <div class="section">
                <h2>📁 FileSystem Changes</h2>
"@

    $diffProgramFiles = "$tempPath\diffProgramFiles.txt"
    if (Test-Path $diffProgramFiles) {
        $fileChanges = Get-Content $diffProgramFiles -ErrorAction SilentlyContinue
        if ($fileChanges) {
            foreach ($change in $fileChanges) {
                $sectionsHtml += @"
                <div class="validation-item">
                    <span>📂 $change</span>
                    <span class="status-badge status-valid">$fileSystemStatus</span>
                </div>
"@
            }
        }
        else {
            $sectionsHtml += @"
                <div class="no-data">
                    <p>No file system changes detected</p>
                </div>
"@
        }
    }

    $sectionsHtml += "</div>"

    # Registry Section
    $sectionsHtml += @"
            <div class="section">
                <h2>📋 Registry Changes</h2>
"@

    # Use the generated diff files for registry changes
    $hasRegistryChanges = $false

    # Check for new uninstall keys (Add/Remove Programs entries)
    $diffUninstallKey = "$tempPath\diffUninstallKey.txt"
    if (Test-Path $diffUninstallKey) {
        try {
            $newUninstallEntries = Get-Content $diffUninstallKey -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
            if ($newUninstallEntries) {
                $hasRegistryChanges = $true
                foreach ($entry in $newUninstallEntries) {
                    $regPath = $entry.PSPath -replace "Microsoft.PowerShell.Core\\Registry::", ""
                    $sectionsHtml += @"
                <div class="validation-item">
                    <span>🔑 $regPath\$($entry.DisplayName) (Version: $($entry.DisplayVersion), Publisher: $($entry.Publisher))</span>
                    <span class="status-badge status-valid">$registryStatus</span>
                </div>
"@
                }
            }
        }
        catch {
            Write-LogEntry "Error reading uninstall registry data: $($_.Exception.Message)" -LogType "Warning"
        }
    }

    # Check for other registry area changes
    $diffRegistryAreas = "$tempPath\diffRegistryAreas.txt"
    if (Test-Path $diffRegistryAreas) {
        try {
            $newRegistryEntries = Get-Content $diffRegistryAreas -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
            if ($newRegistryEntries) {
                $hasRegistryChanges = $true
                foreach ($entry in $newRegistryEntries) {
                    $regPath = $entry.Path -replace "Microsoft.PowerShell.Core\\Registry::", ""
                    $sectionsHtml += @"
                <div class="validation-item">
                    <span>🔑 $regPath\$($entry.Name) = $($entry.Value)</span>
                    <span class="status-badge status-valid">$registryStatus</span>
                </div>
"@
                }
            }
        }
        catch {
            Write-LogEntry "Error reading registry areas data: $($_.Exception.Message)" -LogType "Warning"
        }
    }

    if (-not $hasRegistryChanges) {
        $sectionsHtml += @"
                <div class="no-data">
                    <p>No new registry entries detected</p>
                </div>
"@
    }

    $sectionsHtml += "</div>"

    # Shortcuts Section
    $sectionsHtml += @"
            <div class="section">
                <h2>🔗 Shortcuts (Public Desktop & Start Menu)</h2>
"@

    $diffShortcuts = "$tempPath\diffShortcuts.txt"
    if (Test-Path $diffShortcuts) {
        $shortcutChanges = Get-Content $diffShortcuts -ErrorAction SilentlyContinue
        if ($shortcutChanges) {
            foreach ($shortcut in $shortcutChanges) {
                $sectionsHtml += @"
                <div class="validation-item">
                    <span>🔗 $shortcut</span>
                    <span class="status-badge status-valid">$shortcutStatus</span>
                </div>
"@
            }
        }
        else {
            $sectionsHtml += @"
                <div class="no-data">
                    <p>No shortcuts created</p>
                </div>
"@
        }
    }

    $sectionsHtml += "</div>"

    # ARP Section
    $sectionsHtml += @"
            <div class="section">
                <h2>🗂️ Add/Remove Programs (32bit & 64bit)</h2>
"@

    if (Test-Path "$tempPath\AfUninstallKey.txt") {
        try {
            $afterARP = Get-Content "$tempPath\AfUninstallKey.txt" -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
            $beforeARP = @()
            if (Test-Path "$tempPath\BfUninstallKey.txt") {
                $beforeARP = Get-Content "$tempPath\BfUninstallKey.txt" -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
            }

            # Find new ARP entries
            $newEntries = $afterARP | Where-Object {
                $afterEntry = $_
                -not ($beforeARP | Where-Object { $_.DisplayName -eq $afterEntry.DisplayName })
            }

            if ($newEntries) {
                foreach ($entry in $newEntries) {
                    # Format vendor and version information
                    $vendor = if ($entry.Publisher) { $entry.Publisher } else { "Unknown Vendor" }
                    $version = if ($entry.DisplayVersion) { $entry.DisplayVersion } else { "Unknown Version" }

                    $sectionsHtml += @"
                <div class="validation-item">
                    <span>📦 <strong>$($entry.DisplayName)</strong><br/>
                          &nbsp;&nbsp;&nbsp;&nbsp;Vendor: $vendor<br/>
                          &nbsp;&nbsp;&nbsp;&nbsp;Version: $version</span>
                    <span class="status-badge status-valid">Installed</span>
                </div>
"@
                }
            }
            else {
                $sectionsHtml += @"
                <div class="no-data">
                    <p>No new Add/Remove Programs entries</p>
                </div>
"@
            }
        }
        catch {
            $sectionsHtml += @"
                <div class="no-data">
                    <p>Error reading ARP data</p>
                </div>
"@
        }
    }

    $sectionsHtml += "</div>"

    # UHG Logs Section
    $sectionsHtml += @"
            <div class="section">
                <h2>📄 UHG Logs</h2>
"@

    $diffUHGLogs = "$tempPath\diffUHGLogs.txt"
    if (Test-Path $diffUHGLogs) {
        try {
            $logChanges = Get-Content $diffUHGLogs -Raw -ErrorAction SilentlyContinue | ConvertFrom-Json
            if ($logChanges) {
                foreach ($log in $logChanges) {
                    $logSize = if ($log.Length) { [math]::Round($log.Length / 1KB, 2) } else { 0 }
                    $sectionsHtml += @"
                <div class="validation-item">
                    <span>📄 $($log.Name) ($logSize KB) - $($log.LastWriteTime)</span>
                    <span class="status-badge status-valid">$uhgLogStatus</span>
                </div>
"@
                }
            }
            else {
                $sectionsHtml += @"
                <div class="no-data">
                    <p>No UHG log changes detected</p>
                </div>
"@
            }
        }
        catch {
            $sectionsHtml += @"
                <div class="no-data">
                    <p>Error reading UHG logs data</p>
                </div>
"@
        }
    }
    else {
        $sectionsHtml += @"
                <div class="no-data">
                    <p>No UHG log changes detected</p>
                </div>
"@
    }

    $sectionsHtml += "</div>"

    # Installation Timeline Section (moved to end)
    $sectionsHtml += @"
            <div class="section">
                <h2>⏱️ Installation Timeline</h2>
"@

    $installLog = "C:\temp\PackageDevManager_InstallLog.txt"
    if (Test-Path $installLog) {
        $logEntries = Get-Content $installLog -ErrorAction SilentlyContinue | Select-Object -Last 10
        if ($logEntries) {
            $sectionsHtml += @"
                <div class="timeline">
"@
            foreach ($entry in $logEntries) {
                $sectionsHtml += @"
                    <div class="timeline-item">
                        <strong>$entry</strong>
                    </div>
"@
            }
            $sectionsHtml += @"
                </div>
"@
        }
        else {
            $sectionsHtml += @"
                <div class="no-data">
                    <p>No timeline entries available</p>
                </div>
"@
        }
    }
    else {
        $sectionsHtml += @"
                <div class="no-data">
                    <p>Installation log not found</p>
                </div>
"@
    }

    $sectionsHtml += "</div>"

    return $sectionsHtml
}

# Embedded Core Functions (for standalone EXE)
# These functions are from PckgDevMngr_refactored_clean.ps1

#region Core Functions - Embedded for Standalone EXE

# Global Package Variables
$global:msiExist = $false
$global:mstExist = $false
$global:exeExist = $false
$global:ps1Exist = $false
$global:filemsi = ""
$global:filemst = ""
$global:fileexe = ""
$global:fileps1 = ""

# System Configuration
$identifyMachine = ($env:COMPUTERNAME).IndexOf("VRA")
$toolOutputDir = "$env:SystemDrive\temp\AutoTest\Tool_Output"
$installHistory = "$env:SystemDrive\temp\PackageDevHistory\Install"
$uninstallHistory = "$env:SystemDrive\temp\PackageDevHistory\UnInstall"
$trackPath = "\\Nasdslpps001\drm_pkging\Team\Members\Phase3\track.csv"

# System Paths
$shortcutPath = "$env:Programdata\Microsoft\Windows\Start Menu"
$shortcutPublicDesktop = "$env:PUBLIC\desktop"
$program86 = ${env:ProgramFiles(x86)}
$program64 = "C:\Program Files"
$tempPath = "$env:SystemDrive\temp\AutoTestScan"
$AppLogsPath = "$tempPath\AppLogs"

function Write-LogEntry {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [string]$LogType = "Info"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$LogType] $Message"

    # Write to console
    switch ($LogType) {
        "Error" { Write-Host $logEntry -ForegroundColor Red }
        "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
        "Info" { Write-Host $logEntry -ForegroundColor Green }
        default { Write-Host $logEntry }
    }

    # Write to log file
    $logFile = "$toolOutputDir\PackageDevManager.txt"
    if (-not (Test-Path (Split-Path $logFile))) {
        New-Item -ItemType Directory -Path (Split-Path $logFile) -Force | Out-Null
    }
    Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue
}

function Initialize-WorkingDirectories {
    $directories = @(
        $tempPath,
        "$tempPath\Logs",
        $AppLogsPath,
        $toolOutputDir,
        $installHistory,
        $uninstallHistory
    )

    foreach ($dir in $directories) {
        if (-not (Test-Path -Path $dir)) {
            try {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
                Write-LogEntry "Created directory: $dir"
            }
            catch {
                Write-LogEntry "Failed to create directory: $dir - $($_.Exception.Message)" -LogType "Error"
            }
        }
    }

    # Clear timeline information when tool is launched
    Clear-TimelineInformation
}

function Clear-TimelineInformation {
    <#
    .SYNOPSIS
        Clears timeline information when the tool is launched so it will have only current values
    #>

    try {
        $installLog = "C:\temp\PackageDevManager_InstallLog.txt"
        if (Test-Path $installLog) {
            Write-LogEntry "Clearing previous timeline information..."
            Remove-Item $installLog -Force -ErrorAction SilentlyContinue
            Write-LogEntry "Timeline information cleared - ready for new operations"
        }

        # Also clear any existing temp scan files to ensure fresh start
        $tempScanFiles = @(
            "$tempPath\BfShortcuts.txt",
            "$tempPath\AfShortcuts.txt",
            "$tempPath\BfprogramFiles.txt",
            "$tempPath\AfprogramFiles.txt",
            "$tempPath\BfUninstallKey.txt",
            "$tempPath\AfUninstallKey.txt",
            "$tempPath\BfRegistryAreas.txt",
            "$tempPath\AfRegistryAreas.txt",
            "$tempPath\BfUHGLogs.txt",
            "$tempPath\AfUHGLogs.txt"
        )

        foreach ($file in $tempScanFiles) {
            if (Test-Path $file) {
                Remove-Item $file -Force -ErrorAction SilentlyContinue
            }
        }

        Write-LogEntry "Previous scan data cleared for fresh start"
    }
    catch {
        Write-LogEntry "Warning: Could not clear all timeline information: $($_.Exception.Message)" -LogType "Warning"
    }
}

function Invoke-BaseScan {
    Write-LogEntry "Starting base scan..."

    try {
        # Initialize directories
        Initialize-WorkingDirectories

        # Scan shortcuts (Public Desktop and Start Menu)
        Write-LogEntry "Scanning shortcuts..."
        $shortcuts = @()
        $shortcuts += Get-ChildItem $shortcutPath -Recurse -Include "*.lnk" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        $shortcuts += Get-ChildItem $shortcutPublicDesktop -Include "*.lnk" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        # Add Start Menu shortcuts
        $startMenuPaths = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
            "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
        )
        foreach ($startMenuPath in $startMenuPaths) {
            if (Test-Path $startMenuPath) {
                $shortcuts += Get-ChildItem $startMenuPath -Recurse -Include "*.lnk" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
            }
        }
        # Remove duplicates and empty entries
        $shortcuts = $shortcuts | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique
        $shortcuts | Out-File "$tempPath\BfShortcuts.txt" -Encoding UTF8

        # Scan file system (Program Files with detailed scanning)
        Write-LogEntry "Scanning file system..."
        $programFiles = @()
        # Scan Program Files directories with detailed scanning (depth 2)
        $programFiles += (Get-ChildItem $program86 -Recurse -Depth 2 -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }).FullName
        $programFiles += (Get-ChildItem $program64 -Recurse -Depth 2 -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }).FullName
        # Scan ProgramData separately (not overlapping with above)
        if (Test-Path "C:\ProgramData") {
            $programFiles += (Get-ChildItem "C:\ProgramData" -Depth 1 -Force -Directory -ErrorAction SilentlyContinue).FullName
        }
        # Remove duplicates and empty entries
        $programFiles = $programFiles | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique
        $programFiles | Out-File "$tempPath\BfprogramFiles.txt" -Encoding UTF8

        # Scan registry (ARP entries in both 32bit and 64bit nodes)
        Write-LogEntry "Scanning registry (Add/Remove Programs)..."
        $uninstallKeys = @()
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        foreach ($path in $uninstallPaths) {
            $uninstallKeys += Get-ItemProperty $path -ErrorAction SilentlyContinue |
                             Where-Object { $_.DisplayName } |
                             Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, PSPath
        }
        $uninstallKeys | ConvertTo-Json | Out-File "$tempPath\BfUninstallKey.txt" -Encoding UTF8

        # Scan major registry areas for changes
        Write-LogEntry "Scanning major registry areas..."
        $registryAreas = @()
        $regPaths = @(
            "HKLM:\SOFTWARE\Classes",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        )

        foreach ($regPath in $regPaths) {
            try {
                Write-LogEntry "Scanning registry path: $regPath"
                if (Test-Path $regPath) {
                    # Skip HKLM:\SOFTWARE\Classes as it's too large and causes hangs
                    if ($regPath -eq "HKLM:\SOFTWARE\Classes") {
                        Write-LogEntry "Skipping HKLM:\SOFTWARE\Classes (too large, causes performance issues)"
                        continue
                    }

                    # Use timeout for registry scanning to prevent hangs
                    $regItems = Get-ChildItem $regPath -ErrorAction SilentlyContinue |
                                Select-Object Name, @{Name="Path"; Expression={$_.PSPath}} |
                                Select-Object -First 1000  # Limit to first 1000 items to prevent memory issues
                    $registryAreas += $regItems
                    Write-LogEntry "Scanned $($regItems.Count) items from $regPath"
                }
                else {
                    Write-LogEntry "Registry path not found: $regPath" -LogType "Warning"
                }
            }
            catch {
                Write-LogEntry "Error scanning registry path: $regPath - $($_.Exception.Message)" -LogType "Warning"
            }
        }
        $registryAreas | ConvertTo-Json | Out-File "$tempPath\BfRegistryAreas.txt" -Encoding UTF8

        # Scan UHG logs directory - simple approach
        Write-LogEntry "Scanning UHG logs directory..."
        $uhgLogsPath = "C:\Windows\UHGLOGS"
        $uhgLogs = @()
        if (Test-Path $uhgLogsPath) {
            $uhgLogs = Get-ChildItem $uhgLogsPath -Recurse -File -ErrorAction SilentlyContinue |
                       Select-Object Name, FullName, LastWriteTime, Length
        }
        Write-LogEntry "Found $($uhgLogs.Count) UHG log files (before scan)"
        $uhgLogs | ConvertTo-Json | Out-File "$tempPath\BfUHGLogs.txt" -Encoding UTF8

        Write-LogEntry "Base scan completed successfully"
    }
    catch {
        Write-LogEntry "Error during base scan: $($_.Exception.Message)" -LogType "Error"
    }
}

function Invoke-SecondScan {
    <#
    .SYNOPSIS
        Performs second scan and generates difference files
    .PARAMETER OperationType
        Type of operation: Install or Uninstall
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$OperationType = "Install"
    )

    Write-LogEntry "Starting second scan for $OperationType operation..."

    try {
        # Scan shortcuts after installation (Public Desktop and Start Menu)
        Write-LogEntry "Scanning shortcuts (post-installation)..."
        $shortcuts = @()
        $shortcuts += Get-ChildItem $shortcutPath -Recurse -Include "*.lnk" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        $shortcuts += Get-ChildItem $shortcutPublicDesktop -Include "*.lnk" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
        # Add Start Menu shortcuts
        $startMenuPaths = @(
            "$env:ProgramData\Microsoft\Windows\Start Menu\Programs",
            "$env:APPDATA\Microsoft\Windows\Start Menu\Programs"
        )
        foreach ($startMenuPath in $startMenuPaths) {
            if (Test-Path $startMenuPath) {
                $shortcuts += Get-ChildItem $startMenuPath -Recurse -Include "*.lnk" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
            }
        }
        # Remove duplicates and empty entries
        $shortcuts = $shortcuts | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique
        $shortcuts | Out-File "$tempPath\AfShortcuts.txt" -Encoding UTF8

        # Scan file system after installation (detailed scanning)
        Write-LogEntry "Scanning file system (post-installation)..."
        $programFiles = @()
        # Scan Program Files directories with detailed scanning (depth 2)
        $programFiles += (Get-ChildItem $program86 -Recurse -Depth 2 -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }).FullName
        $programFiles += (Get-ChildItem $program64 -Recurse -Depth 2 -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }).FullName
        # Scan ProgramData separately (not overlapping with above)
        if (Test-Path "C:\ProgramData") {
            $programFiles += (Get-ChildItem "C:\ProgramData" -Depth 1 -Force -Directory -ErrorAction SilentlyContinue).FullName
        }
        # Remove duplicates and empty entries
        $programFiles = $programFiles | Where-Object { $_ -and $_.Trim() } | Sort-Object -Unique
        $programFiles | Out-File "$tempPath\AfprogramFiles.txt" -Encoding UTF8

        # Scan registry after installation (ARP entries in both 32bit and 64bit nodes)
        Write-LogEntry "Scanning registry (post-installation)..."
        $uninstallKeys = @()
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        )

        foreach ($path in $uninstallPaths) {
            $uninstallKeys += Get-ItemProperty $path -ErrorAction SilentlyContinue |
                             Where-Object { $_.DisplayName } |
                             Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, PSPath
        }
        $uninstallKeys | ConvertTo-Json | Out-File "$tempPath\AfUninstallKey.txt" -Encoding UTF8

        # Scan major registry areas after installation
        Write-LogEntry "Scanning major registry areas (post-installation)..."
        $registryAreas = @()
        $regPaths = @(
            "HKLM:\SOFTWARE\Classes",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce",
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        )

        foreach ($regPath in $regPaths) {
            try {
                Write-LogEntry "Scanning registry path (after): $regPath"
                if (Test-Path $regPath) {
                    # Skip HKLM:\SOFTWARE\Classes as it's too large and causes hangs
                    if ($regPath -eq "HKLM:\SOFTWARE\Classes") {
                        Write-LogEntry "Skipping HKLM:\SOFTWARE\Classes (too large, causes performance issues)"
                        continue
                    }

                    # Use timeout for registry scanning to prevent hangs
                    $regItems = Get-ChildItem $regPath -ErrorAction SilentlyContinue |
                                Select-Object Name, @{Name="Path"; Expression={$_.PSPath}} |
                                Select-Object -First 1000  # Limit to first 1000 items to prevent memory issues
                    $registryAreas += $regItems
                    Write-LogEntry "Scanned $($regItems.Count) items from $regPath (after)"
                }
                else {
                    Write-LogEntry "Registry path not found (after): $regPath" -LogType "Warning"
                }
            }
            catch {
                Write-LogEntry "Error scanning registry path (after): $regPath - $($_.Exception.Message)" -LogType "Warning"
            }
        }
        $registryAreas | ConvertTo-Json | Out-File "$tempPath\AfRegistryAreas.txt" -Encoding UTF8

        # Scan UHG logs directory - simple approach
        Write-LogEntry "Scanning UHG logs directory (post-installation)..."
        $uhgLogsPath = "C:\Windows\UHGLOGS"
        $uhgLogs = @()
        if (Test-Path $uhgLogsPath) {
            $uhgLogs = Get-ChildItem $uhgLogsPath -Recurse -File -ErrorAction SilentlyContinue |
                       Select-Object Name, FullName, LastWriteTime, Length
        }
        Write-LogEntry "Found $($uhgLogs.Count) UHG log files (after scan)"
        $uhgLogs | ConvertTo-Json | Out-File "$tempPath\AfUHGLogs.txt" -Encoding UTF8

        # Generate differences
        Write-LogEntry "Generating difference reports..."

        # Determine comparison direction based on operation type
        $comparisonIndicator = if ($OperationType -eq "Uninstall") { "<=" } else { "=>" }
        Write-LogEntry "Using comparison indicator: $comparisonIndicator for $OperationType operation"

        # Compare shortcuts
        if ((Test-Path "$tempPath\BfShortcuts.txt") -and (Test-Path "$tempPath\AfShortcuts.txt")) {
            $beforeShortcuts = Get-Content "$tempPath\BfShortcuts.txt" -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() }
            $afterShortcuts = Get-Content "$tempPath\AfShortcuts.txt" -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() }
            $diffShortcuts = Compare-Object $beforeShortcuts $afterShortcuts | Where-Object { $_.SideIndicator -eq $comparisonIndicator } | Select-Object -ExpandProperty InputObject | Sort-Object -Unique
            if ($diffShortcuts) {
                $diffShortcuts | Out-File "$tempPath\diffShortcuts.txt" -Encoding UTF8
            }
        }

        # Compare program files
        if ((Test-Path "$tempPath\BfprogramFiles.txt") -and (Test-Path "$tempPath\AfprogramFiles.txt")) {
            $beforePrograms = Get-Content "$tempPath\BfprogramFiles.txt" -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() }
            $afterPrograms = Get-Content "$tempPath\AfprogramFiles.txt" -ErrorAction SilentlyContinue | Where-Object { $_ -and $_.Trim() }
            $diffPrograms = Compare-Object $beforePrograms $afterPrograms | Where-Object { $_.SideIndicator -eq $comparisonIndicator } | Select-Object -ExpandProperty InputObject | Sort-Object -Unique
            if ($diffPrograms) {
                $diffPrograms | Out-File "$tempPath\diffProgramFiles.txt" -Encoding UTF8
            }
        }

        # Compare uninstall keys (JSON data)
        if ((Test-Path "$tempPath\BfUninstallKey.txt") -and (Test-Path "$tempPath\AfUninstallKey.txt")) {
            try {
                $beforeUninstallJson = Get-Content "$tempPath\BfUninstallKey.txt" -Raw -ErrorAction SilentlyContinue
                $afterUninstallJson = Get-Content "$tempPath\AfUninstallKey.txt" -Raw -ErrorAction SilentlyContinue

                $beforeUninstall = @()
                $afterUninstall = @()

                if ($beforeUninstallJson) { $beforeUninstall = $beforeUninstallJson | ConvertFrom-Json }
                if ($afterUninstallJson) { $afterUninstall = $afterUninstallJson | ConvertFrom-Json }

                # Compare by DisplayName to find changed entries
                $beforeNames = $beforeUninstall | ForEach-Object { $_.DisplayName } | Where-Object { $_ }
                $afterNames = $afterUninstall | ForEach-Object { $_.DisplayName } | Where-Object { $_ }
                $changedNames = Compare-Object $beforeNames $afterNames | Where-Object { $_.SideIndicator -eq $comparisonIndicator } | Select-Object -ExpandProperty InputObject | Sort-Object -Unique

                if ($changedNames) {
                    if ($OperationType -eq "Uninstall") {
                        # For uninstall, get entries that were in Before but not in After (removed)
                        $diffUninstall = $beforeUninstall | Where-Object { $changedNames -contains $_.DisplayName }
                    } else {
                        # For install, get entries that are in After but not in Before (added)
                        $diffUninstall = $afterUninstall | Where-Object { $changedNames -contains $_.DisplayName }
                    }
                    $diffUninstall | ConvertTo-Json | Out-File "$tempPath\diffUninstallKey.txt" -Encoding UTF8
                }
            }
            catch {
                Write-LogEntry "Error comparing uninstall keys: $($_.Exception.Message)" -LogType "Warning"
            }
        }

        # Compare UHG logs - simplified approach
        if ((Test-Path "$tempPath\BfUHGLogs.txt") -and (Test-Path "$tempPath\AfUHGLogs.txt")) {
            try {
                Write-LogEntry "Comparing UHG logs for $OperationType operation..."

                $beforeUHGJson = Get-Content "$tempPath\BfUHGLogs.txt" -Raw -ErrorAction SilentlyContinue
                $afterUHGJson = Get-Content "$tempPath\AfUHGLogs.txt" -Raw -ErrorAction SilentlyContinue

                $beforeUHG = @()
                $afterUHG = @()

                if ($beforeUHGJson) { $beforeUHG = $beforeUHGJson | ConvertFrom-Json }
                if ($afterUHGJson) { $afterUHG = $afterUHGJson | ConvertFrom-Json }

                Write-LogEntry "UHG Logs: Before=$($beforeUHG.Count), After=$($afterUHG.Count)"

                # Enhanced comparison: detect new, removed, and modified files
                $beforePaths = $beforeUHG | ForEach-Object { $_.FullName }
                $afterPaths = $afterUHG | ForEach-Object { $_.FullName }

                $diffUHGLogs = @()

                if ($OperationType -eq "Install") {
                    # For install: find new files and modified files

                    # 1. Find completely new files (in After but not in Before)
                    $newPaths = $afterPaths | Where-Object { $beforePaths -notcontains $_ }
                    $newFiles = @()
                    if ($newPaths) {
                        $newFiles = @($afterUHG | Where-Object { $newPaths -contains $_.FullName })
                    }

                    # 2. Find modified files (same path but different LastWriteTime or Length)
                    $modifiedFiles = @()
                    foreach ($afterFile in $afterUHG) {
                        $beforeFile = $beforeUHG | Where-Object { $_.FullName -eq $afterFile.FullName }
                        if ($beforeFile) {
                            # Compare LastWriteTime and Length to detect modifications
                            $afterTime = try { [DateTime]::Parse($afterFile.LastWriteTime) } catch { $null }
                            $beforeTime = try { [DateTime]::Parse($beforeFile.LastWriteTime) } catch { $null }

                            if (($afterTime -and $beforeTime -and $afterTime -gt $beforeTime) -or
                                ($afterFile.Length -ne $beforeFile.Length)) {
                                $modifiedFiles += $afterFile
                            }
                        }
                    }

                    # Combine new and modified files using ArrayList for better compatibility
                    $diffUHGLogs = @()
                    if ($newFiles) { $diffUHGLogs += $newFiles }
                    if ($modifiedFiles) { $diffUHGLogs += $modifiedFiles }
                    Write-LogEntry "Install: Found $($newFiles.Count) new files and $($modifiedFiles.Count) modified files"

                } else {
                    # For uninstall: find removed files, modified files, AND newly created files (like uninstall logs)

                    # 1. Find completely removed files (in Before but not in After)
                    $removedPaths = $beforePaths | Where-Object { $afterPaths -notcontains $_ }
                    $removedFiles = @()
                    if ($removedPaths) {
                        $removedFiles = @($beforeUHG | Where-Object { $removedPaths -contains $_.FullName })
                    }

                    # 2. Find newly created files during uninstall (in After but not in Before)
                    $newPaths = $afterPaths | Where-Object { $beforePaths -notcontains $_ }
                    $newFiles = @()
                    if ($newPaths) {
                        $newFiles = @($afterUHG | Where-Object { $newPaths -contains $_.FullName })
                    }

                    # 3. Find modified files during uninstall (same path but different LastWriteTime or Length)
                    $modifiedFiles = @()
                    foreach ($beforeFile in $beforeUHG) {
                        $afterFile = $afterUHG | Where-Object { $_.FullName -eq $beforeFile.FullName }
                        if ($afterFile) {
                            # Compare LastWriteTime and Length to detect modifications
                            $afterTime = try { [DateTime]::Parse($afterFile.LastWriteTime) } catch { $null }
                            $beforeTime = try { [DateTime]::Parse($beforeFile.LastWriteTime) } catch { $null }

                            if (($afterTime -and $beforeTime -and $afterTime -gt $beforeTime) -or
                                ($afterFile.Length -ne $beforeFile.Length)) {
                                # For uninstall, we want to show the "after" state of modified files (current state)
                                $modifiedFiles += $afterFile
                            }
                        }
                    }

                    # Combine removed, new, and modified files using ArrayList for better compatibility
                    $diffUHGLogs = @()
                    if ($removedFiles) { $diffUHGLogs += $removedFiles }
                    if ($newFiles) { $diffUHGLogs += $newFiles }
                    if ($modifiedFiles) { $diffUHGLogs += $modifiedFiles }
                    Write-LogEntry "Uninstall: Found $($removedFiles.Count) removed files, $($newFiles.Count) new files, and $($modifiedFiles.Count) modified files"
                }

                Write-LogEntry "Found $($diffUHGLogs.Count) UHG log differences"

                # Generate diff files
                if ($diffUHGLogs -and $diffUHGLogs.Count -gt 0) {
                    $diffUHGLogs | ConvertTo-Json | Out-File "$tempPath\diffUHGLogs.txt" -Encoding UTF8

                    # Generate difflogFiles.txt with file paths only
                    $logFiles = $diffUHGLogs | ForEach-Object { $_.FullName }
                    if ($logFiles) {
                        $logFiles | Out-File "$tempPath\difflogFiles.txt" -Encoding UTF8
                        Write-LogEntry "Generated difflogFiles.txt with $($logFiles.Count) log files"
                    }

                    Write-LogEntry "Generated diffUHGLogs.txt with $($diffUHGLogs.Count) entries"
                } else {
                    Write-LogEntry "No UHG log differences found"
                }
            }
            catch {
                Write-LogEntry "Error comparing UHG logs: $($_.Exception.Message)" -LogType "Warning"
            }
        }

        # Compare registry areas (JSON data)
        if ((Test-Path "$tempPath\BfRegistryAreas.txt") -and (Test-Path "$tempPath\AfRegistryAreas.txt")) {
            try {
                $beforeRegJson = Get-Content "$tempPath\BfRegistryAreas.txt" -Raw -ErrorAction SilentlyContinue
                $afterRegJson = Get-Content "$tempPath\AfRegistryAreas.txt" -Raw -ErrorAction SilentlyContinue

                $beforeReg = @()
                $afterReg = @()

                if ($beforeRegJson) { $beforeReg = $beforeRegJson | ConvertFrom-Json }
                if ($afterRegJson) { $afterReg = $afterRegJson | ConvertFrom-Json }

                # Compare by Path and Name to find changed entries
                if ($OperationType -eq "Uninstall") {
                    # For uninstall, find entries that were in Before but not in After (removed)
                    $changedRegEntries = $beforeReg | Where-Object {
                        $beforeEntry = $_
                        -not ($afterReg | Where-Object { $_.Path -eq $beforeEntry.Path -and $_.Name -eq $beforeEntry.Name })
                    } | Sort-Object -Property Path, Name -Unique
                } else {
                    # For install, find entries that are in After but not in Before (added)
                    $changedRegEntries = $afterReg | Where-Object {
                        $afterEntry = $_
                        -not ($beforeReg | Where-Object { $_.Path -eq $afterEntry.Path -and $_.Name -eq $afterEntry.Name })
                    } | Sort-Object -Property Path, Name -Unique
                }

                if ($changedRegEntries) {
                    $changedRegEntries | ConvertTo-Json | Out-File "$tempPath\diffRegistryAreas.txt" -Encoding UTF8
                }
            }
            catch {
                Write-LogEntry "Error comparing registry areas: $($_.Exception.Message)" -LogType "Warning"
            }
        }

        Write-LogEntry "Second scan completed successfully"
    }
    catch {
        Write-LogEntry "Error during second scan: $($_.Exception.Message)" -LogType "Error"
    }
}

function Invoke-PackageDirectory {
    <#
    .SYNOPSIS
        Processes package directory and loads available packages
    .PARAMETER archiveLocation
        Path to the package archive location
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$archiveLocation
    )

    try {
        if (Test-Path $archiveLocation) {
            $packages = Get-ChildItem $archiveLocation -Directory -ErrorAction SilentlyContinue

            if ($packages.Count -gt 0) {
                Write-LogEntry "Found $($packages.Count) packages in $archiveLocation"

                # Process each package
                foreach ($package in $packages) {
                    Write-LogEntry "Processing package: $($package.Name)"

                    # Check for different file types
                    $msiFiles = Get-ChildItem $package.FullName -Filter "*.msi" -ErrorAction SilentlyContinue
                    $mstFiles = Get-ChildItem $package.FullName -Filter "*.mst" -ErrorAction SilentlyContinue
                    $exeFiles = Get-ChildItem $package.FullName -Filter "*.exe" -ErrorAction SilentlyContinue
                    $ps1Files = Get-ChildItem $package.FullName -Filter "*.ps1" -ErrorAction SilentlyContinue

                    # Update global flags
                    $global:msiExist = $msiFiles.Count -gt 0
                    $global:mstExist = $mstFiles.Count -gt 0
                    $global:exeExist = $exeFiles.Count -gt 0
                    $global:ps1Exist = $ps1Files.Count -gt 0

                    # Store file paths
                    if ($global:msiExist) { $global:filemsi = $msiFiles[0].FullName }
                    if ($global:mstExist) { $global:filemst = $mstFiles[0].FullName }
                    if ($global:exeExist) { $global:fileexe = $exeFiles[0].FullName }
                    if ($global:ps1Exist) { $global:fileps1 = $ps1Files[0].FullName }
                }
            }
            else {
                Write-LogEntry "No packages found in $archiveLocation" -LogType "Warning"
            }
        }
        else {
            Write-LogEntry "Archive location not accessible: $archiveLocation" -LogType "Error"
        }
    }
    catch {
        Write-LogEntry "Error processing package directory: $($_.Exception.Message)" -LogType "Error"
    }
}

function Invoke-PackageValidation {
    <#
    .SYNOPSIS
        Performs comprehensive package validation
    #>

    Write-LogEntry "Starting package validation process"

    try {
        # Initialize directories
        Initialize-WorkingDirectories

        # Perform base scan
        Invoke-BaseScan

        Write-LogEntry "Package validation completed successfully"
    }
    catch {
        Write-LogEntry "Error during package validation: $($_.Exception.Message)" -LogType "Error"
    }
}

function Get-PackageType {
    <#
    .SYNOPSIS
        Determines the package type based on available files
    #>

    if ($global:msiExist -and $global:mstExist) {
        return "MSI+MST"
    }
    elseif ($global:msiExist) {
        return "MSI only"
    }
    elseif ($global:exeExist -and $global:ps1Exist) {
        return "PSADTK Wrapper"
    }
    elseif ($global:exeExist) {
        return "EXE"
    }
    else {
        return "Not Detected"
    }
}

function Invoke-VendorScan {
    <#
    .SYNOPSIS
        Quickly scans for vendor folder names (like original code)
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$archiveLocation
    )

    try {
        if (Test-Path $archiveLocation) {
            Write-Host "Getting vendor folders..." -ForegroundColor Gray

            # Quick scan - just get directory names (no deep scanning)
            $vendorFolders = Get-ChildItem $archiveLocation -Directory -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name

            if ($vendorFolders.Count -gt 0) {
                Write-Host "Found $($vendorFolders.Count) vendors: $($vendorFolders -join ', ')" -ForegroundColor Green

                # Update vendor dropdown with actual folder names
                Update-VendorDropdown -vendors $vendorFolders
            }
            else {
                Write-Host "No vendor folders found in $archiveLocation" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host "Archive location not accessible: $archiveLocation" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "Error scanning vendors: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Update-VendorDropdown {
    <#
    .SYNOPSIS
        Updates the vendor dropdown with scanned vendor names
    #>
    param(
        [Parameter(Mandatory = $true)]
        [array]$vendors
    )

    try {
        # Find the vendor dropdown control
        if ($global:vendorDropdown) {
            $global:vendorDropdown.Items.Clear()
            $global:vendorDropdown.Items.AddRange($vendors)
            Write-Host "Updated vendor dropdown with $($vendors.Count) vendors" -ForegroundColor Green
        }
        else {
            Write-Host "Vendor dropdown control not found" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error updating vendor dropdown: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Test-Shortcuts {
    <#
    .SYNOPSIS
        Tests and validates shortcuts from the template INI file
    #>

    $templatePath = "$env:SystemDrive\temp\AutoTestScan\template.ini"

    if (-not (Test-Path $templatePath)) {
        Write-LogEntry "Template INI file not found: $templatePath" -LogType "Warning"
        return
    }

    try {
        $fileContent = Get-IniContent $templatePath
        $shortcutCount = $fileContent["Shortcut"].Count

        if ($shortcutCount -gt 0) {
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br><br>"
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<p><b>Shortcut Validation Results</b></p>"
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br>"

            foreach ($key in $fileContent["Shortcut"].Keys) {
                $shortcutPath = $fileContent["Shortcut"][$key]

                if (Test-Path $shortcutPath) {
                    $status = "Valid"
                    $statusColor = "green"
                }
                else {
                    $status = "Invalid"
                    $statusColor = "red"
                }

                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<span style='color:$statusColor'>$shortcutPath - $status</span>"
                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br>"

                $csvLine = "`"Shortcut Validation`",`"$shortcutPath`",`"$status`""
                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.csv" -Value $csvLine
            }

            Write-LogEntry "Shortcut validation completed - $shortcutCount shortcuts tested"
        }
        else {
            Write-LogEntry "No shortcuts found for validation"
        }
    }
    catch {
        Write-LogEntry "Error during shortcut validation: $($_.Exception.Message)" -LogType "Error"
    }
}

function Test-ProgramFiles {
    <#
    .SYNOPSIS
        Tests and validates program file installations
    #>

    $templatePath = "$env:SystemDrive\temp\AutoTestScan\template.ini"

    if (-not (Test-Path $templatePath)) {
        Write-LogEntry "Template INI file not found: $templatePath" -LogType "Warning"
        return
    }

    try {
        $fileContent = Get-IniContent $templatePath
        $folderCount = $fileContent["Folder"].Count

        if ($folderCount -gt 0) {
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br><br>"
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<p><b>Program Files Validation Results</b></p>"
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br>"

            foreach ($key in $fileContent["Folder"].Keys) {
                $folderPath = $fileContent["Folder"][$key]

                if (Test-Path $folderPath) {
                    $status = "Valid"
                    $statusColor = "green"
                }
                else {
                    $status = "Invalid"
                    $statusColor = "red"
                }

                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<span style='color:$statusColor'>$folderPath - $status</span>"
                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br>"

                $csvLine = "`"Program Files Validation`",`"$folderPath`",`"$status`""
                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.csv" -Value $csvLine
            }

            Write-LogEntry "Program files validation completed"
        }
        else {
            Write-LogEntry "No program files found for validation"
        }
    }
    catch {
        Write-LogEntry "Error during program files validation: $($_.Exception.Message)" -LogType "Error"
    }
}

function Test-UninstallKeys {
    <#
    .SYNOPSIS
        Tests and validates uninstall registry keys
    #>

    $templatePath = "$env:SystemDrive\temp\AutoTestScan\template.ini"

    if (-not (Test-Path $templatePath)) {
        Write-LogEntry "Template INI file not found: $templatePath" -LogType "Warning"
        return
    }

    try {
        $fileContent = Get-IniContent $templatePath
        $uninstallKeyCount = $fileContent["UninstallKey"].Count

        if ($uninstallKeyCount -gt 0) {
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br><br>"
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<p><b>Uninstall Keys Validation Results</b></p>"
            Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br>"

            # Registry paths to check
            $uninstallkey32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
            $uninstallkey64 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            $uninstallUserKey = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"

            foreach ($key in $fileContent["UninstallKey"].Keys) {
                $registryKey = $fileContent["UninstallKey"][$key]

                # Check in all three registry locations
                $keyExists = $false
                $registryPaths = @($uninstallkey32, $uninstallkey64, $uninstallUserKey)

                foreach ($regPath in $registryPaths) {
                    $fullPath = Join-Path $regPath (Split-Path $registryKey -Leaf)
                    if (Test-Path $fullPath) {
                        $keyExists = $true
                        break
                    }
                }

                if ($keyExists) {
                    $status = "Valid"
                    $statusColor = "green"
                }
                else {
                    $status = "Invalid"
                    $statusColor = "red"
                }

                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<span style='color:$statusColor'>$registryKey - $status</span>"
                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.html" -Value "<br>"

                $csvLine = "`"Uninstall Keys Validation`",`"$registryKey`",`"$status`""
                Add-Content -Path "$env:SystemDrive\temp\AutoTest\Tool_Output\report.csv" -Value $csvLine
            }

            Write-LogEntry "Uninstall keys validation completed"
        }
        else {
            Write-LogEntry "No uninstall keys found for validation"
        }
    }
    catch {
        Write-LogEntry "Error during uninstall keys validation: $($_.Exception.Message)" -LogType "Error"
    }
}

function Get-IniContent {
    <#
    .SYNOPSIS
        Gets the content of an INI file
    .DESCRIPTION
        Gets the content of an INI file and returns it as a hashtable
    .PARAMETER FilePath
        Path to the INI file to read
    #>

    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [string]$FilePath
    )

    try {
        $ini = @{}
        $section = $null
        $CommentCount = 0

        if (-not (Test-Path $FilePath)) {
            Write-LogEntry "INI file not found: $FilePath" -LogType "Warning"
            return $ini
        }

        switch -regex -file $FilePath {
            "^\[(.+)\]$" {
                # Section
                $section = $matches[1]
                $ini[$section] = @{}
                $CommentCount = 0
            }
            "^(;.*)$" {
                # Comment
                if (!($section)) {
                    $section = "No-Section"
                    $ini[$section] = @{}
                }
                $value = $matches[1]
                $CommentCount = $CommentCount + 1
                $name = "Comment" + $CommentCount
                $ini[$section][$name] = $value
            }
            "(.+?)\s*=\s*(.*)" {
                # Key
                if (!($section)) {
                    $section = "No-Section"
                    $ini[$section] = @{}
                }
                $name, $value = $matches[1..2]
                $ini[$section][$name] = $value
            }
        }

        return $ini
    }
    catch {
        Write-LogEntry "Error reading INI file: $($_.Exception.Message)" -LogType "Error"
        return @{}
    }
}

function Out-IniFile {
    <#
    .SYNOPSIS
        Writes hashtable content to INI file format
    .PARAMETER InputObject
        Hashtable containing the data to write
    .PARAMETER FilePath
        Path to the INI file
    .PARAMETER Append
        Whether to append to existing file
    #>

    [CmdletBinding()]
    param(
        [switch]$Append,
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [Hashtable]$InputObject
    )

    try {
        if ($Append) {
            $outFile = Get-Item $FilePath -ErrorAction SilentlyContinue
        }
        else {
            $outFile = New-Item -ItemType File -Path $FilePath -Force -ErrorAction SilentlyContinue
        }

        if (!($outFile)) {
            Write-LogEntry "Could not create file: $FilePath" -LogType "Error"
            return
        }

        foreach ($section in $InputObject.Keys) {
            if (!($($InputObject[$section].GetType().Name) -eq "Hashtable")) {
                Add-Content -Path $outFile -Value "$section=$($InputObject[$section])"
            }
            else {
                # Write section header
                Add-Content -Path $outFile -Value "[$section]"

                # Write key-value pairs
                foreach ($key in ($InputObject[$section].Keys | Sort-Object)) {
                    if ($key -match "^Comment[\d]+") {
                        Add-Content -Path $outFile -Value "$($InputObject[$section][$key])"
                    }
                    else {
                        Add-Content -Path $outFile -Value "$key=$($InputObject[$section][$key])"
                    }
                }
                Add-Content -Path $outFile -Value ""
            }
        }
    }
    catch {
        Write-LogEntry "Error writing INI file: $($_.Exception.Message)" -LogType "Error"
    }
}

#endregion

# Launch the GUI if script is run directly
if ($MyInvocation.InvocationName -ne '.') {
    Show-PackageDevManagerGUI
}

#endregion
