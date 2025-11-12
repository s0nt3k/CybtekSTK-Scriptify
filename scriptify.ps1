#requires -Version 5.1

<#
.SYNOPSIS
    Launches the Cybtek STK Office Deployment Tool WPF interface.
.DESCRIPTION
    Initializes assemblies, ensures supporting directories exist, loads defaults, and wires all UI controls
    and command handlers before showing the main window. The window remains open until it is closed from the UI.
.EXAMPLE
    Show-CybtekODTGui
#>
function Show-CybtekODTGui {

    Set-StrictMode -Version 2.0

    # PowerShell 5.1-friendly assembly loads
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, System.Windows.Forms

    # WPF requires STA in Windows PowerShell 5.1
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
        $self = $PSCommandPath; if (-not $self) { $self = $MyInvocation.MyCommand.Path }
        if ($self) {
            Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
                -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-STA','-File',"`"$self`"")
            return
        } else {
            throw 'This GUI requires STA. Start Windows PowerShell with -STA and run the script again.'
        }
    }

    $script:AppRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Convert-Path . }
    $script:DefaultsPath = Join-Path $script:AppRoot 'odt-defaults.json'
    $script:LogsPath = Join-Path $script:AppRoot 'logs'
    $script:GeneratedPath = Join-Path $script:AppRoot 'generated'
    $script:AppDataMicrosoft = Join-Path ([Environment]::GetFolderPath('ApplicationData')) 'Microsoft'
    $script:OfficeInstallerToolUri = 'https://aka.ms/odt'
    $script:OfficeInstallerRoot = Join-Path $script:AppDataMicrosoft 'ODT'
    $script:OdtSetupExePath = Join-Path $script:OfficeInstallerRoot 'setup.exe'
    $script:OfficeInstallerToolUrl = 'https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19029-20278.exe'
    $script:DownloadPath = [Environment]::GetFolderPath('Desktop')
    $script:LogHost = $env:COMPUTERNAME
    $script:ScriptFile = $PSCommandPath
    if (-not $script:ScriptFile) {
        $script:ScriptFile = $MyInvocation.MyCommand.Path
    }

    # <LICENSE_DATA_BLOCK>
    $script:EmbeddedLicenseData = @'
{
  "Technician": "",
  "Company": "",
  "Address": "",
  "Phone": "",
  "Email": "",
  "Website": "",
  "License": "",
  "Saved": false
}
'@
    # </LICENSE_DATA_BLOCK>

    function New-DirectoryIfMissing {
        param(
            [Parameter(Mandatory)][string]$Path
        )
        if (-not (Test-Path -LiteralPath $Path)) {
            [void](New-Item -Path $Path -ItemType Directory -Force)
        }
    }

    New-DirectoryIfMissing -Path $script:LogsPath
    New-DirectoryIfMissing -Path $script:GeneratedPath
    New-DirectoryIfMissing -Path $script:AppDataMicrosoft
    New-DirectoryIfMissing -Path $script:OfficeInstallerRoot

    function Install-OfficeDeploymentTool {
        param(
            [string]$Destination = $script:OfficeInstallerRoot,
            [string]$SourceUrl = $script:OfficeInstallerToolUrl,
            [switch]$Force,
            [switch]$KeepInstaller
        )

        $setupPath = Join-Path $Destination 'setup.exe'
        if (-not $Force -and (Test-Path -LiteralPath $setupPath)) {
            return
        }

        if (-not (Test-Path -LiteralPath $Destination)) {
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
        }

        $tempRoot = Join-Path ([IO.Path]::GetTempPath()) ("ODT_" + [guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
        $installer = Join-Path $tempRoot 'officedeploymenttool.exe'

        $oldProgress = $global:ProgressPreference
        $global:ProgressPreference = 'SilentlyContinue'
        try {
            Show-Info 'Downloading Microsoft Office Deployment Tool for configuration operations.'
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $SourceUrl -OutFile $installer -UseBasicParsing
        } finally {
            $global:ProgressPreference = $oldProgress
        }

        if (-not (Test-Path -LiteralPath $installer)) {
            throw "Download failed from $SourceUrl"
        }

        try {
    # Use variables (not $args) and reference them in Start-Process.
    $extractArgs  = "/quiet /extract:`"$Destination`""
    $fallbackArgs = "/extract:`"$Destination`""

    $proc = Start-Process -FilePath $installer -ArgumentList $extractArgs -Wait -PassThru -WindowStyle Hidden
    if ($proc -and $proc.ExitCode -ne 0) {
        $proc = Start-Process -FilePath $installer -ArgumentList $fallbackArgs -Wait -PassThru -WindowStyle Hidden
    }

    if (-not (Test-Path -LiteralPath $setupPath)) {
        throw "Extraction incomplete: 'setup.exe' not found in $Destination"
    }

    Show-Info 'Office Deployment Tool extracted to AppData.'
} catch {
    Show-Error "Failed to download or extract the Office Deployment Tool: $_"
    throw
} finally {
    if (-not $KeepInstaller) {
        Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
    }

    function Initialize-OdtSetup {
        Install-OfficeDeploymentTool
    }

    $officeCatalog = @(
        [pscustomobject]@{
            Name = 'Select Product'
            Editions = @()
            License = ''
            ProductId = ''
        }
        [pscustomobject]@{
            Name = 'MS Office 2016'
            Editions = @(
                [pscustomobject]@{ Name = 'Standard (Retail)'; License = 'Retail'; ProductId = 'Standard2016Retail' }
                [pscustomobject]@{ Name = 'Standard (Volume)'; License = 'Volume'; ProductId = 'Standard2016Volume' }
                [pscustomobject]@{ Name = 'Professional (Retail)'; License = 'Retail'; ProductId = 'Professional2016Retail' }
                [pscustomobject]@{ Name = 'Professional (Volume)'; License = 'Volume'; ProductId = 'Professional2016Volume' }
                [pscustomobject]@{ Name = 'Professional Plus (Retail)'; License = 'Retail'; ProductId = 'ProPlus2016Retail' }
                [pscustomobject]@{ Name = 'Professional Plus (Volume)'; License = 'Volume'; ProductId = 'ProPlus2016Volume' }
                [pscustomobject]@{ Name = 'Home and Office (Retail)'; License = 'Retail'; ProductId = 'HomeBusiness2016Retail' }
                [pscustomobject]@{ Name = 'Home and Office (Volume)'; License = 'Volume'; ProductId = 'HomeBusiness2016Volume' }
            )
        }
        [pscustomobject]@{
            Name = 'MS Office 2019'
            Editions = @(
                [pscustomobject]@{ Name = 'Standard (Retail)'; License = 'Retail'; ProductId = 'Standard2019Retail' }
                [pscustomobject]@{ Name = 'Standard (Volume)'; License = 'Volume'; ProductId = 'Standard2019Volume' }
                [pscustomobject]@{ Name = 'Professional (Retail)'; License = 'Retail'; ProductId = 'Professional2019Retail' }
                [pscustomobject]@{ Name = 'Professional (Volume)'; License = 'Volume'; ProductId = 'Professional2019Volume' }
                [pscustomobject]@{ Name = 'Professional Plus (Retail)'; License = 'Retail'; ProductId = 'ProPlus2019Retail' }
                [pscustomobject]@{ Name = 'Professional Plus (Volume)'; License = 'Volume'; ProductId = 'ProPlus2019Volume' }
                [pscustomobject]@{ Name = 'Home and Office (Retail)'; License = 'Retail'; ProductId = 'HomeBusiness2019Retail' }
                [pscustomobject]@{ Name = 'Home and Office (Volume)'; License = 'Volume'; ProductId = 'HomeBusiness2019Volume' }
            )
        }
        [pscustomobject]@{
            Name = 'MS Office 2021'
            Editions = @(
                [pscustomobject]@{ Name = 'Standard (Retail)'; License = 'Retail'; ProductId = 'Standard2021Retail' }
                [pscustomobject]@{ Name = 'Standard (Volume)'; License = 'Volume'; ProductId = 'Standard2021Volume' }
                [pscustomobject]@{ Name = 'Professional (Retail)'; License = 'Retail'; ProductId = 'Professional2021Retail' }
                [pscustomobject]@{ Name = 'Professional (Volume)'; License = 'Volume'; ProductId = 'Professional2021Volume' }
                [pscustomobject]@{ Name = 'Professional Plus (Retail)'; License = 'Retail'; ProductId = 'ProPlus2021Retail' }
                [pscustomobject]@{ Name = 'Professional Plus (Volume)'; License = 'Volume'; ProductId = 'ProPlus2021Volume' }
                [pscustomobject]@{ Name = 'Home and Office (Retail)'; License = 'Retail'; ProductId = 'HomeBusiness2021Retail' }
                [pscustomobject]@{ Name = 'Home and Office (Volume)'; License = 'Volume'; ProductId = 'HomeBusiness2021Volume' }
            )
        }
        [pscustomobject]@{
            Name = 'MS Office 2024'
            Editions = @(
                [pscustomobject]@{ Name = 'Standard (Retail)'; License = 'Retail'; ProductId = 'Standard2024Retail' }
                [pscustomobject]@{ Name = 'Standard (Volume)'; License = 'Volume'; ProductId = 'Standard2024Volume' }
                [pscustomobject]@{ Name = 'Professional (Retail)'; License = 'Retail'; ProductId = 'Professional2024Retail' }
                [pscustomobject]@{ Name = 'Professional (Volume)'; License = 'Volume'; ProductId = 'Professional2024Volume' }
                [pscustomobject]@{ Name = 'Professional Plus (Retail)'; License = 'Retail'; ProductId = 'ProPlus2024Retail' }
                [pscustomobject]@{ Name = 'Professional Plus (Volume)'; License = 'Volume'; ProductId = 'ProPlus2024Volume' }
                [pscustomobject]@{ Name = 'Home and Office (Retail)'; License = 'Retail'; ProductId = 'HomeBusiness2024Retail' }
                [pscustomobject]@{ Name = 'Home and Office (Volume)'; License = 'Volume'; ProductId = 'HomeBusiness2024Volume' }
            )
        }
        [pscustomobject]@{
            Name = 'MS Office 365'
            Editions = @(
                [pscustomobject]@{ Name = 'Business (Retail)'; License = 'Retail'; ProductId = 'O365BusinessRetail' }
                [pscustomobject]@{ Name = 'Business Premium (Retail)'; License = 'Retail'; ProductId = 'O365BusinessRetail' }
                [pscustomobject]@{ Name = 'ProPlus (Retail)'; License = 'Retail'; ProductId = 'O365ProPlusRetail' }
                [pscustomobject]@{ Name = 'Enterprise (Retail)'; License = 'Retail'; ProductId = 'O365ProPlusRetail' }
            )
        }
    )

    $channelOptions = @(
        [pscustomobject]@{ Label = 'Current Channel'; Value = 'Current' }
        [pscustomobject]@{ Label = 'Monthly Enterprise'; Value = 'MonthlyEnterprise' }
        [pscustomobject]@{ Label = 'Semi-Annual Enterprise'; Value = 'SemiAnnual' }
        [pscustomobject]@{ Label = 'Semi-Annual Preview'; Value = 'SemiAnnualPreview' }
        [pscustomobject]@{ Label = 'Beta Channel'; Value = 'BetaChannel' }
        [pscustomobject]@{ Label = 'Current Channel Preview'; Value = 'CurrentPreview' }
        [pscustomobject]@{ Label = 'Perpetual VL 2021 (Volume)'; Value = 'PerpetualVL2021' }
        [pscustomobject]@{ Label = 'Perpetual VL 2019 (Volume)'; Value = 'PerpetualVL2019' }
    )

    $versionModes = @(
        [pscustomobject]@{ Label = 'Latest Available'; Value = 'Latest' }
        [pscustomobject]@{ Label = 'Specific Version'; Value = 'Specific' }
    )

    $languageCatalog = @(
        @{ Name = 'English (United States)'; Id = 'en-us' }
        @{ Name = 'English (United Kingdom)'; Id = 'en-gb' }
        @{ Name = 'Spanish (Spain)'; Id = 'es-es' }
        @{ Name = 'French (France)'; Id = 'fr-fr' }
        @{ Name = 'German (Germany)'; Id = 'de-de' }
        @{ Name = 'Italian (Italy)'; Id = 'it-it' }
        @{ Name = 'Portuguese (Brazil)'; Id = 'pt-br' }
        @{ Name = 'Japanese'; Id = 'ja-jp' }
        @{ Name = 'Chinese (Simplified)'; Id = 'zh-cn' }
        @{ Name = 'Chinese (Traditional)'; Id = 'zh-tw' }
        @{ Name = 'Korean'; Id = 'ko-kr' }
        @{ Name = 'Russian'; Id = 'ru-ru' }
    )

    $script:SelectedLanguageIds = @('en-us')
    $script:Is32BitArchitecture = $false
    $script:IsDisplayNone = $false
    $script:ShowVolumeEditions = $false
    $script:DisableRestorePoint = $false
    $script:DisableRemoveOffice = $false
    $script:ForceAppClose = $true
    $script:InstallProject = $false
    $script:InstallVisio = $false
    $script:DefaultsSaved = $false
    function New-AppOptions {
        return @(
            [pscustomobject]@{ Name = 'Excel'; Column = 0; Default = $true }
            [pscustomobject]@{ Name = 'OneDrive Desktop'; Column = 0; Default = $true }
            [pscustomobject]@{ Name = 'Outlook (classic)'; Column = 0; Default = $true }
            [pscustomobject]@{ Name = 'Publisher'; Column = 0; Default = $true }
            [pscustomobject]@{ Name = 'OneDrive (Groove)'; Column = 1; Default = $false }
            [pscustomobject]@{ Name = 'OneNote'; Column = 1; Default = $true }
            [pscustomobject]@{ Name = 'PowerPoint'; Column = 1; Default = $true }
            [pscustomobject]@{ Name = 'Word'; Column = 1; Default = $true }
        )
    }

    $script:AppSelectionOptions = New-AppOptions
    $script:AppSelections = @{}
    $script:AppSelections = @{};
    $script:AppSelectionOptions | ForEach-Object { $script:AppSelections[$_.Name] = $_.Default }
    function Show-Info {
        param([string]$Message,[string]$Title = 'Cybtek STK')
        [System.Windows.MessageBox]::Show($Message,$Title,[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) | Out-Null
    }

    function Show-Error {
        param([string]$Message,[string]$Title = 'Cybtek STK')
        [System.Windows.MessageBox]::Show($Message,$Title,[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error) | Out-Null
    }

    Initialize-OdtSetup

    function New-LicenseDefaults {
        return [pscustomobject]@{
            Technician = ''
            Company = ''
            Address = ''
            Phone = ''
            Email = ''
            Website = ''
            License = ''
            Saved = $false
        }
    }

    function Get-LicenseData {
        try {
            $data = $script:EmbeddedLicenseData | ConvertFrom-Json
            if (-not $data) {
                throw 'Embedded license data is empty.'
            }
            if (-not $data.PSObject.Properties['Saved']) {
                $data | Add-Member -MemberType NoteProperty -Name Saved -Value $false
            }
            return $data
        } catch {
            Write-Warning "Unable to parse embedded license data: $_"
            return New-LicenseDefaults
        }
    }

    function Update-EmbeddedLicenseData {
        param([Parameter(Mandatory)][string]$Json)

        if (-not $script:ScriptFile) {
            Show-Error 'Cannot persist license data because the script path is unknown.'
            return $false
        }
        try {
            $scriptText = Get-Content -LiteralPath $script:ScriptFile -Raw
        } catch {
            Show-Error "Unable to read the script file to persist license data: $_"
            return $false
        }
        $block = @"
# <LICENSE_DATA_BLOCK>
`$script:EmbeddedLicenseData = @'
$Json
'@
# </LICENSE_DATA_BLOCK>
"@
        # PS 5.1-safe Regex construction
        $regex = New-Object System.Text.RegularExpressions.Regex '(?s)# <LICENSE_DATA_BLOCK>.*?# </LICENSE_DATA_BLOCK>'
        if (-not $regex.IsMatch($scriptText)) {
            Show-Error 'Embedded license data block was not found; persistence failed.'
            return $false
        }
        $updatedScript = $regex.Replace($scriptText, $block, 1)
        try {
            Set-Content -LiteralPath $script:ScriptFile -Value $updatedScript -Encoding UTF8
            $script:EmbeddedLicenseData = $Json
        } catch {
            Show-Error "Failed to persist license data: $_"
            return $false
        }
        return $true
    }

    function Save-LicenseData {
        param([pscustomobject]$Data)
        if (-not $Data) { return }
        if (-not $Data.PSObject.Properties['Saved']) {
            $Data | Add-Member -MemberType NoteProperty -Name Saved -Value $false
        }
        $Data | Add-Member -MemberType NoteProperty -Name Saved -Value $true -Force
        $json = $Data | ConvertTo-Json -Depth 3
        Update-EmbeddedLicenseData -Json $json | Out-Null
    }

    function Clear-LicenseData {
        $defaults = New-LicenseDefaults
        $json = $defaults | ConvertTo-Json -Depth 3
        Update-EmbeddedLicenseData -Json $json | Out-Null
        return $defaults
    }

function Write-LogEntry {
    param(
        [Parameter(Mandatory)][ValidateSet('download','install')]$Type,
        [Parameter(Mandatory)][string]$Message,
        [pscustomobject]$Metadata
    )
    $logFile = Join-Path $script:LogsPath "$Type.json"
    $entries = @()
    if (Test-Path -LiteralPath $logFile) {
        try {
            $raw = Get-Content -LiteralPath $logFile -Raw
            if ($raw) {
                $parsed = $raw | ConvertFrom-Json
                if ($parsed) { $entries += $parsed }
            }
        } catch {
            Write-Warning "Failed to read ${logFile}: $_"
        }
    }

    $logHost = if ($script:LogHost) { $script:LogHost } else { $env:COMPUTERNAME }
    $metaOut = if ($Metadata) { $Metadata } else { @{} }

    $entry = [pscustomobject]@{
        Timestamp = (Get-Date).ToString('u')
        Message   = $Message
        Host      = $logHost
        Metadata  = $metaOut
    }

    $entries += $entry
    $json = $entries | ConvertTo-Json -Depth 5
    Set-Content -LiteralPath $logFile -Value $json -Encoding UTF8
}


    function Save-LogEntryAccount {
        param(
            [Parameter(Mandatory)][string]$LogFile,
            [Parameter(Mandatory)][pscustomobject]$Entry,
            [Parameter(Mandatory)][string]$Account
        )

        if (-not (Test-Path -LiteralPath $LogFile)) {
            return $false
        }
        try {
            $raw = Get-Content -LiteralPath $LogFile -Raw
            if (-not $raw) { return $false }
            $entries = $raw | ConvertFrom-Json
        } catch {
            Write-Warning "Unable to read log file for account save: $_"
            return $false
        }

        if (-not $entries) { return $false }
        $match = $entries | Where-Object {
            $_.Timestamp -eq $Entry.Timestamp -and $_.Message -eq $Entry.Message -and $_.Host -eq $Entry.Host
        } | Select-Object -First 1
        if (-not $match) { return $false }

        if (-not $match.Metadata) {
            $match.Metadata = @{}
        }
        $match.Metadata.Account = $Account

        try {
            $entries | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $LogFile -Encoding UTF8
            $Entry.Metadata = $match.Metadata
            return $true
        } catch {
            Write-Warning "Unable to persist account info to log: $_"
            return $false
        }
    }

    function Show-LicenseWindow {
        param($Owner)

        $license = Get-LicenseData
        if ($license.License -eq 'RESETLIC') {
            $license = Clear-LicenseData
        }

        [xml]$licenseXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="License Info" Height="320" Width="335" Background="#0F172A" Foreground="White"
        FontFamily="Segoe UI" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>

        <TextBlock Grid.Row="0" Grid.Column="0" Text="Technician:" Margin="0,0,10,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="0" Grid.Column="1" Name="TechnicianText" Margin="0,0,0,6" />

        <TextBlock Grid.Row="1" Grid.Column="0" Text="Company:" Margin="0,0,10,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="1" Grid.Column="1" Name="CompanyText" Margin="0,0,0,6" />

        <TextBlock Grid.Row="2" Grid.Column="0" Text="Address:" Margin="0,0,10,6" VerticalAlignment="Top" />
        <TextBox Grid.Row="2" Grid.Column="1" Name="AddressText" Height="38" TextWrapping="Wrap"
                 AcceptsReturn="True" VerticalScrollBarVisibility="Auto" Margin="0,0,0,6" Background="White" Foreground="#0F172A" />

        <TextBlock Grid.Row="3" Grid.Column="0" Text="Phone:" Margin="0,0,10,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="3" Grid.Column="1" Name="PhoneText" Margin="0,0,0,6" />

        <TextBlock Grid.Row="4" Grid.Column="0" Text="E-Mail:" Margin="0,0,10,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="4" Grid.Column="1" Name="EmailText" Margin="0,0,0,6" />

        <TextBlock Grid.Row="5" Grid.Column="0" Text="Website:" Margin="0,0,10,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="5" Grid.Column="1" Name="WebsiteText" Margin="0,0,0,6" />

        <Separator Grid.Row="6" Grid.ColumnSpan="2" Margin="0,6" />

        <TextBlock Grid.Row="7" Grid.Column="0" Text="LICENSE:" Margin="0,0,10,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="7" Grid.Column="1" Name="LicenseText" IsReadOnly="True" Margin="0,0,0,6" Background="White" Foreground="#0F172A" />

        <StackPanel Grid.Row="8" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button Name="SaveLicenseBtn" Content="Save" Width="90" Height="32" Margin="0,0,12,0" />
            <Button Name="CloseLicenseBtn" Content="Close" Width="90" Height="32" />
        </StackPanel>
    </Grid>
</Window>
"@

        $reader = New-Object System.Xml.XmlNodeReader $licenseXaml
        $licenseWindow = [Windows.Markup.XamlReader]::Load($reader)
        if ($Owner) {
            $licenseWindow.Owner = $Owner
        }

        $licenseWindow.FindName('TechnicianText').Text = $license.Technician
        $licenseWindow.FindName('CompanyText').Text = $license.Company
        $licenseWindow.FindName('AddressText').Text = $license.Address
        $licenseWindow.FindName('PhoneText').Text = $license.Phone
        $licenseWindow.FindName('EmailText').Text = $license.Email
        $licenseWindow.FindName('WebsiteText').Text = $license.Website
        $licenseWindow.FindName('LicenseText').Text = if ($license.License) { $license.License } else { [guid]::NewGuid().ToString() }

        $saveBtn = $licenseWindow.FindName('SaveLicenseBtn')
        $closeBtn = $licenseWindow.FindName('CloseLicenseBtn')
        $licenseFieldNames = @(
            'TechnicianText',
            'CompanyText',
            'AddressText',
            'PhoneText',
            'EmailText',
            'WebsiteText'
        )
        $setLicenseEditable = {
            param([bool]$Editable)
            foreach ($fieldName in $licenseFieldNames) {
                $field = $licenseWindow.FindName($fieldName)
                if ($field) {
                    $field.IsReadOnly = -not $Editable
                }
            }
            if ($saveBtn) {
                $saveBtn.Visibility = if ($Editable) { 'Visible' } else { 'Collapsed' }
            }
        }
        $licenseAlreadySaved = $false
        if ($license -and $license.PSObject.Properties['Saved']) {
            $licenseAlreadySaved = [bool]$license.Saved
        }
        & $setLicenseEditable (-not $licenseAlreadySaved)
        $saveBtn.Add_Click({
            $data = [pscustomobject]@{
                Technician = $licenseWindow.FindName('TechnicianText').Text.Trim()
                Company = $licenseWindow.FindName('CompanyText').Text.Trim()
                Address = $licenseWindow.FindName('AddressText').Text.Trim()
                Phone = $licenseWindow.FindName('PhoneText').Text.Trim()
                Email = $licenseWindow.FindName('EmailText').Text.Trim()
                Website = $licenseWindow.FindName('WebsiteText').Text.Trim()
                License = $licenseWindow.FindName('LicenseText').Text.Trim()
            }
            Save-LicenseData -Data $data
            Show-Info 'License information saved.'
            & $setLicenseEditable $false
        })
        $closeBtn.Add_Click({ $licenseWindow.Close() })
        $licenseWindow.ShowDialog() | Out-Null
    }

    function New-SystemRestorePoint {
        param([string]$ProductName)

        # PS 5.1-safe (no C#-style ternary)
        $descName = if ($ProductName) { $ProductName } else { 'Office' }
        $description = "Office Deployment Tool - $descName"
        try {
            if (-not (Get-Command -Name Checkpoint-Computer -ErrorAction SilentlyContinue)) {
                return
            }
            Checkpoint-Computer -Description $description -RestorePointType 'ApplicationInstall' -ErrorAction Stop
            Write-LogEntry -Type 'install' -Message "Created system restore point '$description'."
        } catch {
            Write-Warning "Unable to create restore point: $_"
        }
    }

    function Show-ProgressWindow {
        param(
            [Parameter(Mandatory)][ScriptBlock]$Work,
            [object[]]$Parameters = @(),
            [string]$Title = 'Working',
            [string]$Message = 'Please wait...'
        )

        $form = New-Object System.Windows.Forms.Form
        $form.Text = $Title
        $form.Width = 360
        $form.Height = 140
        $form.StartPosition = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        $form.MinimizeBox = $false
        $form.MaximizeBox = $false
        $form.TopMost = $true
        $form.ShowInTaskbar = $false

        $label = New-Object System.Windows.Forms.Label
        $label.Text = $Message
        $label.AutoSize = $true
        $label.Location = New-Object System.Drawing.Point(18, 18)

        $progress = New-Object System.Windows.Forms.ProgressBar
        $progress.Style = 'Marquee'
        $progress.MarqueeAnimationSpeed = 30
        $progress.Width = 320
        $progress.Height = 20
        $progress.Location = New-Object System.Drawing.Point(18, 50)

        $form.Controls.AddRange(@($label,$progress))

        $job = Start-Job -ScriptBlock $Work -ArgumentList $Parameters
        try {
            $form.Show()
            while ($job.State -eq 'Running') {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 100
            }
            $child = $job.ChildJobs[0]
            $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
            if ($child.JobStateInfo.State -eq 'Failed') {
                throw $child.JobStateInfo.Reason
            }
            return $result
        } finally {
            if ($job) {
                $job | Remove-Job -Force
            }
            $form.Close()
        }
    }

    function Get-LanguageNames {
        param([string[]]$Ids)
        if (-not $Ids) { return @() }
        $results = @()
        foreach ($langId in $Ids) {
            $match = $languageCatalog | Where-Object { $_.Id -eq $langId }
            if ($match) {
                $results += $match.Name
            } else {
                $results += $langId
            }
        }
        return $results
    }

    function Initialize-LanguageSelection {
        if (-not $script:SelectedLanguageIds -or $script:SelectedLanguageIds.Count -eq 0) {
            $script:SelectedLanguageIds = @('en-us')
        }
    }

    function Show-DownloadFolderDialog {
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = 'Select download folder for the Office bits'
        $dlg.ShowNewFolderButton = $true
        # PS 5.1: use the enum for RootFolder
        $dlg.RootFolder = [System.Environment+SpecialFolder]::MyComputer
        $dlg.SelectedPath = if ($script:DownloadPath) { $script:DownloadPath } else { [Environment]::GetFolderPath('Desktop') }
        $dialog = $dlg.ShowDialog()
        if ($dialog -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dlg.SelectedPath
        }
        return $null
    }

    function Show-LanguagePicker {
        $languageItems = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        foreach ($lang in $languageCatalog) {
            $languageItems.Add([pscustomobject]@{
                    Selected = $script:SelectedLanguageIds -contains $lang.Id
                    Name = $lang.Name
                    Id = $lang.Id
                }) | Out-Null
        }

        [xml]$langXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Languages" Height="342" Width="320" WindowStartupLocation="CenterOwner"
        Background="#1F1F28" Foreground="White" FontFamily="Segoe UI">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <DataGrid Grid.Row="0" ItemsSource="{Binding}" AutoGenerateColumns="False"
                  Foreground="White"
                  RowBackground="#2A2A34" AlternatingRowBackground="#33333E"
                  HeadersVisibility="Column" CanUserAddRows="False" CanUserDeleteRows="False"
                  ColumnWidth="*" >
            <DataGrid.ColumnHeaderStyle>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Foreground" Value="#0F172A" />
                    <Setter Property="Background" Value="#E2E8F0" />
                </Style>
            </DataGrid.ColumnHeaderStyle>
            <DataGrid.Columns>
                <DataGridCheckBoxColumn Binding="{Binding Selected}" Header="Use" Width="60" />
                <DataGridTextColumn Binding="{Binding Name}" Header="Language" Width="*" />
                <DataGridTextColumn Binding="{Binding Id}" Header="Locale" Width="110" />
            </DataGrid.Columns>
        </DataGrid>
        <StackPanel Grid.Row="1" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button Name="LangCancelBtn" Content="Cancel" Width="90" Margin="0,0,8,0" />
            <Button Name="LangApplyBtn" Content="Apply" Width="120" Background="#3A88F6" />
        </StackPanel>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $langXaml
        $langWindow = [Windows.Markup.XamlReader]::Load($reader)
        $langWindow.DataContext = $languageItems
        $langApply = $langWindow.FindName('LangApplyBtn')
        $langCancel = $langWindow.FindName('LangCancelBtn')

        $langApply.Add_Click({
                $langWindow.DialogResult = $true
                $langWindow.Close()
            })
        $langCancel.Add_Click({
                $langWindow.DialogResult = $false
                $langWindow.Close()
            })

        $langWindow.Owner = $window
        $dialogResult = $langWindow.ShowDialog()
        if ($dialogResult) {
            $selection = @($languageItems | Where-Object { $_.Selected } | Select-Object -ExpandProperty Id)
            if ($selection.Count -eq 0) {
                Show-Error 'Select at least one language.'
                return
            }
            $script:SelectedLanguageIds = $selection
            $friendly = Get-LanguageNames -Ids $selection
            Show-Info ("Languages updated to: {0}" -f ($friendly -join ', '))
        }
    }

    function Update-SettingsMenuState {
        if ($controls -and $controls.Install32BitMenu) {
            $controls.Install32BitMenu.IsChecked = $script:Is32BitArchitecture
        }
        if ($controls -and $controls.DisplayNoneMenu) {
            $controls.DisplayNoneMenu.IsChecked = $script:IsDisplayNone
        }
        if ($controls -and $controls.ShowVolumeEditionsMenu) {
            $controls.ShowVolumeEditionsMenu.IsChecked = $script:ShowVolumeEditions
        }
        if ($controls -and $controls.DisableRestorePointMenu) {
            $controls.DisableRestorePointMenu.IsChecked = $script:DisableRestorePoint
        }
        if ($controls -and $controls.DisableRemoveOfficeMenu) {
            $controls.DisableRemoveOfficeMenu.IsChecked = $script:DisableRemoveOffice
        }
        if ($controls -and $controls.ForceAppCloseMenu) {
            $controls.ForceAppCloseMenu.IsChecked = $script:ForceAppClose
        }
        if ($controls -and $controls.InstallProjectMenu) {
            $controls.InstallProjectMenu.IsChecked = $script:InstallProject
        }
        if ($controls -and $controls.InstallVisioMenu) {
            $controls.InstallVisioMenu.IsChecked = $script:InstallVisio
        }
    }

    function New-ODTConfigurationXml {
        param(
            [pscustomobject]$State
        )

        $license = if ($State.License) { $State.License } else { 'Retail' }
        $productId = if ($State.ProductId) { $State.ProductId } else { 'ProPlusRetail' }
        $languageIds = if ($State.PSObject.Properties.Match('Languages').Count -gt 0 -and $State.Languages) {
            $State.Languages
        } else {
            $script:SelectedLanguageIds
        }

        $xml = New-Object System.Xml.XmlDocument
        $xml.AppendChild($xml.CreateXmlDeclaration('1.0','UTF-8',$null)) | Out-Null
        $config = $xml.CreateElement('Configuration')
        $xml.AppendChild($config) | Out-Null

        $add = $xml.CreateElement('Add')
        $add.SetAttribute('OfficeClientEdition',$State.Architecture)
        $add.SetAttribute('Channel',$State.ChannelValue)
        if ($State.DownloadPath) {
            $add.SetAttribute('SourcePath',$State.DownloadPath)
        }
        if ($State.VersionMode -eq 'Specific' -and $State.Version) {
            $add.SetAttribute('Version',$State.Version)
        }
        $config.AppendChild($add) | Out-Null

        $product = $xml.CreateElement('Product')
        $product.SetAttribute('ID',$productId)
        if ($license -eq 'Volume' -and $State.ProductKey) {
            $product.SetAttribute('PIDKEY',$State.ProductKey)
        }
        $add.AppendChild($product) | Out-Null

        foreach ($langId in $languageIds) {
            $langNode = $xml.CreateElement('Language')
            $langNode.SetAttribute('ID',$langId)
            $product.AppendChild($langNode) | Out-Null
        }

        if ($State.InstallProject) {
            $projectProduct = $xml.CreateElement('Product')
            $projectProduct.SetAttribute('ID','ProjectPro2021Retail')
            foreach ($langId in $languageIds) {
                $langNode = $xml.CreateElement('Language')
                $langNode.SetAttribute('ID',$langId)
                $projectProduct.AppendChild($langNode) | Out-Null
            }
            $add.AppendChild($projectProduct) | Out-Null
        }

        if ($State.InstallVisio) {
            $visioProduct = $xml.CreateElement('Product')
            $visioProduct.SetAttribute('ID','VisioPro2021Retail')
            foreach ($langId in $languageIds) {
                $langNode = $xml.CreateElement('Language')
                $langNode.SetAttribute('ID',$langId)
                $visioProduct.AppendChild($langNode) | Out-Null
            }
            $add.AppendChild($visioProduct) | Out-Null
        }

        if ($State.ProductKey -and $license -eq 'Retail') {
            $property = $xml.CreateElement('Property')
            $property.SetAttribute('Name','ProductKey')
            $property.SetAttribute('Value',$State.ProductKey)
            $config.AppendChild($property) | Out-Null
        }

        $display = $xml.CreateElement('Display')
        $display.SetAttribute('Level',$State.DisplayLevel)
        $display.SetAttribute('AcceptEULA','TRUE')
        $config.AppendChild($display) | Out-Null

        if (-not $State.DisableRemoveOffice) {
            $removeMsi = $xml.CreateElement('RemoveMSI')
            $config.AppendChild($removeMsi) | Out-Null

            $removeAll = $xml.CreateElement('Remove')
            $removeAll.SetAttribute('All','TRUE')
            $config.AppendChild($removeAll) | Out-Null
        }

        if ($State.EnabledApps) {
            foreach ($app in $State.EnabledApps.GetEnumerator()) {
                if (-not $app.Value) {
                    $exclude = $xml.CreateElement('ExcludeApp')
                    $exclude.SetAttribute('ID',$app.Key)
                    $config.AppendChild($exclude) | Out-Null
                }
            }
        }

        $forceProperty = $xml.CreateElement('Property')
        $forceProperty.SetAttribute('Name','FORCEAPPSHUTDOWN')
        $forceValue = if ($State.ForceAppClose) { 'FALSE' } else { 'TRUE' }
        $forceProperty.SetAttribute('Value', $forceValue)
        $config.AppendChild($forceProperty) | Out-Null

        $propertyUpdates = $xml.CreateElement('Updates')
        $propertyUpdates.SetAttribute('Enabled','TRUE')
        $propertyUpdates.SetAttribute('Channel',$State.ChannelValue)
        $config.AppendChild($propertyUpdates) | Out-Null

        return $xml
    }

function Get-UIState {
    Initialize-LanguageSelection

    $productName      = if ($controls.ProductCombo.SelectedItem) { $controls.ProductCombo.SelectedItem.Name } else { $null }
    $editionItem      = $controls.EditionCombo.SelectedItem
    $editionName      = if ($editionItem) { $editionItem.Name } else { $null }
    $editionProductId = if ($editionItem) { $editionItem.ProductId } else { $null }

    $channelItem  = $controls.ChannelCombo.SelectedItem
    $channelLabel = if ($channelItem) { $channelItem.Label } else { $null }
    $channelValue = if ($channelItem) { $channelItem.Value } else { $null }

    $versionMode = if ($controls.VersionModeCombo.SelectedItem) { $controls.VersionModeCombo.SelectedItem.Value } else { $null }
    $version     = $controls.VersionText.Text.Trim()
    $prodKey     = $controls.ProductKeyText.Text.Trim()

    $arch         = if ($script:Is32BitArchitecture) { '32' } else { '64' }
    $displayLevel = if ($script:IsDisplayNone) { 'None' } else { 'Full' }
    $prodId       = if ($editionProductId) { $editionProductId } else { 'ProPlusRetail' }
    $dlPath       = $script:DownloadPath

    [pscustomobject]@{
        Product             = $productName
        Edition             = $editionName
        EditionProductId    = $editionProductId
        ChannelLabel        = $channelLabel
        ChannelValue        = $channelValue
        VersionMode         = $versionMode
        Version             = $version
        ProductKey          = $prodKey
        Languages           = @($script:SelectedLanguageIds)
        License             = if ($editionItem) { $editionItem.License } else { $null }
        Architecture        = $arch
        DisplayLevel        = $displayLevel
        ShowVolumeEditions  = $script:ShowVolumeEditions
        DisableRestorePoint = $script:DisableRestorePoint
        DisableRemoveOffice = $script:DisableRemoveOffice
        ForceAppClose       = $script:ForceAppClose
        InstallProject      = $script:InstallProject
        InstallVisio        = $script:InstallVisio
        EnabledApps         = $script:AppSelections
        ProductId           = $prodId
        DownloadPath        = $dlPath
    }
}


function Get-SelectionMetadata {
    param([pscustomobject]$State)

    $prod    = if ($State.Product) { $State.Product } else { 'Unknown' }
    $edition = if ($State.Edition) { $State.Edition } else { 'Unknown' }
    $ver     = if ($State.VersionMode -eq 'Specific' -and $State.Version) { $State.Version } else { 'Latest' }
    $chan    = if ($State.ChannelValue) { $State.ChannelValue } else { 'Default' }
    $key     = if ($State.ProductKey) { $State.ProductKey } else { 'None' }

    $meta = [ordered]@{
        Product    = $prod
        Edition    = $edition
        Version    = $ver
        Channel    = $chan
        ProductKey = $key
    }

    if ($State.PSObject.Properties.Match('DownloadPath').Count -gt 0 -and $State.DownloadPath) {
        $meta.DownloadedTo = $State.DownloadPath
    }

    return [pscustomobject]$meta
}

    function Save-CurrentDefaults {
        try {
            $state = Get-UIState
            $json = $state | ConvertTo-Json -Depth 5
            Set-Content -Path $script:DefaultsPath -Value $json -Encoding UTF8
            Show-Info "Defaults saved to $($script:DefaultsPath)."
            $script:DefaultsSaved = $true
        } catch {
            Show-Error "Failed to save defaults. $_"
        }
    }

    function Show-SaveDefaultsAsDialog {
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.InitialDirectory = $script:AppRoot
        $dlg.FileName = Split-Path $script:DefaultsPath -Leaf
        $dlg.Filter = 'JSON Files (*.json)|*.json'
        $dlg.Title = 'Save Configuration As'
        if ($dlg.ShowDialog()) {
            $script:DefaultsPath = $dlg.FileName
            Save-CurrentDefaults
            $script:DefaultsSaved = $true
        }
    }

    function Save-CurrentConfiguration {
        if (-not $script:DefaultsSaved -or -not (Test-Path -LiteralPath $script:DefaultsPath)) {
            Show-SaveDefaultsAsDialog
            return
        }
        Save-CurrentDefaults
    }

    function Reset-UIToDefaults {
        Clear-Defaults
        Initialize-LanguageSelection
        $script:SelectedLanguageIds = @('en-us')
        $script:Is32BitArchitecture = $false
        $script:IsDisplayNone = $false
        $script:ShowVolumeEditions = $false
        $script:DisableRestorePoint = $false
        $script:DisableRemoveOffice = $false
        $script:ForceAppClose = $true
        Reset-AppSelections
        if ($controls) {
            $controls.ProductCombo.SelectedIndex = 0
            Update-EditionOptions -PreserveSelection $false
            $controls.ChannelCombo.SelectedIndex = 0
            $controls.VersionModeCombo.SelectedIndex = 0
            Set-VersionTextState
            $controls.ProductKeyText.Text = ''
            Update-SettingsMenuState
        }
    }

    function Show-GeneratedScriptsWindow {
        if (-not (Test-Path -LiteralPath $script:GeneratedPath)) {
            Show-Error 'Generated folder not found.'
            return
        }
        $dlg = New-Object Microsoft.Win32.OpenFileDialog
        $dlg.InitialDirectory = $script:GeneratedPath
        $dlg.Filter = 'Scripts (*.xml)|*.xml'
        $dlg.Title = 'Open Generated Script'
        if ($dlg.ShowDialog()) {
            Start-Process -FilePath $dlg.FileName
        }
    }

    function Show-CurrentScriptWindow {
        param([pscustomobject]$State)

        if (-not $State) {
            $State = Get-UIState
        }
        $xmlDoc = New-ODTConfigurationXml -State $State
        $settings = New-Object System.Xml.XmlWriterSettings
        $settings.Indent = $true
        $settings.OmitXmlDeclaration = $false
        $sb = New-Object System.Text.StringBuilder
        $writer = [System.Xml.XmlWriter]::Create($sb,$settings)
        $xmlDoc.Save($writer)
        $writer.Flush()
        $scriptText = $sb.ToString()

        [xml]$scriptXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Current Configuration" Height="520" Width="640"
        Background="#111827" Foreground="White" FontFamily="Consolas" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <TextBox Name="ScriptText" Grid.Row="0" TextWrapping="Wrap" AcceptsReturn="True"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"
                 Background="#0F172A" Foreground="White" IsReadOnly="True" FontFamily="Consolas" FontSize="12" />
        <Button Name="CloseScriptBtn" Grid.Row="1" Content="Close" Width="90" Height="32" HorizontalAlignment="Right" Margin="0,12,0,0" />
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $scriptXaml
        try {
            $scriptWindow = [Windows.Markup.XamlReader]::Load($reader)
        } catch {
            Show-Error "Unable to render script preview: $_"
            return
        }
        if (-not $scriptWindow) {
            Show-Error 'Unable to display script preview window.'
            return
        }
        $scriptTextBox = $scriptWindow.FindName('ScriptText')
        if ($scriptTextBox) {
            $scriptTextBox.Text = $scriptText
        }
        $closeBtn = $scriptWindow.FindName('CloseScriptBtn')
        if ($closeBtn) {
            $closeBtn.Add_Click({ $scriptWindow.Close() })
        }
        $scriptWindow.ShowDialog() | Out-Null
    }

    function Show-AppSelectionWindow {
        $columnTemplates = @{
            0 = ($script:AppSelectionOptions | Where-Object { $_.Column -eq 0 })
            1 = ($script:AppSelectionOptions | Where-Object { $_.Column -eq 1 })
        }
        $checkBoxMarkup = @{}
        foreach ($column in 0..1) {
            $checkBoxMarkup[$column] = ($columnTemplates[$column] | ForEach-Object {
                $safeName = "AppCheck$([regex]::Replace($_.Name,'[^0-9A-Za-z]',''))"
                "<CheckBox Name=`"$safeName`" Content=`"$($_.Name)`" Margin=`"0,6,0,0`" Foreground=`"#FFFFFF`" />"
            }) -join "`n"
        }

        $message = "Turn apps on or off to include or exclude them from being deployed."
        [xml]$appXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Applications" Height="250" Width="325" Background="#0F172A" Foreground="White"
        FontFamily="Segoe UI" WindowStartupLocation="CenterOwner">
    <DockPanel Margin="12">
        <TextBlock Text="$message" TextWrapping="Wrap" DockPanel.Dock="Top" Margin="0,0,0,12" />
        <Grid DockPanel.Dock="Top">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <StackPanel Grid.Column="0">
$($checkBoxMarkup[0])
            </StackPanel>
            <StackPanel Grid.Column="1">
$($checkBoxMarkup[1])
            </StackPanel>
        </Grid>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" DockPanel.Dock="Bottom" Margin="0,12,0,0">
            <Button Name="CloseAppsBtn" Content="Close" Width="90" Height="32" Margin="0,0,8,0" />
            <Button Name="ApplyAppsBtn" Content="Apply" Width="90" Height="32" Background="#2563EB" />
        </StackPanel>
    </DockPanel>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $appXaml
        $appWindow = [Windows.Markup.XamlReader]::Load($reader)
        if (-not $appWindow) {
            Show-Error 'Unable to open applications dialog.'
            return
        }
        foreach ($option in $script:AppSelectionOptions) {
            $safeName = "AppCheck$([regex]::Replace($option.Name,'[^0-9A-Za-z]',''))"
            $check = $appWindow.FindName($safeName)
            if ($check) {
                $check.IsChecked = [bool]$script:AppSelections[$option.Name]
                $check.Add_Checked({ $script:AppSelections[$option.Name] = $true })
                $check.Add_Unchecked({ $script:AppSelections[$option.Name] = $false })
            }
        }
        $apply = $appWindow.FindName('ApplyAppsBtn')
        $close = $appWindow.FindName('CloseAppsBtn')
        if ($apply) {
            $apply.Add_Click({
                $appWindow.Close()
            })
        }
        if ($close) {
            $close.Add_Click({ $appWindow.Close() })
        }
        $appWindow.ShowDialog() | Out-Null
    }

    function Invoke-RemovePreviousInstallsScript {
        if (-not (Test-Path -LiteralPath $script:OdtSetupExePath)) {
            Show-Error 'The Office Deployment Tool is not available. Run the GUI once to download it first.'
            return
        }
        $configPath = Join-Path $script:GeneratedPath 'remove-previous-installs.xml'
        $configContent = @"
<Configuration>
  <Display Level="Full" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <RemoveMSI />
  <Remove All="TRUE" />
</Configuration>
"@
        Set-Content -LiteralPath $configPath -Value $configContent -Encoding UTF8
        try {
            Start-Process -FilePath $script:OdtSetupExePath -ArgumentList '/configure', "`"$configPath`"" -Wait -NoNewWindow
            Show-Info 'Previous Office installations were removed. Check the log for details.'
        } catch {
            Show-Error "Failed to execute removal script: $_"
        }
    }

    function Show-RemovePreviousInstallsWindow {
        $message = "By continuing all opened Office applications will be forced closed and all previously MSI and Click-to-Run installations of Office and Office Apps will be removed from the system.`n`nIf all your work has been saved and you wish to continue click Proceed otherwise click Cancel."
        [xml]$removeXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Remove Office" Height="230" Width="420"
        Background="#0F172A" Foreground="White" FontFamily="Segoe UI" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <TextBlock Text="$message" TextWrapping="Wrap" Grid.Row="0" />
        <StackPanel Orientation="Horizontal" Grid.Row="1" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button Name="CancelRemoveBtn" Content="Cancel" Width="90" Height="32" Margin="0,0,8,0" />
            <Button Name="ProceedRemoveBtn" Content="Proceed" Width="90" Height="32" Background="#DC2626" />
        </StackPanel>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $removeXaml
        $removeWindow = [Windows.Markup.XamlReader]::Load($reader)
        $cancel = $removeWindow.FindName('CancelRemoveBtn')
        $proceed = $removeWindow.FindName('ProceedRemoveBtn')
        $cancel.Add_Click({ $removeWindow.Close() })
        $proceed.Add_Click({
            $removeWindow.Close()
            Invoke-RemovePreviousInstallsScript
        })
        $removeWindow.ShowDialog() | Out-Null
    }

    function Clear-Defaults {
        if (Test-Path -LiteralPath $script:DefaultsPath) {
            Remove-Item -LiteralPath $script:DefaultsPath -Force
        }
        Show-Info 'Defaults cleared. Current session uses in-memory selections.'
    }

    function Get-AppSafeName {
        param([string]$Name)
        return ($Name -replace '[^0-9A-Za-z]', '')
    }

    function Reset-AppSelections {
        $script:AppSelections = @{}
        foreach ($option in $script:AppSelectionOptions) {
            $script:AppSelections[$option.Name] = $option.Default
        }
    }

    function Update-AppSelectionsFromState {
        param([pscustomobject]$State)
        Reset-AppSelections
        if ($State -and $State.PSObject.Properties.Match('EnabledApps').Count -gt 0) {
            foreach ($entry in $State.EnabledApps.GetEnumerator()) {
                if ($script:AppSelections.Contains($entry.Key)) {
                    $script:AppSelections[$entry.Key] = [bool]$entry.Value
                }
            }
        }
    }

    function Set-UiState {
        param($state)
        if (-not $state) { return }

        $product = $officeCatalog | Where-Object { $_.Name -eq $state.Product }
        if ($product) {
            $controls.ProductCombo.SelectedItem = $product
            Update-EditionOptions -PreserveSelection $true
        }
        if ($state.Edition) {
            $edition = $null
            if ($state.EditionProductId) {
                $edition = $controls.EditionCombo.Items | Where-Object { $_.ProductId -eq $state.EditionProductId }
            }
            if (-not $edition) {
                $edition = $controls.EditionCombo.Items | Where-Object { $_.Name -eq $state.Edition }
            }
            if ($edition) { $controls.EditionCombo.SelectedItem = $edition }
        }
        if ($state.ChannelValue) {
            $channel = $channelOptions | Where-Object { $_.Value -eq $state.ChannelValue }
            if ($channel) { $controls.ChannelCombo.SelectedItem = $channel }
        }
        if ($state.VersionMode) {
            $mode = $versionModes | Where-Object { $_.Value -eq $state.VersionMode }
            if ($mode) { $controls.VersionModeCombo.SelectedItem = $mode }
        }
        if ($state.Version) {
            $controls.VersionText.Text = $state.Version
        }
        if ($state.ProductKey) { $controls.ProductKeyText.Text = $state.ProductKey }
        if ($state.Languages) {
            $script:SelectedLanguageIds = @($state.Languages)
        } else {
            Initialize-LanguageSelection
        }
        if ($state.PSObject.Properties.Match('ShowVolumeEditions').Count -gt 0) {
            $script:ShowVolumeEditions = [bool]$state.ShowVolumeEditions
        } else {
            $script:ShowVolumeEditions = $false
        }
        if ($state.PSObject.Properties.Match('DisableRestorePoint').Count -gt 0) {
            $script:DisableRestorePoint = [bool]$state.DisableRestorePoint
        } else {
            $script:DisableRestorePoint = $false
        }
        if ($state.PSObject.Properties.Match('DisableRemoveOffice').Count -gt 0) {
            $script:DisableRemoveOffice = [bool]$state.DisableRemoveOffice
        } else {
            $script:DisableRemoveOffice = $false
        }
        if ($state.PSObject.Properties.Match('ForceAppClose').Count -gt 0) {
            $script:ForceAppClose = [bool]$state.ForceAppClose
        } else {
            $script:ForceAppClose = $true
        }
        if ($state.PSObject.Properties.Match('InstallProject').Count -gt 0) {
            $script:InstallProject = [bool]$state.InstallProject
        } else {
            $script:InstallProject = $false
        }
        if ($state.PSObject.Properties.Match('InstallVisio').Count -gt 0) {
            $script:InstallVisio = [bool]$state.InstallVisio
        } else {
            $script:InstallVisio = $false
        }
        Update-AppSelectionsFromState -State $state
        Update-EditionOptions -PreserveSelection $true
        $script:Is32BitArchitecture = $state.Architecture -eq '32'
        $script:IsDisplayNone = $state.DisplayLevel -eq 'None'
        Update-SettingsMenuState
    }

    function Get-DefaultsFile {
        if (-not (Test-Path -LiteralPath $script:DefaultsPath)) { return $null }
        try {
            return Get-Content -Path $script:DefaultsPath -Raw | ConvertFrom-Json
        } catch {
            Write-Warning "Failed to parse defaults file: $_"
            return $null
        }
    }

    function Get-FilteredEditions {
        param([pscustomobject]$Product)
        if (-not $Product) { return @() }
        $editions = if ($script:ShowVolumeEditions) {
            $Product.Editions
        } else {
            $Product.Editions | Where-Object { $_.License -eq 'Retail' }
        }
        if (-not $editions -or $editions.Count -eq 0) {
            return $Product.Editions
        }
        return $editions
    }

    function Update-EditionOptions {
        param([switch]$PreserveSelection)
        $selectedProduct = $controls.ProductCombo.SelectedItem
        if (-not $selectedProduct) { return }
        if ($selectedProduct.Name -eq 'Select Product') {
            $controls.EditionCombo.ItemsSource = @()
            $controls.EditionCombo.IsEnabled = $false
            return
        }
        $editionSource = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        foreach ($ed in (Get-FilteredEditions -Product $selectedProduct)) {
            [void]$editionSource.Add($ed)
        }
        $previous = $controls.EditionCombo.SelectedItem
        $controls.EditionCombo.ItemsSource = $editionSource
        if ($PreserveSelection -and $previous) {
            $match = $editionSource | Where-Object { $_.Name -eq $previous.Name }
            if ($match) {
                $controls.EditionCombo.SelectedItem = $match
                return
            }
        }
        if ($editionSource.Count -eq 0) {
            $controls.EditionCombo.IsEnabled = $false
            $controls.EditionCombo.SelectedIndex = -1
        } else {
            $controls.EditionCombo.IsEnabled = $true
            $controls.EditionCombo.SelectedIndex = 0
        }
    }

    function Set-VersionTextState {
        $versionModeItem = $controls.VersionModeCombo.SelectedItem
        if ($versionModeItem -and $versionModeItem.Value -eq 'Specific') {
            $controls.VersionText.IsEnabled = $true
        } else {
            $controls.VersionText.IsEnabled = $false
            $controls.VersionText.Text = ''
        }
    }

    function Invoke-ODTAction {
        param(
            [Parameter(Mandatory)][ValidateSet('Download','Install')]$Action
        )
        $state = Get-UIState
        if (-not $state.Product -or -not $state.Edition) {
            Show-Error 'Select an Office product and edition.'
            return
        }
        if ($state.VersionMode -eq 'Specific' -and -not $state.Version) {
            Show-Error 'Enter a version or switch to Latest Available.'
            return
        }
        if ($Action -eq 'Download') {
            $downloadFolder = Show-DownloadFolderDialog
            if (-not $downloadFolder) {
                Show-Info 'Download canceled.'
                return
            }
            New-DirectoryIfMissing -Path $downloadFolder
            $script:DownloadPath = $downloadFolder
            $state | Add-Member -MemberType NoteProperty -Name 'DownloadPath' -Value $downloadFolder -Force
        }

        $xml = New-ODTConfigurationXml -State $state
        $configPath = Join-Path $script:GeneratedPath ("configuration-{0}.xml" -f $Action.ToLower())
        $xml.Save($configPath)

        $setupExe = $script:OdtSetupExePath
        $logType = if ($Action -eq 'Download') { 'download' } else { 'install' }

        if (-not (Test-Path -LiteralPath $setupExe)) {
            Show-Error 'ODT.exe is missing. The downloader failed to fetch the executable.'
            return
        }

        $argument = if ($Action -eq 'Download') {
            "/download `"$configPath`""
        } else {
            "/configure `"$configPath`""
        }
        try {
            if ($Action -eq 'Install') {
                if (-not $script:DisableRestorePoint) {
                    New-SystemRestorePoint -ProductName $state.Product
                } else {
                    Write-LogEntry -Type 'install' -Message 'Skipped restore point creation per Settings.' -Metadata (Get-SelectionMetadata -State $state)
                }
                $productName = if ($state.Product) { $state.Product } else { 'Office' }
                $exitCode = Show-ProgressWindow -Title 'Installing Office' -Message "Installing $productName..." -Work {
                    param($exe,$installArgs)
                    $process = Start-Process -FilePath $exe -ArgumentList $installArgs -PassThru -Wait
                    return $process.ExitCode
                } -Parameters @($setupExe,$argument)
            } else {
                $process = Start-Process -FilePath $setupExe -ArgumentList $argument -PassThru -Wait
                $exitCode = $process.ExitCode
            }
            $selectionMeta = Get-SelectionMetadata -State $state
            Write-LogEntry -Type $logType -Message "$Action completed with exit code $exitCode." -Metadata $selectionMeta
            Show-Info "$Action completed (exit code $exitCode)."
        } catch {
            $selectionMeta = Get-SelectionMetadata -State $state
            Write-LogEntry -Type $logType -Message "$Action failed: $_" -Metadata $selectionMeta
            Show-Error "Failed to run $Action. $_"
        }
    }

    function Show-LogEntryDetails {
        param(
            [Parameter(Mandatory)][pscustomobject]$Entry,
            [System.Windows.Window]$Owner,
            [Parameter()][string]$LogFile,
            [string]$LogTitle = 'Log Entry'
        )

        $timestamp = $Entry.Timestamp
        $message = $Entry.Message
        $hostValue = $Entry.Host
        $metadata = if ($Entry.Metadata) { $Entry.Metadata } else { [pscustomobject]@{} }
        if (-not ($metadata -is [System.Management.Automation.PSCustomObject])) {
            $metadata = [pscustomobject]$metadata
        }
        foreach ($requiredProp in @('Product','Edition','Version','Channel','ProductKey','Account')) {
            if (-not $metadata.PSObject.Properties[$requiredProp]) {
                $metadata | Add-Member -MemberType NoteProperty -Name $requiredProp -Value '' -Force
            }
        }

        [xml]$detailXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$LogTitle" Height="420" Width="400" Background="#0F172A" Foreground="White"
        FontFamily="Segoe UI" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="120" />
            <ColumnDefinition Width="*" />
        </Grid.ColumnDefinitions>

        <TextBlock Grid.Row="0" Grid.Column="0" Text="Timestamp:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="0" Grid.Column="1" Name="TimestampText" IsReadOnly="True" Margin="0,0,0,6" />

        <TextBlock Grid.Row="1" Grid.Column="0" Text="Account:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="1" Grid.Column="1" Name="AccountText" Margin="0,0,0,6" />

        <TextBlock Grid.Row="2" Grid.Column="0" Text="Host:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="2" Grid.Column="1" Name="HostText" IsReadOnly="True" Margin="0,0,0,6" />

        <TextBlock Grid.Row="3" Grid.Column="0" Text="Product:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="3" Grid.Column="1" Name="ProductText" IsReadOnly="True" Margin="0,0,0,6" />

        <TextBlock Grid.Row="4" Grid.Column="0" Text="Edition:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="4" Grid.Column="1" Name="EditionText" IsReadOnly="True" Margin="0,0,0,6" />

        <TextBlock Grid.Row="5" Grid.Column="0" Text="Version:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="5" Grid.Column="1" Name="VersionText" IsReadOnly="True" Margin="0,0,0,6" />

        <TextBlock Grid.Row="6" Grid.Column="0" Text="Channel:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="6" Grid.Column="1" Name="ChannelText" IsReadOnly="True" Margin="0,0,0,6" />

        <TextBlock Grid.Row="7" Grid.Column="0" Text="ProductKey:" Margin="0,0,8,6" VerticalAlignment="Center" />
        <TextBox Grid.Row="7" Grid.Column="1" Name="ProductKeyText" IsReadOnly="True" Margin="0,0,0,6" />

        <TextBlock Grid.Row="8" Grid.Column="0" Text="Message:" Margin="0,0,8,6" VerticalAlignment="Top" />
        <TextBox Grid.Row="8" Grid.Column="1" Name="MessageText" TextWrapping="Wrap" AcceptsReturn="True"
                 VerticalScrollBarVisibility="Auto" Height="90" IsReadOnly="True" Background="#111827" Foreground="White" />

        <StackPanel Grid.Row="9" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
            <Button Name="SaveAccountBtn" Content="Save" Width="90" Height="32" Margin="0,0,8,0" />
            <Button Name="PrintAccountBtn" Content="Print" Width="90" Height="32" Margin="0,0,8,0" Visibility="Collapsed" />
            <Button Name="CloseBtn" Content="Close" Width="100" Height="32" />
        </StackPanel>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $detailXaml
        $detailWindow = [Windows.Markup.XamlReader]::Load($reader)
        if ($Owner) {
            $detailWindow.Owner = $Owner
        }
        $detailWindow.FindName('TimestampText').Text = $timestamp
        $detailWindow.FindName('HostText').Text = $hostValue
        $detailWindow.FindName('ProductText').Text = if ($metadata.Product) { $metadata.Product } else { '' }
        $detailWindow.FindName('EditionText').Text = if ($metadata.Edition) { $metadata.Edition } else { '' }
        $detailWindow.FindName('VersionText').Text = if ($metadata.Version) { $metadata.Version } else { '' }
        $detailWindow.FindName('ChannelText').Text = if ($metadata.Channel) { $metadata.Channel } else { '' }
        $detailWindow.FindName('ProductKeyText').Text = if ($metadata.ProductKey) { $metadata.ProductKey } else { '' }
        $detailWindow.FindName('MessageText').Text = $message

        $accountText = $detailWindow.FindName('AccountText')
        $saveAccountBtn = $detailWindow.FindName('SaveAccountBtn')
        $printAccountBtn = $detailWindow.FindName('PrintAccountBtn')

        if ($accountText) {
            $accountText.Text = if ($metadata.Account) { $metadata.Account } else { '' }
        }

        $setAccountEditable = {
            param([bool]$Editable)
            if ($accountText) {
                $accountText.IsReadOnly = -not $Editable
            }
            if ($saveAccountBtn) {
                $saveAccountBtn.Visibility = if ($Editable) { 'Visible' } else { 'Collapsed' }
            }
            if ($printAccountBtn) {
                $printAccountBtn.Visibility = if ($Editable) { 'Collapsed' } else { 'Visible' }
            }
        }

        $hasAccount = $accountText -and -not [string]::IsNullOrWhiteSpace($accountText.Text)
        & $setAccountEditable (-not $hasAccount)

        if ($saveAccountBtn) {
            $saveAccountBtn.Add_Click({
                if (-not $accountText) { return }
                $accountValue = $accountText.Text.Trim()
                if (-not $accountValue) {
                    Show-Error 'Enter an account value before saving.'
                    return
                }
                if (-not $LogFile) {
                    Show-Error 'Unable to save account without a log file path.'
                    return
                }
                if (Save-LogEntryAccount -LogFile $LogFile -Entry $Entry -Account $accountValue) {
                    Show-Info 'Account saved.'
                    $accountText.Text = $accountValue
                    & $setAccountEditable $false
                } else {
                    Show-Error 'Failed to persist account information.'
                }
            })
        }

        if ($printAccountBtn) {
            $printAccountBtn.Add_Click({
                $printer = New-Object System.Windows.Controls.PrintDialog
                if ($printer.ShowDialog()) {
                    $printer.PrintVisual($detailWindow.Content, "$LogTitle Entry")
                }
            })
        }

        $close = $detailWindow.FindName('CloseBtn')
        $close.Add_Click({ $detailWindow.Close() })
        $detailWindow.ShowDialog() | Out-Null
    }

    function Show-LogWindow {
        param(
            [Parameter(Mandatory)][string]$Title,
            [Parameter(Mandatory)][string]$LogFile
        )
        if (-not (Test-Path -LiteralPath $LogFile)) {
            Show-Info "Log file not found:`n$LogFile"
            return
        }
        $entries = @()
        try {
            $raw = Get-Content -LiteralPath $LogFile -Raw
            if ($raw) {
                $parsed = $raw | ConvertFrom-Json
                if ($parsed) {
                    $entries += $parsed
                }
            }
        } catch {
            Write-Warning "Unable to read JSON log: $_"
        }

        $displayEntries = @()
        $index = 1
        foreach ($entry in $entries) {
            if (-not $entry) { continue }
            $displayEntries += [pscustomobject]@{
                Index = $index++
                Timestamp = $entry.Timestamp
                Message = $entry.Message
                Host = $entry.Host
                EntryData = $entry
            }
        }

        [xml]$logXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$Title" Height="320" Width="700" Background="#101018" Foreground="White"
        FontFamily="Segoe UI" WindowStartupLocation="CenterOwner">
    <Grid Margin="12">
        <DataGrid Name="LogGrid" ItemsSource="{Binding}" AutoGenerateColumns="False" HeadersVisibility="Column"
                  RowBackground="#1E1E2A" AlternatingRowBackground="#252533"
                  CanUserAddRows="False" CanUserDeleteRows="False" IsReadOnly="True">
            <DataGrid.Resources>
                <Style x:Key="LogCellTextStyle" TargetType="TextBlock">
                    <Setter Property="Foreground" Value="White" />
                </Style>
            </DataGrid.Resources>
            <DataGrid.Columns>
                <DataGridTextColumn Binding="{Binding Index}" Header="#" Width="40" ElementStyle="{StaticResource LogCellTextStyle}" />
                <DataGridTextColumn Binding="{Binding Timestamp}" Header="Timestamp" Width="180" ElementStyle="{StaticResource LogCellTextStyle}" />
                <DataGridTextColumn Binding="{Binding Host}" Header="Host" Width="160" ElementStyle="{StaticResource LogCellTextStyle}" />
                <DataGridTextColumn Binding="{Binding Message}" Header="Message" Width="*" ElementStyle="{StaticResource LogCellTextStyle}" />
            </DataGrid.Columns>
        </DataGrid>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $logXaml
        $logWindow = [Windows.Markup.XamlReader]::Load($reader)
        $logWindow.Owner = $window
        $logWindow.DataContext = $displayEntries
        $logGrid = $logWindow.FindName('LogGrid')
        $logGrid.ItemsSource = $displayEntries
        $logGrid.Add_MouseDoubleClick({
            $selected = $logGrid.SelectedItem
            if ($selected -and $selected.EntryData) {
                Show-LogEntryDetails -Entry $selected.EntryData -Owner $logWindow -LogFile $LogFile -LogTitle ("{0} Entry" -f $Title)
            }
        })
        $logWindow.ShowDialog() | Out-Null
    }
    function Invoke-Repair {
        param([ValidateSet('QuickRepair','FullRepair')]$Mode)
        $repairArgs = @('scenario=Repair', "RepairType=$Mode", 'forceappshutdown=True', 'DisplayLevel=True')
        $exe = 'OfficeClickToRun.exe'
        if (-not (Get-Command $exe -ErrorAction SilentlyContinue)) {
            Show-Error "$exe not found in PATH. Launch the repair from an elevated Microsoft Office Command Prompt."
            return
        }
        try {
            Start-Process -FilePath $exe -ArgumentList $repairArgs -Verb RunAs
            Show-Info "$Mode started via OfficeClickToRun.exe."
        } catch {
            Show-Error "Failed to start $Mode. $_"
        }
    }

    function Show-PstRepairWindow {
        [xml]$pstXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PST Repair" Height="260" Width="520" Background="#0F172A" Foreground="White"
        FontFamily="Segoe UI" WindowStartupLocation="CenterOwner">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
        </Grid.RowDefinitions>
        <TextBlock Text="Outlook PST Path" FontSize="14" />
        <DockPanel Grid.Row="1" Margin="0,6,0,0">
            <Button Name="BrowsePstButton" Content="Browse" Width="80" Margin="0,0,8,0" DockPanel.Dock="Right" />
            <TextBox Name="PstPathText" Height="32" Background="#18213B" BorderBrush="#2F3B5E" />
        </DockPanel>
        <ProgressBar Name="PstProgress" Grid.Row="2" Height="18" Margin="0,18,0,0" Minimum="0" Maximum="100" />
        <Button Name="StartPstRepairButton" Grid.Row="3" Content="Repair PST" Height="40" Margin="0,18,0,0"
                Background="#3A88F6" />
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $pstXaml
        $pstWindow = [Windows.Markup.XamlReader]::Load($reader)
        $pstWindow.Owner = $window
        $pstPathText = $pstWindow.FindName('PstPathText')
        $browseBtn = $pstWindow.FindName('BrowsePstButton')
        $repairBtn = $pstWindow.FindName('StartPstRepairButton')
        $progress = $pstWindow.FindName('PstProgress')

        $browseBtn.Add_Click({
                $dlg = New-Object Microsoft.Win32.OpenFileDialog
                $dlg.Filter = 'Outlook PST (*.pst)|*.pst'
                if ($dlg.ShowDialog()) {
                    $pstPathText.Text = $dlg.FileName
                }
            })

        $repairBtn.Add_Click({
                $path = $pstPathText.Text.Trim()
                if (-not (Test-Path -LiteralPath $path)) {
                    Show-Error 'Select a valid PST file.'
                    return
                }
                $repairBtn.IsEnabled = $false
                $progress.Value = 0
                $exe = 'SCANPST.EXE'
                $scanner = Get-Command $exe -ErrorAction SilentlyContinue
                $timer = New-Object System.Windows.Threading.DispatcherTimer
                $timer.Interval = [TimeSpan]::FromMilliseconds(80)
                $timer.Add_Tick({
                        if ($progress.Value -ge 100) {
                            $timer.Stop()
                            $repairBtn.IsEnabled = $true
                            if ($scanner) {
                                try {
                                    Start-Process -FilePath $scanner.Source -ArgumentList "`"$path`"" -Verb RunAs
                                    Show-Info 'SCANPST launched to complete the repair.'
                                } catch {
                                    Show-Error "Failed to start SCANPST: $_"
                                }
                            } else {
                                Show-Info 'Progress complete. Install Outlook tools to run SCANPST.EXE automatically.'
                            }
                            return
                        }
                        $progress.Value += 5
                    })
                $timer.Start()
            })

        $pstWindow.ShowDialog() | Out-Null
    }

    function Show-DocumentationWindow {
        $html = @"
<html>
<head>
    <meta http-equiv='X-UA-Compatible' content='IE=Edge' />
    <style>
        body { font-family: 'Segoe UI', sans-serif; margin: 8px; background-color: #0f172a; color: #e2e8f0; }
        h1 { font-size: 20px; color: #93c5fd; }
        h2 { font-size: 16px; margin-top: 18px; color: #7dd3fc; }
        ul { margin-left: 18px; }
        a { color: #93c5fd; }
        code { background: #1e293b; padding: 2px 4px; border-radius: 3px; }
    </style>
</head>
<body>
    <h1><b><i>Scriptify</i></b> (ODT Utility) v3.5.0</h1>
    <p>Scriptify ODT Utility is a WPF front end for the Microsoft Office Deployment Tool. It downloads the Office Deployment Tool binary, generates a configuration script based on user parameters and applies the script to the ODT.</p>
    <h2>Generating Configurations</h2>
    <ul>
        <li>Select your Office product and edition; retail-only items are shown by default, and the Settings menu lets you toggle volume editions.</li>
        <li>Pick a servicing channel and choose whether to run the latest build or enter a specific version.</li>
        <li>Enter a product key if required. Use the Settings menu to add languages, switch to 32-bit installation, or force the Display level to None.</li>
        <li>Click <strong>Download</strong> to pick a local folder for the bits; the chosen path is recorded as <code>SourcePath</code> inside the config.</li>
        <li>Use <strong>Install</strong> to run the same configuration through the Microsoft Office Deployment Tool <code>setup.exe</code> that lives under <code>%AppData%\Microsoft\ODT</code> (it is downloaded and extracted automatically when the GUI launches).</li>
    </ul>
    <h2>Settings & Defaults</h2>
    <ul>
        <li><strong>File → New/Open/Save/Save As/Exit</strong> manages configuration state, persistence, and the View Script preview.</li>
        <li><strong>Tools → Set Defaults/Clear Defaults</strong> stores or clears <code>odt-defaults.json</code>; defaults include architecture, channel, language selection, force-app-close, remove-office, and application selections.</li>
        <li><strong>Settings → Add Languages</strong> adds locales, and <strong>Select Applications</strong> toggles which Office apps are included—unchecked apps emit <code>&lt;ExcludeApp ID="AppName"/&gt;</code>.</li>
        <li>Remaining Settings toggles cover volume editions, restore-point creation, MSI removal, force-app-close, and the documentation/About/Licensing helpers.</li>
        <li>Logs live under the <code>logs</code> folder beside the script, and the View menu opens each log with a responsive grid.</li>
    </ul>
    <h2>Servicing Channels</h2>
    <ul>
        <li><strong>Current</strong> – Latest retail builds with automatic feature updates.</li>
        <li><strong>MonthlyEnterprise</strong> – Monthly targeted enterprise release cadence.</li>
        <li><strong>SemiAnnual</strong> – Broad enterprise semi-annual release.</li>
        <li><strong>SemiAnnualPreview</strong> – Preview builds for the Semi-Annual channel.</li>
        <li><strong>BetaChannel</strong> – Insider Fast/Beta ring for early testing.</li>
        <li><strong>CurrentPreview</strong> – Preview stream for the Current channel.</li>
        <li><strong>PerpetualVL2021</strong> – Perpetual Volume Licensing for Office 2021.</li>
        <li><strong>PerpetualVL2019</strong> – Perpetual Volume Licensing for Office 2019.</li>
    </ul>
    <h2>Repairs</h2>
    <ul>
        <li>The Repair menu launches Quick or Full repairs via <code>OfficeClickToRun.exe</code>.</li>
        <li>The PST repair window lets you browse to a PST file and starts <code>SCANPST.EXE</code> once the animated progress completes.</li>
    </ul>
</body>
</html>
"@

        [xml]$docXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Documents" Height="520" Width="620" Background="#0F172A" Foreground="White"
        FontFamily="Segoe UI" WindowStartupLocation="CenterOwner">
    <Grid Margin="8">
        <WebBrowser Name="DocBrowser" />
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $docXaml
        $docWindow = [Windows.Markup.XamlReader]::Load($reader)
        $docWindow.Owner = $window
        $browser = $docWindow.FindName('DocBrowser')
        $browser.NavigateToString($html)
        $docWindow.ShowDialog() | Out-Null
    }

    function Show-AboutWindow {
        [xml]$aboutXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="About Cybtek STK" Height="180" Width="360"
        Background="#111827" Foreground="White" FontFamily="Segoe UI"
        WindowStartupLocation="CenterOwner" ResizeMode="NoResize">
    <Grid Margin="16">
        <StackPanel>
            <TextBlock Text="Scriptify (ODT Utility)" FontSize="20" FontWeight="Bold" FontStyle="Italic"
                       FontFamily="Space Grotesk Medium" />
            <TextBlock Text="Version 3.5.0" Margin="0,6,0,0" />
            <TextBlock Text=" " Margin="0,6,0,0" />
            <TextBlock Text="A CybtekSTK Software Application." Margin="0,6,0,0" FontFamily="Space Grotesk" />
            <TextBlock Margin="0,6,0,0">
                <Hyperlink x:Name="ScriptifyLink"
                           NavigateUri="https://github.com/s0nt3k/CybtekSTK-Scriptify"
                           FontFamily="Space Grotesk"
                           Foreground="#3A88F6"
                           TextDecorations="Underline">https://github.com/s0nt3k/CybtekSTK-Scriptify</Hyperlink>
            </TextBlock>
            
        </StackPanel>
    </Grid>
</Window>
"@
        $reader = New-Object System.Xml.XmlNodeReader $aboutXaml
        $aboutWindow = [Windows.Markup.XamlReader]::Load($reader)
        $aboutWindow.Owner = $window
        $scriptifyLink = $aboutWindow.FindName('ScriptifyLink')
        if ($scriptifyLink) {
            $scriptifyLink.Add_RequestNavigate({
                param($eventSender,$navArgs)
                try {
                    Start-Process $navArgs.Uri.AbsoluteUri
                } catch {
                    Show-Error "Failed to open link: $_"
                }
                $navArgs.Handled = $true
            })
        }
        $aboutWindow.ShowDialog() | Out-Null
    }
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Scriptify v3.5.0" Width="408" Height="620"
        WindowStartupLocation="CenterScreen" Background="#0B1221" Foreground="#F8FAFC"
        FontFamily="Space Grotesk" FontSize="14">
    <Window.Resources>
        <Style x:Key="LightComboBoxItemStyle" TargetType="ComboBoxItem">
            <Setter Property="Foreground" Value="#000000" />
            <Setter Property="Background" Value="#FFFFFF" />
            <Setter Property="Padding" Value="6,2" />
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="#E2E8F0" />
                    <Setter Property="Foreground" Value="#000000" />
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#CBD5F5" />
                    <Setter Property="Foreground" Value="#000000" />
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Foreground" Value="#9CA3AF" />
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style x:Key="LightComboBoxStyle" TargetType="ComboBox">
            <Setter Property="Foreground" Value="#000000" />
            <Setter Property="Background" Value="#FFFFFF" />
            <Setter Property="BorderBrush" Value="#CBD5F5" />
            <Setter Property="ItemContainerStyle" Value="{StaticResource LightComboBoxItemStyle}" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
        </Style>
        <Style x:Key="LightTextBoxStyle" TargetType="TextBox">
            <Setter Property="Foreground" Value="#0F172A" />
            <Setter Property="Background" Value="#FFFFFF" />
            <Setter Property="BorderBrush" Value="#CBD5F5" />
            <Setter Property="VerticalContentAlignment" Value="Center" />
        </Style>
        <DataTemplate x:Key="ChannelItemTemplate">
            <TextBlock Text="{Binding Label}" Foreground="#000000" />
        </DataTemplate>
        <DataTemplate x:Key="VersionModeTemplate">
            <TextBlock Text="{Binding Label}" Foreground="#000000" />
        </DataTemplate>
    </Window.Resources>
    <DockPanel>
        <Menu DockPanel.Dock="Top" Background="#F8FAFC" Foreground="#0F172A">
            <Menu.Resources>
                <Style TargetType="MenuItem">
                    <Setter Property="Foreground" Value="#0F172A" />
                    <Setter Property="Background" Value="#F8FAFC" />
                </Style>
            </Menu.Resources>
            <MenuItem Header="_File">
                <MenuItem Header="New" Name="FileNewMenu" />
                <MenuItem Header="Open" Name="FileOpenMenu" />
                <MenuItem Header="Save" Name="FileSaveMenu" />
                <MenuItem Header="Save As" Name="FileSaveAsMenu" />
                <Separator />
                <MenuItem Header="View Script" Name="FileViewScriptMenu" />
                <MenuItem Header="Exit" Name="FileExitMenu" />
            </MenuItem>
            <MenuItem Header="_View">
                <MenuItem Header="Download Log" Name="ViewDownloadLogMenu" />
                <MenuItem Header="Install Log" Name="ViewInstallLogMenu" />
            </MenuItem>
            <MenuItem Header="_Repair">
                <MenuItem Header="Quick Repair" Name="QuickRepairMenu" />
                <MenuItem Header="Full Repair" Name="FullRepairMenu" />
                <Separator />
                <MenuItem Header="PST Repair..." Name="PstRepairMenu" />
            </MenuItem>
            <MenuItem Header="_Tools">
                <MenuItem Header="Set Defaults" Name="SetDefaultsMenu" />
                <MenuItem Header="Clear Defaults" Name="ClearDefaultsMenu" />
                <MenuItem Header="Remove Office" Name="RemovePreviousInstallsMenu" />
            </MenuItem>
        <MenuItem Header="_Settings">
        <MenuItem Header="Add Languages..." Name="SettingsAddLanguagesMenu" />
        <MenuItem Header="Select Applications..." Name="SettingsSelectAppsMenu" />
        <MenuItem Header="Show Volume Editions" Name="ShowVolumeEditionsMenu" IsCheckable="True" />
        <MenuItem Header="Disable Restore Point" Name="DisableRestorePointMenu" IsCheckable="True" />
        <MenuItem Header="Disable Remove Office" Name="DisableRemoveOfficeMenu" IsCheckable="True" />
        <MenuItem Header="Install Microsoft Project" Name="InstallProjectMenu" IsCheckable="True" />
        <MenuItem Header="Install Microsoft Visio" Name="InstallVisioMenu" IsCheckable="True" />
        <MenuItem Header="Force App Close" Name="ForceAppCloseMenu" IsCheckable="True" IsChecked="True" />
        <Separator />
        <MenuItem Header="Install 32-Bit" Name="Install32BitMenu" IsCheckable="True" />
        <MenuItem Header="Display Level (None)" Name="DisplayNoneMenu" IsCheckable="True" />
        </MenuItem>
            <MenuItem Header="_Help">
                <MenuItem Header="Documents" Name="DocsMenu" />
                <MenuItem Header="License" Name="LicenseMenu" />
                <MenuItem Header="About" Name="AboutMenu" />
            </MenuItem>
        </Menu>
        <Grid Margin="24">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*" />
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="*" />
                <RowDefinition Height="Auto" />
            </Grid.RowDefinitions>

            <Border Padding="10" CornerRadius="12" Background="#131C32" BorderBrush="#1F2A46" BorderThickness="1">
                <StackPanel>
                <TextBlock Text="Scriptify!" FontSize="28" FontWeight="Bold" FontStyle="Italic" FontFamily="Space Grotesk Medium" />
                    <TextBlock Text="A CybtekSTK Software Application." Margin="0,6,0,0" Foreground="#94A3B8" FontFamily="Space Grotesk" />
                </StackPanel>
            </Border>

            <StackPanel Grid.Row="1" Grid.Column="0" Margin="0,20,12,0">
                <TextBlock Text="Office Product" />
                <ComboBox Name="ProductCombo" Height="34" Width="360" Margin="0,6,0,12" Style="{StaticResource LightComboBoxStyle}" />

                <TextBlock Text="Office Edition" />
                <ComboBox Name="EditionCombo" Height="34" Width="360" Margin="0,6,0,12" Style="{StaticResource LightComboBoxStyle}" />

                <TextBlock Text="Servicing Channel" />
                <ComboBox Name="ChannelCombo" Height="34" Width="360" Margin="0,6,0,12" Style="{StaticResource LightComboBoxStyle}" ItemTemplate="{StaticResource ChannelItemTemplate}" />

                <TextBlock Text="Version Selection" />
                <StackPanel Orientation="Horizontal" Margin="0,6,0,4" HorizontalAlignment="Stretch">
                    <ComboBox Name="VersionModeCombo" Width="180" Style="{StaticResource LightComboBoxStyle}" ItemTemplate="{StaticResource VersionModeTemplate}" />
                    <TextBox Name="VersionText" Height="34" Margin="12,0,0,0" Width="180" Style="{StaticResource LightTextBoxStyle}" />
                </StackPanel>

                <TextBlock Text="Product Key (optional)" />
                <TextBox Name="ProductKeyText" Height="34" Width="360" Margin="0,6,0,0" Style="{StaticResource LightTextBoxStyle}" />
            </StackPanel>

            <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,40,0,0">
                <Button Name="DownloadButton" Content="Download" Width="160" Height="44" Margin="0,0,18,0" Background="#2563EB" />
                <Button Name="InstallButton" Content="Install" Width="160" Height="44" Background="#22C55E" />
            </StackPanel>
        </Grid>
    </DockPanel>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    $controls = @{
        ProductCombo = $window.FindName('ProductCombo')
        EditionCombo = $window.FindName('EditionCombo')
        ChannelCombo = $window.FindName('ChannelCombo')
        VersionModeCombo = $window.FindName('VersionModeCombo')
        VersionText = $window.FindName('VersionText')
        ProductKeyText = $window.FindName('ProductKeyText')
        DownloadButton = $window.FindName('DownloadButton')
        InstallButton = $window.FindName('InstallButton')
        FileNewMenu = $window.FindName('FileNewMenu')
        FileOpenMenu = $window.FindName('FileOpenMenu')
        FileSaveMenu = $window.FindName('FileSaveMenu')
        FileSaveAsMenu = $window.FindName('FileSaveAsMenu')
        FileViewScriptMenu = $window.FindName('FileViewScriptMenu')
        FileExitMenu = $window.FindName('FileExitMenu')
        ViewDownloadLogMenu = $window.FindName('ViewDownloadLogMenu')
        ViewInstallLogMenu = $window.FindName('ViewInstallLogMenu')
        QuickRepairMenu = $window.FindName('QuickRepairMenu')
        FullRepairMenu = $window.FindName('FullRepairMenu')
        PstRepairMenu = $window.FindName('PstRepairMenu')
        SetDefaultsMenu = $window.FindName('SetDefaultsMenu')
        ClearDefaultsMenu = $window.FindName('ClearDefaultsMenu')
        SettingsAddLanguagesMenu = $window.FindName('SettingsAddLanguagesMenu')
        Install32BitMenu = $window.FindName('Install32BitMenu')
        DisplayNoneMenu = $window.FindName('DisplayNoneMenu')
        ShowVolumeEditionsMenu = $window.FindName('ShowVolumeEditionsMenu')
        DisableRestorePointMenu = $window.FindName('DisableRestorePointMenu')
        DisableRemoveOfficeMenu = $window.FindName('DisableRemoveOfficeMenu')
        ForceAppCloseMenu = $window.FindName('ForceAppCloseMenu')
        InstallProjectMenu = $window.FindName('InstallProjectMenu')
        InstallVisioMenu = $window.FindName('InstallVisioMenu')
        SettingsSelectAppsMenu = $window.FindName('SettingsSelectAppsMenu')
        RemovePreviousInstallsMenu = $window.FindName('RemovePreviousInstallsMenu')
        DocsMenu = $window.FindName('DocsMenu')
        LicenseMenu = $window.FindName('LicenseMenu')
        AboutMenu = $window.FindName('AboutMenu')
    }
    # Populate controls
    $controls.ProductCombo.ItemsSource = $officeCatalog
    $controls.ProductCombo.DisplayMemberPath = 'Name'
    $controls.ProductCombo.SelectedIndex = 0
    Update-EditionOptions
    $controls.EditionCombo.DisplayMemberPath = 'Name'
    $controls.EditionCombo.IsEnabled = $false

    $controls.ChannelCombo.ItemsSource = $channelOptions
    $controls.ChannelCombo.SelectedIndex = 0

    $controls.VersionModeCombo.ItemsSource = $versionModes
    $controls.VersionModeCombo.SelectedIndex = 0
    Set-VersionTextState

    # Wire events
    $controls.ProductCombo.Add_SelectionChanged({ Update-EditionOptions })
    $controls.VersionModeCombo.Add_SelectionChanged({ Set-VersionTextState })
    $controls.DownloadButton.Add_Click({ Invoke-ODTAction -Action 'Download' })
    $controls.InstallButton.Add_Click({ Invoke-ODTAction -Action 'Install' })
    $controls.ViewDownloadLogMenu.Add_Click({
            $log = Join-Path $script:LogsPath 'download.json'
            Show-LogWindow -Title 'Download Log' -LogFile $log
        })
    $controls.ViewInstallLogMenu.Add_Click({
            $log = Join-Path $script:LogsPath 'install.json'
            Show-LogWindow -Title 'Install Log' -LogFile $log
        })
    $controls.QuickRepairMenu.Add_Click({ Invoke-Repair -Mode 'QuickRepair' })
    $controls.FullRepairMenu.Add_Click({ Invoke-Repair -Mode 'FullRepair' })
    $controls.PstRepairMenu.Add_Click({ Show-PstRepairWindow })
    $controls.SetDefaultsMenu.Add_Click({ Save-CurrentDefaults })
    $controls.ClearDefaultsMenu.Add_Click({ Clear-Defaults })
    $controls.SettingsAddLanguagesMenu.Add_Click({ Show-LanguagePicker })
    $controls.SettingsSelectAppsMenu.Add_Click({ Show-AppSelectionWindow })
    $controls.ShowVolumeEditionsMenu.Add_Click({
            $script:ShowVolumeEditions = $controls.ShowVolumeEditionsMenu.IsChecked
            Update-EditionOptions -PreserveSelection $true
            Update-SettingsMenuState
        })
    $controls.DisableRestorePointMenu.Add_Click({
            $script:DisableRestorePoint = $controls.DisableRestorePointMenu.IsChecked
            Update-SettingsMenuState
        })
    $controls.DisableRemoveOfficeMenu.Add_Click({
            $script:DisableRemoveOffice = $controls.DisableRemoveOfficeMenu.IsChecked
            Update-SettingsMenuState
        })
    $controls.ForceAppCloseMenu.Add_Click({
            $script:ForceAppClose = $controls.ForceAppCloseMenu.IsChecked
            Update-SettingsMenuState
        })
     $controls.InstallProjectMenu.Add_Click({
            $script:InstallProject = $controls.InstallProjectMenu.IsChecked
            Update-SettingsMenuState
        })
     $controls.InstallVisioMenu.Add_Click({
            $script:InstallVisio = $controls.InstallVisioMenu.IsChecked
            Update-SettingsMenuState
        })
    $controls.Install32BitMenu.Add_Click({
            $script:Is32BitArchitecture = $controls.Install32BitMenu.IsChecked
            Update-SettingsMenuState
        })
    $controls.DisplayNoneMenu.Add_Click({
            $script:IsDisplayNone = $controls.DisplayNoneMenu.IsChecked
            Update-SettingsMenuState
        })
    $controls.DocsMenu.Add_Click({ Show-DocumentationWindow })
    $controls.AboutMenu.Add_Click({ Show-AboutWindow })
    $controls.LicenseMenu.Add_Click({
            $productKey = if ($controls.ProductKeyText) { $controls.ProductKeyText.Text.Trim().ToUpper() } else { '' }
            if ($productKey -eq 'RESETLIC') {
                Clear-LicenseData | Out-Null
            }
            Show-LicenseWindow -Owner $window
        })
    $controls.FileNewMenu.Add_Click({ Reset-UIToDefaults })
    $controls.FileOpenMenu.Add_Click({ Show-GeneratedScriptsWindow })
    $controls.FileSaveMenu.Add_Click({ Save-CurrentConfiguration })
    $controls.FileSaveAsMenu.Add_Click({ Show-SaveDefaultsAsDialog })
    $controls.FileViewScriptMenu.Add_Click({
            $state = Get-UIState
            Show-CurrentScriptWindow -State $state
        })
    $controls.FileExitMenu.Add_Click({ $window.Close() })
    $controls.RemovePreviousInstallsMenu.Add_Click({ Show-RemovePreviousInstallsWindow })

    # Apply saved defaults if available
    $defaults = Get-DefaultsFile
    if ($defaults) {
        Set-UiState -state $defaults
    }
    Update-SettingsMenuState

    $window.ShowDialog() | Out-Null

}

if ($MyInvocation.InvocationName -ne '.') {
    Show-CybtekODTGui
}
