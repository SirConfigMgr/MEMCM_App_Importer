<#	
	.NOTES
	===========================================================================
	 Created on:   	2021.07.28
	 Last Updated:  2021.09.16
         Version:       1.1
	 Author:	Rene Hartmann
	 Filename:     	MEMCM_App_Importer.ps1
	===========================================================================
	.DESCRIPTION
	Create MEMCM Applications (optional with PSADT).

#>

#region # Initializing
### Initial Configuration
$Version = "1.1"
$User = [Environment]::UserName
$Path = "$PSScriptRoot"
$LogFolder = "$Path\Logs"
If (!(Test-Path $LogFolder)) {New-Item -Path $LogFolder -ItemType Directory -Force}
New-PSDrive -Name LogPath -PSProvider FileSystem -Root $LogFolder
$LogPath = "LogPath:\\Importer_$User.log"
$AssemblyLocation = "$Path\bin"
$pathPanel= split-path -parent $MyInvocation.MyCommand.Definition
$Date = Get-Date -Format "dd\/MM\/yyy"

### Load Config
If (Test-Path "$Path\Config.cfg") {Get-Content "$Path\Config.cfg" | foreach-object -begin {$Config=@{}} -process { $ConfigValues = [regex]::split($_,'='); if(($ConfigValues[0].CompareTo("") -ne 0) -and ($ConfigValues[0].StartsWith("[") -ne $True)) { $Config.Add($ConfigValues[0], $ConfigValues[1]) } }}

### Load WPF Framework
[System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')  | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')   | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework')  | out-null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsBase')    | out-null
[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore')   | out-null
Foreach ($Assembly in (Dir $AssemblyLocation -Filter *.dll)) {
     [System.Reflection.Assembly]::LoadFrom($Assembly.fullName) | out-null
     }
#endregion # Initializing

#region # Functions
Function Write-Log {

[CmdletBinding()]
Param(
    [parameter(Mandatory=$true)][String]$LogPath,
    [parameter(Mandatory=$true)][String]$Message,
    [parameter(Mandatory=$true)][String]$Component,
    [Parameter(Mandatory=$true)][ValidateSet("Info", "Warning", "Error")][String]$Type
    )

Switch ($Type) {
    "Info" {[int]$Type = 1}
    "Warning" {[int]$Type = 2}
    "Error" {[int]$Type = 3}
    }

$Content = "<![LOG[$Message]LOG]!>" +`
        "<time=`"$(Get-Date -Format "HH:mm:ss.ffffff")`" " +`
        "date=`"$(Get-Date -Format "M-d-yyyy")`" " +`
        "component=`"$Component`" " +`
        "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " +`
        "type=`"$Type`" " +`
        "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " +`
        "file=`"`">"

Add-Content -Path $LogPath -Value $Content
}

Function Connect-MEMCM {

Param(
    [String]$SiteCode,
    [String]$ProviderMachineName
    )

    $Info = "Enter Function Connect-MEMCM"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "MEMCM_Connect" -Type Info

    if((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" 
        }

    if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
        }

    Set-Location "$($SiteCode):\"

    $DPGroups = (Get-CMDistributionPointGroup).Name
    Foreach ($DPGroup in $DPGroups) {$ComboBox_DistributionPointGroup.AddText($DPGroup)}
        }

Function Read-Installer {
    $Info = "Enter Function Read-Installer"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Read_Installer" -Type Info
    [void]
    [System.Reflection.Assembly]::LoadWithPartialName
    ("System.Windows.Forms")
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.filter = "MSI files (*.msi)| *.msi|EXE files (*.exe)| *.exe"
    $OpenFileDialog.ShowDialog()
    $FileFolder = Split-Path $OpenFileDialog.filename -Parent
    $FileName = Split-Path $OpenFileDialog.filename -Leaf

    If ($FileName) {
        $Info = "Filename: $FileName"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Read_Installer" -Type Info

        If ($FileName -like "*.msi") {
            [IO.FileInfo]$Path = $OpenFileDialog.filename
            $WindowsInstaller = New-Object -com WindowsInstaller.Installer
            $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase","InvokeMethod",$Null,$WindowsInstaller,@($Path.FullName,0))
            $View = $MSIDatabase.GetType().InvokeMember("OpenView","InvokeMethod",$null,$MSIDatabase,"SELECT * FROM Property")

            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
            $MSIProperties = while($Record = $View.GetType().InvokeMember("Fetch","InvokeMethod",$null,$View,$null)){@{$Record.GetType().InvokeMember("StringData","GetProperty",$null,$Record,1) = $Record.GetType().InvokeMember("StringData","GetProperty",$null,$Record,2)}}
            $ComboBox_Installtype.SelectedIndex = 0
            $TextBox_Installfile.AppendText($FileName)
            $Global:MSIProductcode = $MSIProperties.ProductCode
            $TextBox_Vendor.AppendText($MSIProperties.Manufacturer)
            $TextBox_Name.AppendText($MSIProperties.ProductName)
            $TextBox_Version.AppendText($MSIProperties.ProductVersion)
            $TextBox_Sourcefolder.AppendText($FileFolder)
            }

        If ($FileName -like "*.exe") {
            $EXEProperties = Get-ChildItem $OpenFileDialog.filename | % {$_.VersionInfo} | Select *
            $ComboBox_Installtype.SelectedIndex = 1
            $TextBox_Installfile.AppendText($FileName)
            $TextBox_Vendor.AppendText($EXEProperties.CompanyName)
            $TextBox_Name.AppendText($EXEProperties.ProductName)
            $TextBox_Version.AppendText($EXEProperties.ProductVersion)
            $TextBox_Sourcefolder.AppendText($FileFolder)
            }
        }
    }

Function Generate-Package {

Param(
    [String]$DestinationPath,
    [String]$SourcePath,
    [String]$Vendor,
    [String]$Name,
    [String]$Version,
    [String]$Language,
    [String]$Architecture,
    [String]$Revision,
    [String]$User,
    [String]$Installtype,
    [String]$Installfile,
    [String]$Installparameter,
    [String]$Uninstalltype,
    [String]$UninstallNameOrCode,
    [String]$UninstallParameter,
    [String]$Message
    )

    $Info = "Enter Function Generate-Package"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
    $Info = "Check Destination Path"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info

    If (Test-Path "filesystem::$DestinationPath") {
        $Info = "--> Destination Path Exist"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
        $Info = "Checking Values"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info

        # Check Mandatory Values
        If (!($Vendor)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Vendor Field"
            Return
            }
        ElseIf (!($Name)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Name Field"
            Return
            }
        ElseIf (!($Version)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Version Field"
            Return
            }
        ElseIf (!($Language)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Language Field"
            Return
            }
        ElseIf (!($Architecture)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Architecture Box"
            Return
            }
        ElseIf (!($Revision)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Revision Field"
            Return
            }
        ElseIf (!($Installtype)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Installtype Box"
            Return
            }
        ElseIf (!($Installfile)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Installfile Field"
            Return
            }
        ElseIf (!($Uninstalltype)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Uninstalltype Box"
            Return
            }
        ElseIf (!($UninstallNameOrCode)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Uninstall Name Or Code Field"
            Return
            }
        ElseIf (!($ComboBox_Detection)) {
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Please Check Detection"
            Return
            }
        $Info = "--> Values Correct"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info

        # Generate Folder Name
        $Info = "Generate Folder Name"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
        $NewAppFolderName =$TextBox_NamingAppFolder.Text -replace([regex]::Escape("[Vendor]"),$Vendor) `
            -replace ([regex]::Escape("[Name]"),$Name) `
            -replace ([regex]::Escape("[Version]"),$Version) `
            -replace ([regex]::Escape("[Language]"),$Language) `
            -replace ([regex]::Escape("[Architecture]"),$Architecture) `
            -replace ([regex]::Escape("[Revision]"),$Revision)
        $Info = "--> Folder Name: $NewAppFolderName"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info

        # Create Application Folder
        $Info = "Check Application Folder"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
        $Global:NewAppDestinationPath = $DestinationPath + "\" + $NewAppFolderName
        Write-Host $NewAppDestinationPath
        If (Test-Path "filesystem::$NewAppDestinationPath") {
            $Info = "--> Application Folder Already Exist"
            Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Error
            $Window_Package_Generation.IsOpen = $true
            $Label_PG_ChildWindow.Content = "Application Folder Already Exist - Please Check Folder"
            Return
            }
        Else {
            $Info = "--> Application Folder Does Not Exist"
            Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
            New-Item -Path "filesystem::$NewAppDestinationPath" -ItemType directory -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
            If ($ErrorAction) {
                Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Generate_Package" -Type Error
                $Window_Package_Generation.IsOpen = $true
                $Label_PG_ChildWindow.Content = "Cannot Create Folder"
                }
            Else {
                $Info = "--> Folder Created"
                Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
                }
            }
        
        # Copy Files
        $Info = "Copy Files"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
        If ($CheckBox_PSADT.IsChecked -eq $true) {
            $Info = "--> Use PSADT - CheckBox Checked"
            Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
            
            # Copy PSADT
            Get-ChildItem -Path "$Path\PSADT\" | Copy-Item -Destination "filesystem::$NewAppDestinationPath" -Recurse -Container -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
            If ($ErrorAction) {
                Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Generate_Package" -Type Error
                $Window_Package_Generation.IsOpen = $true
                $Label_PG_ChildWindow.Content = "Cannot Copy PSADT Files"
                }
            Else {
                $Info = "--> PSADT Files Copied"
                Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
                }

            # Copy Install Files
            Get-ChildItem -Path "$SourcePath" | Copy-Item -Destination "filesystem::$NewAppDestinationPath\Files" -Recurse -Container -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
            If ($ErrorAction) {
                Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Generate_Package" -Type Error
                $Window_Package_Generation.IsOpen = $true
                $Label_PG_ChildWindow.Content = "Cannot Copy Install Files"
                }
            Else {
                $Info = "--> Install Files Copied"
                Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
                }

            # Copy App Icon
            If ($Image_Icon) {
                If (!(Test-Path "filesystem::$NewAppDestinationPath\Icon")) {New-Item -Path "filesystem::$NewAppDestinationPath\Icon" -ItemType directory -Force} 
                [String]$IconPath = $Image_Icon.Source
                Copy-Item "$PSScriptRoot\$IconPath" -Destination "filesystem::$NewAppDestinationPath\Icon" -Force -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
                If ($ErrorAction) {
                    Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Generate_Package" -Type Error
                    $Window_Package_Generation.IsOpen = $true
                    $Label_PG_ChildWindow.Content = "Cannot Copy Icon"
                    }
                Else {
                    $Info = "--> Icon Copied"
                    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
                    }
                }

            # Manipulate Depoloy-Application.ps1
            $Info = "Edit Depoloy-Application.ps1"
            Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
            If ($Installtype -eq "MSI") {$ExecutionString = "Execute-MSI -Action 'Install' -Path '$Installfile' $Installparameter"}
            Else {$ExecutionString = "Execute-Process -Path '$Installfile' $Installparameter"}
            If ($Uninstalltype -eq "Product Name") {$UninstallString = "Remove-MSIApplications -Name '$UninstallNameOrCode' $Uninstallparameter"}
            Elseif ($Uninstalltype -eq "Product Code") {$UninstallString = "Execute-MSI -Action 'Uninstall' -Path '$UninstallNameOrCode' $Uninstallparameter"}
            Elseif ($Uninstalltype -eq "Script") {$UninstallString = "Execute-Process -Path '$UninstallNameOrCode' $Uninstallparameter"}
            If ($Message) {$MessageString = "Show-InstallationPrompt -Message '$Message' -ButtonRightText 'OK' -Icon Information -NoWait"}
                Else {$MessageString = ""}
            Get-Content -Path "$Path\PSADT_Custom\Deploy-Application.ps1" | Foreach-Object {
                $_ -replace "<Vendor>", "[string]`$appVendor = '$Vendor'" `
                -replace "<Name>", "[string]`$appName = '$Name'" `
                -replace "<Version>", "[string]`$appVersion = '$Version'" `
                -replace "<Arch>", "[string]`$appArch = '$Architecture'" `
                -replace "<Lang>", "[string]`$appLang = '$Language'" `
                -replace "<Revision>", "[string]`$appRevision = '$Revision'" `
                -replace "<ScriptVersion>", "[string]`$appScriptVersion = '1.0'" `
                -replace "<Date>", "[string]`$appScriptDate = '$Date'" `
                -replace "<Creator>", "[string]`$appScriptVersion = '$User'" `
                -replace "<InstallCmdline>", "$ExecutionString" `
                -replace "<UninstallCmdline>", "$UninstallString" `
                -replace "<Message>", "$MessageString"
                } | Set-Content "filesystem::$NewAppDestinationPath\Deploy-Application.ps1" -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
                
                If ($ErrorAction) {
                    Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Generate_Package" -Type Error
                    $Window_Package_Generation.IsOpen = $true
                    $Label_PG_ChildWindow.Content = "Edit Depoloy-Application.ps1 failed - View Log"
                    }
                Else {
                    $Info = "--> Depoloy-Application.ps1 Successfully Edited"
                    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
                    }           
            
            }
        Else {
            $Info = "--> Copy Files Without PSADT - CheckBox Not Checked"
            Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info

            # Copy Install Files
            Get-ChildItem -Path "$SourcePath" | Copy-Item -Destination "filesystem::$NewAppDestinationPath" -Recurse -Container -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
            If ($ErrorAction) {
                Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Generate_Package" -Type Error
                $Window_Package_Generation.IsOpen = $true
                $Label_PG_ChildWindow.Content = "Cannot Copy Install Files"
                }
            Else {
                $Info = "--> Install Files Copied"
                Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
                }
            
            # Copy App Icon
            If ($Image_Icon) {
                If (!(Test-Path "filesystem::$NewAppDestinationPath\Icon")) {New-Item -Path "filesystem::$NewAppDestinationPath\Icon" -ItemType directory -Force}
                [String]$IconPath = $Image_Icon.Source
                Copy-Item "$PSScriptRoot\$IconPath" -Destination "filesystem::$NewAppDestinationPath\Icon" -Force -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
                If ($ErrorAction) {
                    Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Generate_Package" -Type Error
                    $Window_Package_Generation.IsOpen = $true
                    $Label_PG_ChildWindow.Content = "Cannot Copy Icon"
                    }
                Else {
                    $Info = "--> Icon Copied"
                    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
                    }
                }
            }
        }
    Else {
        $Info = "--> Destination Path Does Not Exist"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Error
        $Window_Package_Generation.IsOpen = $true
        $Label_PG_ChildWindow.Content = "Destination Path Does Not Exist - Please Check Folder"
        Return
        }

    # Generate Package-Info-File
    $Info = "Generate Info-File"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
    "Vendor=" + $TextBox_Vendor.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Name=" + $TextBox_Name.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Version=" + $TextBox_Version.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Language=" + $TextBox_Language.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Architecture=" + $ComboBox_Architecture.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Revision=" + $TextBox_Revision.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    If ($Image_Icon) {"Icon=" + ([String]($Image_Icon.Source)).split("\")[1] | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append}
    "Installtype=" + $ComboBox_Installtype.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Installfile=" + $TextBox_Installfile.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Installparameter=" + $TextBox_Installparameter.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Detection=" + $ComboBox_Detection.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Uninstalltype=" + $ComboBox_Uninstalltype.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "UninstallNameOrCode=" + $TextBox_UninstallNameOrCode.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Uninstallparameter=" + $TextBox_Uninstallparameter.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Sourcefolder=" + $TextBox_Sourcefolder.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "Message=" + $TextBox_Message.Text | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    "DestinationPath=" + $NewAppDestinationPath | Out-File -FilePath "$NewAppDestinationPath\Pkg.info" -Encoding utf8 -Append
    If ($TextBox_DetectionScript_ChildWindow.Text) {$TextBox_DetectionScript_ChildWindow.Text | Out-File -FilePath "$NewAppDestinationPath\DetectionScript.ps1" -Encoding utf8 -Append}
    If (Test-Path "filesystem::$NewAppDestinationPath\Pkg.info") {
        $Info = "--> Info-File Generated"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
        }
    Else {
        $Info = "--> Info-File Generation Failed"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Error
        $Window_Package_Generation.IsOpen = $true
        $Label_PG_ChildWindow.Content = "Package Creation Failed - View Logs"
        Return
        }

    $Info = "--> Package Successfully Created"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Generate_Package" -Type Info
    $Window_Package_Generation.IsOpen = $true
    $Label_PG_ChildWindow.Content = "Package Creation Successfully"
    }

Function Load-PkgInfo {
    $Info = "Enter Function Load-PkgInfo"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Load_PkgInfo" -Type Info
    [void]
    [System.Reflection.Assembly]::LoadWithPartialName
    ("System.Windows.Forms")
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.filter = "MEMCM App Importer files (*.info)| *.info"
    $OpenFileDialog.ShowDialog()
    $FileFolder = Split-Path $OpenFileDialog.filename -Parent
    $FileName = Split-Path $OpenFileDialog.filename -Leaf

    Get-Content $OpenFileDialog.filename | foreach-object -begin {$PkgInfo=@{}} -process { $PkgInfoValues = [regex]::split($_,'='); if(($PkgInfoValues[0].CompareTo("") -ne 0) -and ($PkgInfoValues[0].StartsWith("[") -ne $True)) { $PkgInfo.Add($PkgInfoValues[0], $PkgInfoValues[1]) } }

    If ($PkgInfo.Icon) {Copy-Item "$FileFolder\Icon\$($PkgInfo.Icon)" -Destination "$PSScriptRoot\images" -Force}
    $TextBox_Vendor.AppendText($PkgInfo.Vendor)
    $TextBox_Name.AppendText($PkgInfo.Name)
    $TextBox_Version.AppendText($PkgInfo.Version)
    $TextBox_Language.AppendText($PkgInfo.Language)
    $ComboBox_Architecture.Text= $PkgInfo.Architecture
    $TextBox_Revision.AppendText($PkgInfo.Revision)
    $ComboBox_Installtype.Text= $PkgInfo.Installtype
    $TextBox_Installfile.AppendText($PkgInfo.Installfile)
    $TextBox_Installparameter.AppendText($PkgInfo.Installparameter)
    $ComboBox_Detection.Text= $PkgInfo.Detection
    $ComboBox_Uninstalltype.Text= $PkgInfo.Uninstalltype
    $TextBox_UninstallNameOrCode.AppendText($PkgInfo.UninstallNameOrCode)
    $Global:MSIProductCode = $PkgInfo.UninstallNameOrCode
    $TextBox_Uninstallparameter.AppendText($PkgInfo.Uninstallparameter)
    $TextBox_Sourcefolder.AppendText($PkgInfo.Sourcefolder)
    $TextBox_Message.AppendText($PkgInfo.Message)
    $Image_Icon.Source="images\$($PkgInfo.Icon)"
    $Global:NewAppDestinationPath = $PkgInfo.DestinationPath
    If (Test-Path "filesystem::$FileFolder\DetectionScript.ps1") {
        $DetectionScript = Get-Content "filesystem::$FileFolder\DetectionScript.ps1"
        Foreach ($Line in $DetectionScript) {
            $TextBox_DetectionScript_ChildWindow.AppendText($Line)
            $TextBox_DetectionScript_ChildWindow.AppendText("`n`n")
            }
        }

    If ($FileName) {
        $Info = "Load-PkgInfo for Application $($PkgInfo.Name) $($PkgInfo.Version)"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Load_PkgInfo" -Type Info
        }
    }

Function Load-Icon {
    $Info = "Enter Function Load-Icon"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Load_Icon" -Type Info
    [void]
    [System.Reflection.Assembly]::LoadWithPartialName
    ("System.Windows.Forms")
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.filter = "Images|*.jpg;*.png;*.bmp"
    $OpenFileDialog.ShowDialog()
    $FileFolder = Split-Path $OpenFileDialog.filename -Parent
    $FileName = Split-Path $OpenFileDialog.filename -Leaf

    Add-Type -AssemblyName System.Drawing
    $Icon = New-Object System.Drawing.Bitmap $OpenFileDialog.filename
    If ($Icon.with -le "512" -and $Icon.height -le "512") {
        $Info = "Icon Size Ok - Copy To Image-Folder"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Load_Icon" -Type Info
        If (Test-Path $PSScriptRoot\Images\$FileName) {Remove-Item $PSScriptRoot\Images\$FileName -Force}
        Copy-Item -Path $OpenFileDialog.filename -Destination $PSScriptRoot\Images -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
            If ($ErrorAction) {
                Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Load_Icon" -Type Error
                $Window_Package_Generation.IsOpen = $true
                $Label_PG_ChildWindow.Content = "Cannot Copy Icon File To PSScriptRoot - View Log"
                }
            Else {
                $Image_Icon.Source="images\$FileName"
                $Info = "--> Icon File Copied To PSScriptRoot"
                Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Load_Icon" -Type Info
                }
        
        }
    Else {
        $Info = "Icon Greater Than 512x512"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Load_Icon" -Type Error
        $Window_Package_Generation.IsOpen = $true
        $Label_PG_ChildWindow.Content = "Please Check Icon Dimension - Max. 512x512"
        }
    }

Function Create-Application {

Param(
    [String]$DestinationPath,
    [String]$Vendor,
    [String]$Name,
    [String]$Version,
    [String]$Language,
    [String]$Architecture,
    [String]$Revision,
    [String]$User,
    [String]$Installtype,
    [String]$Installfile,
    [String]$Installparameter,
    [String]$Uninstalltype,
    [String]$UninstallNameOrCode,
    [String]$UninstallParameter,
    [String]$DetectionScript,
    [String]$DetectionKey,
    [String]$DetectionFile
    )

    $Info = "Enter Function Create-Application"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info

    # Generate Localized Application Name
    $Info = "Generate Localized Application Name"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info
    $LocalizedName =$TextBox_NamingLocalizedName.Text -replace([regex]::Escape("[Vendor]"),$Vendor) `
        -replace ([regex]::Escape("[Name]"),$Name) `
        -replace ([regex]::Escape("[Version]"),$Version) `
        -replace ([regex]::Escape("[Language]"),$Language) `
        -replace ([regex]::Escape("[Architecture]"),$Architecture) `
        -replace ([regex]::Escape("[Revision]"),$Revision)
    $Info = "--> Localized Application Name: $LocalizedName"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info

    # Generate Application Name
    $Info = "Generate Localized Application Name"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info
    $Global:AppName =$TextBox_NamingApp.Text -replace([regex]::Escape("[Vendor]"),$Vendor) `
        -replace ([regex]::Escape("[Name]"),$Name) `
        -replace ([regex]::Escape("[Version]"),$Version) `
        -replace ([regex]::Escape("[Language]"),$Language) `
        -replace ([regex]::Escape("[Architecture]"),$Architecture) `
        -replace ([regex]::Escape("[Revision]"),$Revision)
    $Info = "--> Application Name: $AppName"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info

    # Generate Deployment Type Name
    $Info = "Generate Deployment Type Name"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info
    $DeploymentTypeName =$TextBox_NamingDeploymentType.Text -replace([regex]::Escape("[Vendor]"),$Vendor) `
        -replace ([regex]::Escape("[Name]"),$Name) `
        -replace ([regex]::Escape("[Version]"),$Version) `
        -replace ([regex]::Escape("[Language]"),$Language) `
        -replace ([regex]::Escape("[Architecture]"),$Architecture) `
        -replace ([regex]::Escape("[Revision]"),$Revision)
    $Info = "--> Deployment Type Name: $DeploymentTypeName"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info

    $Info = "Create Application"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info

    [String]$ImagePath = $Image_Icon.Source
    If ($Image_Icon.Source) {New-CMApplication -Name $AppName -Publisher $Vendor -SoftwareVersion $Version -LocalizedName $LocalizedName -IconLocationFile "$PSScriptRoot\$ImagePath" -ErrorAction SilentlyContinue -ErrorVariable ErrorAction}
    Else {New-CMApplication -Name $AppName -Publisher $Vendor -SoftwareVersion $Version -LocalizedName $LocalizedName -ErrorAction SilentlyContinue -ErrorVariable ErrorAction}
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Application" -Type Error
        $Window_MEMCM_Connection.IsOpen = $true
        $Label_MEMCM_ChildWindow.Content = "Cannot Create Application - View Log File"
        $Global:AppCreated = $false
        Return
        }
    Else {
        $Info = "--> Application Created"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info
        }


        $Info = "Create Deployment Type"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info
        
        If ($ComboBox_Detection.Text -eq "Product Code") {
            If ($CheckBox_PSADT.IsChecked -eq $true) {Add-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -ContentLocation $DestinationPath -InstallCommand "Deploy-Application.exe" -UninstallCommand "Deploy-Application.exe -Uninstall" -InstallationBehaviorType $ComboBox_InstallBehaviorType.Text -LogonRequirementType $ComboBox_LogonRequirementType.Text -ProductCode $UninstallNameOrCode -UninstallOption SameAsInstall -ErrorAction SilentlyContinue -ErrorVariable ErrorAction}
            Else {
                If ($ComboBox_UninstallType.Text -eq "MSI") {$UninstallCommand = "msiexec /x $($MSIProductcode)"}
                If ($ComboBox_UninstallType.Text -eq "Script") {$UninstallCommand = "$($TextBox_UninstallFile.Text) $($TextBox_UninstallParameter.Text)"}
                If ($ComboBox_InstallType.Text -eq "MSI") {Add-CMMSIDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -ContentLocation "$($DestinationPath)\$($TextBox_Installfile.Text)" -InstallCommand "msiexec /i $($TextBox_InstallFile.Text) $($TextBox_InstallParameter.Text)" -UninstallCommand $UninstallCommand -InstallationBehaviorType $ComboBox_InstallBehaviorType.Text -LogonRequirementType $ComboBox_LogonRequirementType.Text -ProductCode $UninstallNameOrCode -UninstallOption SameAsInstall -ErrorAction SilentlyContinue -ErrorVariable ErrorAction -Force}
                If ($ComboBox_InstallType.Text -eq "Script") {Add-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -ContentLocation $DestinationPath -InstallCommand "$($TextBox_InstallFile.Text) $($TextBox_InstallParameter.Text)"  -UninstallCommand $UninstallCommand -InstallationBehaviorType $ComboBox_InstallBehaviorType.Text -LogonRequirementType $ComboBox_LogonRequirementType.Text -ProductCode $UninstallNameOrCode -UninstallOption SameAsInstall -ErrorAction SilentlyContinue -ErrorVariable ErrorAction}
                }
            If ($ErrorAction) {
                Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Application" -Type Error
                $Window_MEMCM_Connection.IsOpen = $true
                $Label_MEMCM_ChildWindow.Content = "Cannot Create Deployment Type - View Log File"
                $Global:AppCreated = $false
                Return
                }
            Else {
                $Info = "--> Deployment Type Created"
                Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info
                }
            }

        If ($ComboBox_Detection.Text -eq "File") {
            # Next Releases
            }

        If ($ComboBox_Detection.Text -eq "Reg Key") {
            # Next Releases
            }

        If ($ComboBox_Detection.Text -eq "PS Script") {
            If ($CheckBox_PSADT.IsChecked -eq $true) {Add-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -ContentLocation $DestinationPath -InstallCommand "Deploy-Application.exe" -UninstallCommand "Deploy-Application.exe -Uninstall" -InstallationBehaviorType $ComboBox_InstallBehaviorType.Text -LogonRequirementType $ComboBox_LogonRequirementType.Text -ScriptText $DetectionScript -ScriptLanguage PowerShell -UninstallOption SameAsInstall -ErrorAction SilentlyContinue -ErrorVariable ErrorAction}
            Else {
                If ($ComboBox_UninstallType.Text -eq "MSI") {$UninstallCommand = "msiexec /x $($MSIProductcode)"}
                If ($ComboBox_UninstallType.Text -eq "Script") {$UninstallCommand = "$($TextBox_UninstallFile.Text) $($TextBox_UninstallParameter.Text)"}
                If ($ComboBox_InstallType.Text -eq "MSI") {Add-CMMSIDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -ContentLocation "$($DestinationPath)\$($TextBox_Installfile.Text)" -InstallCommand "msiexec /i $($TextBox_InstallFile.Text) $($TextBox_InstallParameter.Text)" -UninstallCommand $UninstallCommand -InstallationBehaviorType $ComboBox_InstallBehaviorType.Text -LogonRequirementType $ComboBox_LogonRequirementType.Text -ScriptText $DetectionScript -ScriptLanguage PowerShell -UninstallOption SameAsInstall -ErrorAction SilentlyContinue -ErrorVariable ErrorAction -Force}
                If ($ComboBox_InstallType.Text -eq "Script") {Add-CMScriptDeploymentType -ApplicationName $AppName -DeploymentTypeName $DeploymentTypeName -ContentLocation $DestinationPath -InstallCommand "$($TextBox_InstallFile.Text) $($TextBox_InstallParameter.Text)"  -UninstallCommand $UninstallCommand -InstallationBehaviorType $ComboBox_InstallBehaviorType.Text -InstallationBehaviorType $ComboBox_InstallBehaviorType.Text -LogonRequirementType $ComboBox_LogonRequirementType.Text -ScriptText $DetectionScript -ScriptLanguage PowerShell -UninstallOption SameAsInstall -ErrorAction SilentlyContinue -ErrorVariable ErrorAction}
                }
            If ($ErrorAction) {
                Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Application" -Type Error
                $Window_MEMCM_Connection.IsOpen = $true
                $Label_MEMCM_ChildWindow.Content = "Cannot Create Deployment Type - View Log File"
                $Global:AppCreated = $false
                Return
                }
            Else {
                $Info = "--> Deployment Type Created"
                Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Application" -Type Info
                }
            }

    $Button_CreateApp.IsEnabled = $false
    $TextBox_CreateApp.Background = "Green"
    $TextBox_CreateApp.TextAlignment = "Center"
    $TextBox_CreateApp.Text = "Created"
    $Button_CollectionSettings.IsEnabled = "true"
    $Global:AppCreated = $true
    }

Function Create-Collections {

Param (
    [String]$Vendor,
    [String]$Name,
    [String]$Version,
    [String]$Language,
    [String]$Architecture
    )

    $Info = "Enter Function Create-Collections"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info

    # Generate Install Collection Name
    $Info = "Generate Install Collection Name"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
    $Global:CompleteAppNameInstCol =$TextBox_NamingCollectionInstall.Text -replace([regex]::Escape("[Vendor]"),$Vendor) `
        -replace ([regex]::Escape("[Name]"),$Name) `
        -replace ([regex]::Escape("[Version]"),$Version) `
        -replace ([regex]::Escape("[Language]"),$Language) `
        -replace ([regex]::Escape("[Architecture]"),$Architecture) `
        -replace ([regex]::Escape("[Revision]"),$Revision)
    $Info = "Install Collection Name: $CompleteAppNameInstCol"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info

    # Generate Uninstall Collection Name
    $Info = "Generate Uninstall Collection Name"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
    $Global:CompleteAppNameUninstCol =$TextBox_NamingCollectionUninstall.Text -replace([regex]::Escape("[Vendor]"),$Vendor) `
        -replace ([regex]::Escape("[Name]"),$Name) `
        -replace ([regex]::Escape("[Version]"),$Version) `
        -replace ([regex]::Escape("[Language]"),$Language) `
        -replace ([regex]::Escape("[Architecture]"),$Architecture) `
        -replace ([regex]::Escape("[Revision]"),$Revision)
    $Info = "Uninstall Collection Name: $CompleteAppNameUninstCol"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info

    # Create Collections
    $Info = "Create Install Collection"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
    $CollectionFolderPath = $TextBox_SiteCode.Text + ":\DeviceCollection\" + $TextBox_Folder.Text
    New-CMDeviceCollection -Name $CompleteAppNameInstCol -LimitingCollectionName $TextBox_LimitingCollection.Text -ErrorAction SilentlyContinue -ErrorVariable ErrorAction 
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Collections" -Type Error
        $TextBox_CreateCollections.Background = "Red"
        $TextBox_CreateCollections.TextAlignment = "Center"
        $TextBox_CreateCollections.Text = "Failed"
        $Global:CollectionsCreated = $false
        Return
        }
    Else {
        $Info = "--> Install Collection Created"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
        }

    $Info = "Create Uninstall Collection"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
    New-CMDeviceCollection -Name $CompleteAppNameUninstCol -LimitingCollectionName $TextBox_LimitingCollection.Text -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Collections" -Type Error
        $TextBox_CreateCollections.Background = "Red"
        $TextBox_CreateCollections.TextAlignment = "Center"
        $TextBox_CreateCollections.Text = "Failed"
        $Global:CollectionsCreated = $false
        Return
        }
    Else {
        $Info = "--> Uninstall Collection Created"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
        }

    # Move Collections
    $Info = "Move Install Collection"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
    Move-CMObject -InputObject (Get-CMDeviceCollection -Name $CompleteAppNameInstCol) -FolderPath $CollectionFolderPath -ErrorAction SilentlyContinue -ErrorVariable ErrorAction 
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Collections" -Type Error
        $TextBox_CreateCollections.Background = "Red"
        $TextBox_CreateCollections.TextAlignment = "Center"
        $TextBox_CreateCollections.Text = "Failed"
        $Global:CollectionsCreated = $false
        Return
        }
    Else {
        $Info = "--> Install Collection Moved"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
        }

    $Info = "Move Uninstall Collection"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
    Move-CMObject -InputObject (Get-CMDeviceCollection -Name $CompleteAppNameUninstCol) -FolderPath $CollectionFolderPath -ErrorAction SilentlyContinue -ErrorVariable ErrorAction 
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Collections" -Type Error
        $TextBox_CreateCollections.Background = "Red"
        $TextBox_CreateCollections.TextAlignment = "Center"
        $TextBox_CreateCollections.Text = "Failed"
        $Global:CollectionsCreated = $false
        Return
        }
    Else {
        $Info = "--> Uninstall Collection Moved"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Collections" -Type Info
        }

    $TextBox_CreateCollections.Background = "Green"
    $TextBox_CreateCollections.TextAlignment = "Center"
    $TextBox_CreateCollections.Text = "Created"
    $Button_CreateCollections.IsEnabled = $false
    $Button_DistributionSettings.IsEnabled = $true
    $Global:CollectionsCreated = $true
    }

Function Distribute-Content {

Param(
    [String]$Name,
    [String]$Version
    )

    $Info = "Enter Function Distribute-Content"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Distribute_Content" -Type Info
    Start-CMContentDistribution -ApplicationName $AppName -DistributionPointGroupName $ComboBox_DistributionPointGroup.Text -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
    
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Distribute_Content" -Type Error
        $TextBox_DistributeContent.Text = "Failed"
        $TextBox_DistributeContent.TextAlignment = "Center"
        $TextBox_DistributeContent.Background = "Red"
        $Global:ContentDistributed = $false
        Return
        }
    Else {
        $Info = "--> Content Distribution Started"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Distribute_Content" -Type Info
        }

    $TextBox_DistributeContent.Text = "Started"
    $TextBox_DistributeContent.TextAlignment = "Center"
    $TextBox_DistributeContent.Background = "Green"
    $Button_DistributeContent.IsEnabled = $false
    $Button_DeploymentSettings.IsEnabled = $true
    $Global:ContentDistributed = $true
    }

Function Create-Deployments {

Param (
    [String]$Name,
    [String]$Version
    )

    $Info = "Enter Function Create-Deployments"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Deployments" -Type Info
    $Info = "Create Install Deployment"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Deployments" -Type Info
    New-CMApplicationDeployment -CollectionName $CompleteAppNameInstCol -Name $AppName -DeployAction Install -DeployPurpose $ComboBox_DeployPurpose.Text -DeadlineDateTime (Get-Date) -AvailableDateTime (Get-Date) -UserNotification $ComboBox_UserNotification.Text -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Deployments" -Type Error
        $TextBox_CreateDeployments.Text = "Failed"
        $TextBox_CreateDeployments.TextAlignment = "Center"
        $TextBox_CreateDeployments.Background = "Red"
        Return
        }
    Else {
        $Info = "--> Install Deployment Created"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Deployments" -Type Info
        }

    $Info = "Create Uninstall Deployment"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Deployments" -Type Info
    New-CMApplicationDeployment -CollectionName $CompleteAppNameUninstCol -Name $AppName -DeployAction Uninstall -DeployPurpose $ComboBox_DeployPurpose.Text -DeadlineDateTime (Get-Date) -AvailableDateTime (Get-Date) -UserNotification $ComboBox_UserNotification.Text -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ("-->" + $ErrorAction | Out-String) -Component "Create_Deployments" -Type Error
        $TextBox_CreateDeployments.Text = "Failed"
        $TextBox_CreateDeployments.TextAlignment = "Center"
        $TextBox_CreateDeployments.Background = "Red"
        Return
        }
    Else {
        $Info = "--> Uninstall Deployment Created"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Create_Deployments" -Type Info
        }

    $TextBox_CreateDeployments.Text = "Created"
    $TextBox_CreateDeployments.TextAlignment = "Center"
    $TextBox_CreateDeployments.Background = "Green"
    $Button_CreateDeployments.IsEnabled = $false
    }

Function Load-Xaml ($Filename){
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($Filename)
    return $XamlLoader
    }


#endregion # Functions

#region # Create Logfile
$Info = "############################"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = "Start Script"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = "Version $Version"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = Get-Date
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = "User: $User"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = "Current Path: $Path"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = "Log Folder: $LogFolder"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = "Log Path: $LogPath"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
$Info = "Assembly Location: $AssemblyLocation"
Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Start" -Type Info
#endregion # Create Logfile

#region # XML
### Load XML
$XamlMainWindow=Load-Xaml($Path+"\MainWindow.xaml")
$Reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form = [Windows.Markup.XamlReader]::Load($Reader)

### Gui Objects
# Main Window
$Button_About = $Form.Findname("Button_About")
$Button_About_ChildWindow_Close = $Form.Findname("Button_About_ChildWindow_Close")
$Button_Exit = $Form.Findname("Button_Exit")
$TextBlock_Version = $Form.Findname("TextBlock_Version")
$Window_About = $Form.Findname("Window_About")

# App Tab
$Button_LoadInstaller = $Form.Findname("Button_LoadInstaller")
$Button_GeneratePackage = $Form.Findname("Button_GeneratePackage")
$Button_DetectionButton = $Form.Findname("Button_DetectionButton")
$Button_LoadPkgInfo = $Form.Findname("Button_LoadPkgInfo")
$Button_Icon = $Form.Findname("Button_Icon")
$Image_Icon = $Form.Findname("Image_Icon")
$TextBox_Vendor = $Form.Findname("TextBox_Vendor")
$TextBox_Name = $Form.Findname("TextBox_Name")
$TextBox_Version = $Form.Findname("TextBox_Version")
$TextBox_Language = $Form.Findname("TextBox_Language")
$TextBox_Revision = $Form.Findname("TextBox_Revision")
$TextBox_Installfile = $Form.Findname("TextBox_Installfile")
$TextBox_Installparameter = $Form.Findname("TextBox_Installparameter")
$TextBox_Sourcefolder = $Form.Findname("TextBox_Sourcefolder")
$TextBox_Message = $Form.Findname("TextBox_Message")
$TextBox_UninstallNameOrCode = $Form.Findname("TextBox_UninstallNameOrCode")
$TextBox_Uninstallparameter = $Form.Findname("TextBox_Uninstallparameter")
$ComboBox_Architecture = $Form.Findname("ComboBox_Architecture")
$ComboBox_Installtype = $Form.Findname("ComboBox_Installtype")
$ComboBox_Uninstalltype = $Form.Findname("ComboBox_Uninstalltype")
$ComboBox_Detection = $Form.Findname("ComboBox_Detection")
$Label_UninstallNameOrCode = $Form.Findname("Label_UninstallNameOrCode")

# App Tab Childwindow
$Window_Package_Generation = $Form.Findname("Window_Package_Generation")
$Label_PG_ChildWindow = $Form.Findname("Label_PG_ChildWindow")
$Button_PG_ChildWindow_Close = $Form.Findname("Button_PG_ChildWindow_Close")

# App Tab Detection Childwindow 
$Window_Detection_Script = $Form.Findname("Window_Detection_Script")
$TextBox_DetectionScript_ChildWindow = $Form.Findname("TextBox_DetectionScript_ChildWindow")
$Button_DetectionScript_ChildWindow_Close = $Form.Findname("Button_DetectionScript_ChildWindow_Close")

# MEMCM Tab
$Button_ConnectMEMCM = $Form.Findname("Button_ConnectMEMCM")
$Button_CreateApp = $Form.Findname("Button_CreateApp")
$Button_CreateCollections = $Form.Findname("Button_CreateCollections")
$Button_CreateDeployments = $Form.Findname("Button_CreateDeployments")
$Button_DistributeContent = $Form.Findname("Button_DistributeContent")
$Button_AppSettings = $Form.Findname("Button_AppSettings")
$Button_CollectionSettings = $Form.Findname("Button_CollectionSettings")
$Button_DistributionSettings = $Form.Findname("Button_DistributionSettings")
$Button_DeploymentSettings = $Form.Findname("Button_DeploymentSettings")
$TextBox_Sitecode = $Form.Findname("TextBox_Sitecode")
$TextBox_Managementpoint = $Form.Findname("TextBox_Managementpoint")
$TextBox_ConnectMEMCM = $Form.Findname("TextBox_ConnectMEMCM")
$TextBox_CreateApp = $Form.Findname("TextBox_CreateApp")
$TextBox_CreateCollections = $Form.Findname("TextBox_CreateCollections")
$TextBox_CreateDeployments = $Form.Findname("TextBox_CreateDeployments")
$TextBox_DistributeContent = $Form.Findname("TextBox_DistributeContent")

# MEMCM Tab Childwindow
$Window_MEMCM_Connection = $Form.Findname("Window_MEMCM_Connection")
$Label_MEMCM_ChildWindow = $Form.Findname("Label_MEMCM_ChildWindow")
$Button_MEMCM_ChildWindow_Close = $Form.Findname("Button_MEMCM_ChildWindow_Close")

# MEMCM Tab ApplicationSettings Childwindow
$Window_MEMCM_ApplicationSettings = $Form.Findname("Window_MEMCM_ApplicationSettings")
$Button_AS_ChildWindow_Close = $Form.Findname("Button_AS_ChildWindow_Close")
$ComboBox_InstallBehaviorType = $Form.Findname("ComboBox_InstallBehaviorType")
$ComboBox_LogonRequirementType = $Form.Findname("ComboBox_LogonRequirementType")

# MEMCM Tab CollectionSettings Childwindow
$Window_MEMCM_CollectionSettings = $Form.Findname("Window_MEMCM_CollectionSettings")
$Button_CS_ChildWindow_Close = $Form.Findname("Button_CS_ChildWindow_Close")
$TextBox_LimitingCollection = $Form.Findname("TextBox_LimitingCollection")
$TextBox_Folder = $Form.Findname("TextBox_Folder")

# MEMCM Tab DistributionSettings Childwindow
$Window_MEMCM_DistributionSettings = $Form.Findname("Window_MEMCM_DistributionSettings")
$Button_DistS_ChildWindow_Close = $Form.Findname("Button_DistS_ChildWindow_Close")
$ComboBox_DistributionPointGroup = $Form.Findname("ComboBox_DistributionPointGroup")

# MEMCM Tab DeploymentSettings Childwindow
$Window_MEMCM_DeploymentSettings = $Form.Findname("Window_MEMCM_DeploymentSettings")
$Button_DeplS_ChildWindow_Close = $Form.Findname("Button_DeplS_ChildWindow_Close")
$ComboBox_DeployPurpose = $Form.Findname("ComboBox_DeployPurpose")
$ComboBox_UserNotification = $Form.Findname("ComboBox_UserNotification")

# Config Tab
$Button_SaveConfig = $Form.Findname("Button_SaveConfig")
$TextBox_Destination = $Form.Findname("TextBox_Destination")
$TextBox_NamingAppFolder = $Form.Findname("TextBox_NamingAppFolder")
$TextBox_NamingApp = $Form.Findname("TextBox_NamingApp")
$TextBox_NamingLocalizedName = $Form.Findname("TextBox_NamingLocalizedName")
$TextBox_NamingDeploymentType = $Form.Findname("TextBox_NamingDeploymentType")
$TextBox_NamingCollectionInstall = $Form.Findname("TextBox_NamingCollectionInstall")
$TextBox_NamingCollectionUninstall = $Form.Findname("TextBox_NamingCollectionUninstall")
$CheckBox_PSADT = $Form.Findname("CheckBox_PSADT")

# Config Tab Childwindow
$Window_Configuration_Saved = $Form.Findname("Window_Configuration_Saved")
$Label_SC_ChildWindow = $Form.Findname("Label_SC_ChildWindow")
$Button_SC_ChildWindow_Close = $Form.Findname("Button_SC_ChildWindow_Close")

### Actions
# Main Window
$Button_About_ChildWindow_Close.Add_Click({
    $Window_About.IsOpen = $false
    })

$Button_About.Add_Click({
    $Window_About.IsOpen = $true
    })

$Button_Exit.Add_Click({
    $Info = "Exit GUI"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Exit" -Type Info
    $Info = "Clear Image Folder"
    Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "Exit" -Type Info
    #$Image_Icon.Source = $null
    $Images = (Get-ChildItem -Path $PSScriptRoot\Images -Exclude "GoLogo.png").FullName
    Foreach ($Image in $Images) {Remove-Item -Path $Image -Force}
    $Form.Close()
    })

# App Tab
$Button_LoadInstaller.Add_Click({
    Read-Installer
    })

$Button_GeneratePackage.Add_Click({
    Generate-Package -DestinationPath $TextBox_Destination.Text -SourcePath $TextBox_Sourcefolder.Text -Vendor $TextBox_Vendor.Text -Name $TextBox_Name.Text -Version $TextBox_Version.Text -Architecture $ComboBox_Architecture.Text -Revision $TextBox_Revision.Text -User $User -Installtype $ComboBox_Installtype.Text -Installfile $Textbox_Installfile.Text -Installparameter $Textbox_Installparameter.Text -Uninstalltype $ComboBox_Uninstalltype.Text -UninstallNameOrCode $TextBox_UninstallNameOrCode.Text -UninstallParameter $TextBox_Uninstallparameter.Text -Message $TextBox_Message.Text -Language $TextBox_Language.Text
    })

$Button_LoadPkgInfo.Add_Click({
    Load-PkgInfo
    })

$Button_Icon.Add_Click({
    Load-Icon
    })
   
$Button_PG_ChildWindow_Close.Add_Click({
    $Window_Package_Generation.IsOpen = $false
    })

$Button_DetectionButton.Add_Click({
    $Window_Detection_Script.IsOpen = $true
    })

$Button_DetectionScript_ChildWindow_Close.Add_Click({
    $Window_Detection_Script.IsOpen = $false
    $Global:DetectionScript = $TextBox_DetectionScript_ChildWindow.Text
    })

$ComboBox_Uninstalltype.add_Selectionchanged({
    If ($ComboBox_Uninstalltype.SelectedItem -eq "Product Name") {
            $Label_UninstallNameOrCode.Content = "Product Name"
            $TextBox_UninstallNameOrCode.text = $TextBox_Name.Text
            }
    If ($ComboBox_Uninstalltype.SelectedItem -eq "Product Code") {
            $Label_UninstallNameOrCode.Content = "Product Code"
            $TextBox_UninstallNameOrCode.text = $MSIProductcode 
            }
    If ($ComboBox_Uninstalltype.SelectedItem -eq "Script") {
            $Label_UninstallNameOrCode.Content = "Script"
            $TextBox_UninstallNameOrCode.text = "" 
            }
    })

$ComboBox_Detection.add_Selectionchanged({
    Switch ($ComboBox_Detection.SelectedIndex) {
        "0" {
            $Button_DetectionButton.Visibility = "Hidden"
            }
        "1" {
            $Button_DetectionButton.Visibility = "Visible"
            }
        } 
    })

# MEMCM Tab
$Button_ConnectMEMCM.Add_Click({
    Connect-MEMCM -SiteCode $TextBox_Sitecode.Text -ProviderMachineName $TextBox_Managementpoint.Text -ErrorAction SilentlyContinue -ErrorVariable ErrorAction
    If ($ErrorAction) {
        Write-Log -LogPath $LogPath -Message ($ErrorAction | Out-String) -Component "MEMCM_Connect" -Type Error
        $TextBox_ConnectMEMCM.Background = "Red"
        $TextBox_ConnectMEMCM.Text = "Failed"
        $TextBox_ConnectMEMCM.TextAlignment = "Center"
        $Global:MEMCMConnected = $false
        }
    Else {
        $Info = "Connected To MEMCM"
        Write-Log -LogPath $LogPath -Message ($Info | Out-String) -Component "MEMCM_Connect" -Type Info
        $Button_ConnectMEMCM.IsEnabled = $false
        $TextBox_ConnectMEMCM.Background = "Green"
        $TextBox_ConnectMEMCM.Text = "Connected"
        $TextBox_ConnectMEMCM.TextAlignment = "Center"
        $Button_AppSettings.IsEnabled = $true
        $Global:MEMCMConnected = $true
        }
    })

$Button_AppSettings.Add_Click({
    $Window_MEMCM_ApplicationSettings.IsOpen = $true
    })

$Button_AS_ChildWindow_Close.Add_Click({
    $Window_MEMCM_ApplicationSettings.IsOpen = $false
    If (($ComboBox_InstallBehaviorType.Text -ne "") -and ($ComboBox_LogonRequirementType.Text -ne "") -and ($MEMCMConnected -eq $true)) {
        $Button_CreateApp.IsEnabled = $true
        }
    })

$Button_CreateApp.Add_Click({
    Create-Application -DestinationPath $NewAppDestinationPath -Vendor $TextBox_Vendor.Text -Name $TextBox_Name.Text -Version $TextBox_Version.Text -Architecture $ComboBox_Architecture.Text -Language $TextBox_Language.Text -Installtype $ComboBox_Installtype.Text -Installfile $Textbox_Installfile.Text -Installparameter $Textbox_Installparameter.Text -Uninstalltype $ComboBox_Uninstalltype.Text -UninstallNameOrCode $TextBox_UninstallNameOrCode.Text -UninstallParameter $TextBox_Uninstallparameter.Text -User $User -DetectionScript $TextBox_DetectionScript_ChildWindow.Text -DetectionKey "" -DetectionFile ""
    })

$Button_CollectionSettings.Add_Click({
    $Window_MEMCM_CollectionSettings.IsOpen = $true
    })

$Button_CS_ChildWindow_Close.Add_Click({
    $Window_MEMCM_CollectionSettings.IsOpen = $false
    If (($TextBox_LimitingCollection.Text -ne "") -and ($TextBox_Folder.Text -ne "") -and ($AppCreated -eq $true)) {
        $Button_CreateCollections.IsEnabled = $true
        }
    })

$Button_CreateCollections.Add_Click({
    Create-Collections -Vendor $TextBox_Vendor.Text -Name $TextBox_Name.Text -Version $TextBox_Version.Text -Language $TextBox_Language.Text -Architecture $ComboBox_Architecture.Text
    })

$Button_DistributionSettings.Add_Click({
    $Window_MEMCM_DistributionSettings.IsOpen = $true
    })

$Button_DistS_ChildWindow_Close.Add_Click({
    $Window_MEMCM_DistributionSettings.IsOpen = $false
    If (($ComboBox_DistributionPointGroup.Text -ne "") -and ($CollectionsCreated -eq $true)) {
        $Button_DistributeContent.IsEnabled = $true
        }
    })

$Button_DistributeContent.Add_Click({
    Distribute-Content -Name $TextBox_Name.Text -Version $TextBox_Version.Text
    })

$Button_DeploymentSettings.Add_Click({
    $Window_MEMCM_DeploymentSettings.IsOpen = $true
    })

$Button_DeplS_ChildWindow_Close.Add_Click({
    $Window_MEMCM_DeploymentSettings.IsOpen = $false
    If (($ComboBox_DeployPurpose.Text -ne "") -and ($ComboBox_UserNotification.Text -ne "") -and ($ContentDistributed -eq $true)) {
        $Button_CreateDeployments.IsEnabled = $true
        }
    })

$Button_CreateDeployments.Add_Click({
    Create-Deployments -Name $TextBox_Name.Text -Version $TextBox_Version.Text
    })

$Button_MEMCM_ChildWindow_Close.Add_Click({
    $Window_MEMCM_Connection.IsOpen = $false
    })

# Config Tab
$Button_SaveConfig.Add_Click({
    If (Test-Path "$Path\Config.cfg") {Remove-Item "$Path\Config.cfg" -Force}
    If ($CheckBox_PSADT.IsChecked -eq $true) {"PSADT=1" | Out-File "$Path\Config.cfg" -Append}
        Else {"PSADT=0" | Out-File "$Path\Config.cfg" -Append}
    "DEST=" + $TextBox_Destination.Text | Out-File "$Path\Config.cfg" -Append
    "SITECODE=" + $TextBox_Sitecode.Text | Out-File "$Path\Config.cfg" -Append
    "SITESERVER=" + $TextBox_Managementpoint.Text | Out-File "$Path\Config.cfg" -Append
    "NAMEFOLDER=" + $TextBox_NamingAppFolder.Text | Out-File "$Path\Config.cfg" -Append
    "NAMEAPP=" + $TextBox_NamingApp.Text | Out-File "$Path\Config.cfg" -Append
    "NAMELOCALIZED=" + $TextBox_NamingLocalizedName.Text | Out-File "$Path\Config.cfg" -Append
    "NAMEDEPLOYMENTTYPE=" + $TextBox_NamingDeploymentType.Text | Out-File "$Path\Config.cfg" -Append
    "NAMEINSTALLCOLLECTION=" + $TextBox_NamingCollectionInstall.Text | Out-File "$Path\Config.cfg" -Append
    "NAMEUNINSTALLCOLLECTION=" + $TextBox_NamingCollectionUninstall.Text | Out-File "$Path\Config.cfg" -Append
    $Window_Configuration_Saved.IsOpen = $true
    $Label_SC_ChildWindow.Content = "Configuration Saved"
    })

$Button_SC_ChildWindow_Close.Add_Click({
    $Window_Configuration_Saved.IsOpen = $false
    })

$CheckBox_PSADT.Add_Checked({
    $ComboBox_Uninstalltype.Items.Clear()
    $ComboBox_Uninstalltype.Items.Add("Product Name")
    $ComboBox_Uninstalltype.Items.Add("Product Code")
    $ComboBox_Uninstalltype.Items.Add("Script")
    })

$CheckBox_PSADT.Add_UnChecked({
    $ComboBox_Uninstalltype.Items.Clear()
    $ComboBox_Uninstalltype.Items.Add("Product Code")
    $ComboBox_Uninstalltype.Items.Add("Script")
    $ComboBox_Uninstalltype.SelectedItem = "Product Code"
    })


### Preloaded Configuration
$TextBlock_Version.Text = "Version $Version"
If ($Config.PSADT -eq "1") {$CheckBox_PSADT.IsChecked = $true}
$TextBox_Sitecode.AppendText($Config.SITECODE)
$TextBox_Managementpoint.AppendText($Config.SITESERVER)
$TextBox_Destination.AppendText($Config.DEST)
$TextBox_NamingAppFolder.AppendText($Config.NAMEFOLDER)
$TextBox_NamingApp.AppendText($Config.NAMEAPP)
$TextBox_NamingLocalizedName.AppendText($Config.NAMELOCALIZED)
$TextBox_NamingDeploymentType.AppendText($Config.NAMEDEPLOYMENTTYPE)
$TextBox_NamingCollectionInstall.AppendText($Config.NAMEINSTALLCOLLECTION)
$TextBox_NamingCollectionUninstall.AppendText($Config.NAMEUNINSTALLCOLLECTION)

### Start GUI
$Form.ShowDialog() | Out-Null
#endregion # XML
