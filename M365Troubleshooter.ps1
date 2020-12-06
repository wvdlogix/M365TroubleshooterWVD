cd "C:\Users\Patrick\OneDrive - koehlercloud2019\02 - WVDLogix\WVDLogix\07 - Tools\MicrosoftAppsForEntperirseTroubleshooter"
$InstallPath = Get-Childitem -Path "C:\Temp\M365Troubleshooter" -ErrorVariable $InstallApp
$InstallApp = New-Item -Path "C:\Temp\M365Troubleshooter" -ItemType "Directory"
cd $InstallApp
Invoke-WebRequest -Uri https://github.com/wvdlogix/M365TroubleshooterWVD/blob/main/M365TSWD.xaml -OutFile "MainWindow.xaml"

###Windows Forms###
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
$xamlFile = "MainWindow.xaml"

###GUI###
#create window
$inputXML = Get-Content $xamlFile -Raw
$inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
[XML]$XAML = $inputXML

#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
try {
    $window = [Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Warning $_.Exception
    throw
}

# Create variables based on form control names.
# Variable will be named as 'var_<control name>'

$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    #"trying item $($_.Name)"
    try {
        Set-Variable -Name "var_$($_.Name)" -Value $window.FindName($_.Name) -ErrorAction Stop
    } catch {
        throw
    }
}
Get-Variable var_*

###GUI Frame Metadata###
$var_lblTime.Content = Get-Date -Format "HH:MM - dd/MM/yy"
$var_lblUser.content = $env:USERNAME + "." + $env:UserDOMAIN

$var_btnWorkingDirectoryBrowse.Add_Click( {

###Extract content to one folder###
$WorkingFolder = New-Object System.Windows.Forms.FolderBrowserDialog
$WorkingFolder.rootfolder = "MyComputer"
$WorkingFolder.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })) 
$WorkingFolderPath = $WorkingFolder.SelectedPath
$ParamsSys = @{
"$WorkingFolderPath" = $WorkingFolder.SelectedPath    
}

Set-Location -Path $WorkingFolderPath
$var_lblWorkingDirectory.content = $WorkingFolderPath

###Preparation of Tools###
Invoke-WebRequest -Uri https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip -OutFile WPJCleanUp.zip 
Invoke-WebRequest -Uri https://download.microsoft.com/download/e/1/b/e1bbdc16-fad4-4aa2-a309-2ba3cae8d424/OLicenseCleanup.zip -OutFile OLicenseCleanup.zip
Invoke-WebRequest -Uri https://download.microsoft.com/download/f/8/7/f8745d3b-49ad-4eac-b49a-2fa60b929e7d/signoutofwamaccounts.zip -OutFile signoutofwamaccounts.zip

###Extraction of ZIP containers###
Expand-Archive -Path $WorkingFolderPath\OLicenseCleanup.zip -DestinationPath $WorkingFolderPath
Expand-Archive -Path $WorkingFolderPath\signoutofwamaccounts.zip -DestinationPath $WorkingFolderPath
Expand-Archive -Path $WorkingFolderPath\WPJCleanUp.zip -DestinationPath $WorkingFolderPath

###Cleanup files###
Remove-Item * -Include *.zip
})




###Gather issue relevant information###
$officeVersion = Get-WmiObject win32_product | where{$_.Name -like "*Office 16 Click-to-Run Licensing Component*"} | select Name,Version
$Office365RegistryLocation = Get-ItemProperty -path HKLM:SOFTWARE\Microsoft\Office\ClickToRun\Configuration\
$activationType = $Office365RegistryLicensingLocation.SharedComputerLicensing
$SCACacheOverride = $Office365RegistryLicensingLocation.SCLCacheOverride
$SCACacheOverrideDirectory = $Office365RegistryLicensingLocation.SCLCacheOverrideDirectory 
$activationTokenLocation = "C:\Users\$env:USERNAME.$env:UserDOMAIN\AppData\Local\Microsoft\Office\16.0\Licensing"
$activationTokenLocationTokenPresent = Get-ChildItem $activationTokenLocation | Measure-Object 
$activationTokenLocationTokenPresentWriteTime = Get-ChildItem $activationTokenLocation
$lastM365Activation = $activationTokenLocationTokenPresentWriteTime | Select LastWriteTime | Select-Object -first 1

###GUI Information Fill###
$var_lblVersion.content = $officeVersion.Version
if ($activationType -eq $null) {$var_lblActivationType.content = "Local Computer Activation"} else {$var_lblActivationType.content = "Shared Computer Activation"}
if ($SCACacheOverride -eq $null) {$var_lblRoamingTokenEnabled.content = "N/A"} else {$var_lblRoamingTokenEnabled.content = "Roaming Enabled"}
if ($SCACacheOverrideDirectory -eq $null) {$var_lblRoamingTokenLocation.content = "N/A"} else {$var_lblRoamingTokenLocation.content = $SCACacheOverrideDirectory}
if ($lastM365Activation -eq $null) {
$var_lblLastSuccessfulActivation.content = "N/A"
$var_btnRUN.Visibility = "Visible"
} else {
$var_lblLastSuccessfulActivation.content = $lastM365Activation
$var_lblNoInteract.Visibility = "Visible"
}

$var_btnRUN.Add_Click( {

###Troubleshoot activation issue###
if ($activationType -eq "1") 
    {
        Write-Host "Your Microsoft Office Version is"  $officeVersion.Version  "this computer is activated through Shared Computer licensing" -ForegroundColor White -BackgroundColor Green
        Write-Host "Validation of licensing token in local user profile" 
        if ($activationTokenLocation) 
            {
            Write-Host "The token folder is present underneath" + $activationTokenLocation -ForegroundColor White -BackgroundColor Green
            Write-Host "Validating Token presence..." 

                if ($activationTokenLocationTokenPresent.count -eq "2")
                    {
                    Write-Host "The Microsoft365 licensing files are present - No action to perform" -ForegroundColor White -BackgroundColor Green
                    Write-Host "The activation for this user occured" $lastM365Activation.LastWriteTime -ForegroundColor White -BackgroundColor Green
                    }
                    else 
                    {
                    Write-Host "Executing reset process" -ForegroundColor White -BackgroundColor Red
                     .\OLicenseCleanup.vbs
                     Start-Sleep -Seconds 10 
                     .\WPJCleanUp\WPJCleanUp.cmd
                     Start-Sleep -Seconds 5 
                        Write-Host "Reset process sucessful! Starting Microsoft Excel for validation - Please log on if Office asks you"
                            cd "C:\Program Files\Microsoft Office\root\Office16\"
                            .\EXCEL.exe
                    }
            
            }
            else 
            { 
                   Write-Host "Executing reset process" -ForegroundColor White -BackgroundColor Red
                     .\OLicenseCleanup.vbs
                     Start-Sleep -Seconds 10 
                     .\WPJCleanUp\WPJCleanUp.cmd
                     Start-Sleep -Seconds 5 
                        Write-Host "Reset process sucessful! Starting Microsoft Excel for validation - Please log on if Office asks you"
                            cd "C:\Program Files\Microsoft Office\root\Office16\"
                            .\EXCEL.exe}
    } 
   
else {Write-Host "You are using a local installation of Microsoft 365 Apps for Enterprise - No action can be taken" -ForegroundColor White -BackgroundColor Green}

 } ) 
$Null = $window.ShowDialog()