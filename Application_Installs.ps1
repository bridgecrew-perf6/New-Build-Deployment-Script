# Global Variables
$ScriptName = "Application_Installs.ps1"
$ScriptDir = "App_Installs\"

# Set Execution Policy
Set-ExecutionPolicy -ExecutionPolicy Bypass

# Install Attachmate
function Install-Attachmate {

    $AEInstall = "Attachmate\Attachmate-Extra-v93.EXE"
    $AEAppName = "Attachmate EXTRA! X-treme 9.3"
    $AEAppInstalled = (

        Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $AEAppName }

        )

    try {

        if(-Not $AEAppInstalled) {

            Write-Host "Installing $AEAppName.."
            Start-Process -FilePath $AEInstall -Wait           
            Write-Host "$AEAppName installation completed successfully!" -ForegroundColor Green
        }

        else { Write-Host "$AEAppName is already installed on your system" -ForegroundColor Yellow }

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}


# Install Avaya One-X Agent
function Install-AvayaOneXAgent {

    $AvayaInstall = "Avaya One-X Agent\Avaya one-X Agent - 2.5.13.msi"
    $AvayaAppName = "Avaya one-X Agent - 2.5.13"
    $AvayaAppInstalled = (

        Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $AvayaAppName }

        )

    try {

        if(-Not $AvayaAppInstalled) {

            Write-Host "Installing $AvayaAppName..."
            Start-Process -FilePath $AvayaInstall -ArgumentList "/qn" -Wait           
            Write-Host "$AvayaAppName installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "$AvayaAppName is already installed on your system" -ForegroundColor Yellow }
    
    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }   

}

# Install Xorceview
function Install-Xorceview {

    $XVInstall = "Xorceview\GLO_Spectrum Corp_XorceView_2.2.0.1_MSI_v1.msi"
    $XVAppName = "Xorceview"

    try {

        if(-Not $XVAppInstalled) {

            Write-Host "Installing $XVAppName..."
            Start-Process -FilePath $XVInstall -ArgumentList "/qn" -Wait            
            Write-Host "$XVAppName installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "$XVAppName is already installed on your system" -ForegroundColor Yellow }

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Install VCAT TLS Update
function Install-VCAT_TLS_Update {

    $VTLSInstall = "VCAT TLS Update\Install.bat"
    $VTLSAppName = "VCAT 7.1 TLS Update"

    try {

        Write-Host "Installing $VTLSAppName..."
        Start-Process -FilePath $VTLSInstall -Wait
        Write-Host "$VTLSAppName installation completed successfully!" -ForegroundColor Green

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Install Global Protect
function Install-GlobalProtect {

    $GPInstall = "Global Protect\Deploy-Application.exe"
    $GPVersion = "5.1.8"
    $GPAppName = "GlobalProtect"
    $GPVersInstalled = (

        Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayVersion -Match $GPVersion }

        )

    $GPAppInstalled = (

        Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $GPAppName }

        )

    $GProcess = @("PanGPA", "PanGPS")
    $GProcessCheck = Get-Process $GProcess -ErrorAction SilentlyContinue


    try {

        Write-Host "Installing $GPAppName $GPVersion..."
        Write-Host "Checking if any $GPAppName processes are running..."

        if ($null -ne $GProcessCheck) {

            Write-Host "Stopping $GPAppName Processes..."
            $GProcessCheck | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Host "Processes successfully stopped. Proceeding with installation..."
        }

        else { Write-Host "Unable to find any running processes. Proceeding with installation..." }


        if((-Not $GPAppInstalled) -or (-Not $GPVersInstalled)) {

            Start-Process -FilePath $GPInstall -Wait
            Write-Host "$GPAppName $GPVersion installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "$GPAppName $GPVersion is already installed on your system" -ForegroundColor Yellow }
    
    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }  

}

# Install Oracle 11g
function Install-Oracle {

    $OraInstall = "Oracle 11\Oracle_Client_11g_ENG_R1_Install.vbs"
    $OraAppName = "Oracle Client 11g"
    $OraAppInstalled = (

        Get-ItemProperty -Path HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $OraAppName }

        )

    $TNSPath = "HKLM:\SOFTWARE\Wow6432Node\ORACLE\KEY_OraClient11g_home1"
    $TNSKey = "TNS_ADMIN"
    $TNSValue = "\\path\to\tns\directory"


    try {

        if(-Not $OraAppInstalled) {

            Write-Host "Installing $OraAppName..."
            Start-Process -FilePath $OraInstall -Wait         
            Write-Host "$OraAppName installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "Oracle 11g is already installed on your system" -ForegroundColor Yellow }

        # Add the TNS_NAMES registry key
        Get-Item -Path $TNSPath | New-ItemProperty -Name $TNSKey -Value $TNSValue -PropertyType String -Force
        Get-ItemProperty -Path $TNSPath -Name $TNSKey
        Write-Host "Successfully added TNS_NAMES key to the registry" -ForegroundColor Green

    }    

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Install Rightfax
function Install-Rightfax {

    $RFXInstall = "RightFax\RightFax Product Suite - Client.EXE"
    $RFXAppName = "RightFax Product Suite - Client"
    $RFXAppInstalled = (

        Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $RFXAppName }

        )

    try {

        if(-Not $RFXAppInstalled) {

            Write-Host "Installing $RFXAppName..."
            Start-Process -FilePath $RFXInstall -Wait         
            Write-Host "$RFXAppName installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "$RFXAppName is already installed on your system" -ForegroundColor Yellow }

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Install NICE codec package
function Install-NICE {

    $NICEInstallDir = "NICE\Install"
    $NICEAppName = "NICE Engage Platform Release 6.5 - Player Codec Pack"
    $NICEAppInstalled = (

        Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $NICEAppName }

        )

    try {

        if(-Not $NICEAppInstalled) {

            Write-Host "Installing $NICEAppName..."
            Get-ChildItem -Path $NICEInstallDir | ForEach-Object { Start-Process -FilePath $_.FullName -ArgumentList "/qn" -Wait }         
            Write-Host "$NICEAppName installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "$NICEAppName is already installed on your system" -ForegroundColor Yellow }

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Install Lotus Notes
function Install-IBM_Notes {

    $NotesInstall = "IBM Notes 9\Files\setup.exe"
    $NotesAppName = "IBM Notes 9.0.1 Social Edition"
    $NotesAppInstalled = (

        Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $NotesAppName }

        )

    try {

        if(-Not $NotesAppInstalled) {

            Write-Host "Installing $NotesAppName..."
            Start-Process -FilePath $NotesInstall -Wait          
            Write-Host "$NotesAppName installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "$NotesAppName is already installed on your system" -ForegroundColor Yellow }

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Install SAP
function Install-SAP {

    $SAPInstallDir = "SAP\Install"
    $SAPAppName = "SAP GUI for Windows 7.60"
    $SAPAppInstalled = (

        Get-ItemProperty -Path HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | 
        Where-Object { $_.DisplayName -Match $SAPAppName }

        )

    try {

        if(-Not $SAPAppInstalled) {

            Write-Host "Installing $SAPAppName..."
            Get-ChildItem -Path $SAPInstallDir | ForEach-Object { Start-Process $_.FullName -Wait }
            Write-Host "$SAPAppName installation completed successfully!" -ForegroundColor Green

        }

        else { Write-Host "$SAPAppName is already installed on your system" -ForegroundColor Yellow }

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Install VCAT
function Install-VCAT {

    $VCATInstall = "VCAT_7_1\VCAT Setup Package.msi"
    $VCATAppName = "VCAT"

    try {

        Write-Host "Installing $VCATAppName..."
        Start-Process -FilePath $VCATInstall -ArgumentList "/qn" -Wait          
        Write-Host "$VCATAppName installation completed successfully!" -ForegroundColor Green

    }

    catch [System.Management.Automation.ItemNotFoundException] {

        Write-Host "Installation files not found. Please ensure that the $ScriptName file is in the $ScriptDir directory" -ForegroundColor Red

    }

}

# Add "Authenticated Users" group to VCAT registry key with "Full Control" permissions
function Add-AuthenticatedUsersGroup { 

    $RegKey = "HKLM:\SOFTWARE\WOW6432Node\VCAT"
    $acl = Get-Acl -Path $RegKey
    $idName = "NT AUTHORITY\Authenticated Users"
    $idRef = [System.Security.Principal.NTAccount]($idName)
    $regRights = [System.Security.AccessControl.RegistryRights]::FullControl
    $inhFlags = [System.Security.AccessControl.InheritanceFlags]::None
    $propFlags = [System.Security.AccessControl.PropagationFlags]::None
    $acType = [System.Security.AccessControl.AccessControlType]::Allow
    $typeName = "System.Security.AccessControl.RegistryAccessRule"
    $argsList = @($idRef, $regRights, $inhFlags, $propFlags, $acType)
    $RegAccessRule = New-Object -TypeName $typeName -ArgumentList $argsList
    
    # Add the new registry access rule if not already present
    
    $CheckRegAccessRule = $acl.Access.Where({$_.IdentityReference -eq $idName})
    if(-Not($CheckRegAccessRule)) {
    
        Write-Host "Adding Authenticated Users group..."
        # Create a backup of the registry key
        Invoke-Command -ScriptBlock {reg export HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\VCAT VCATRegKeyBackup.reg /y}
        $acl.AddAccessRule($RegAccessRule)
        Set-Acl -Path $RegKey -AclObject $acl
        $acl.Access
        Write-Host "Authenticated Users group added successfully!" -ForegroundColor Green
    }
    
    else {Write-Host "The Authenticated Users group is already present!" -ForegroundColor Yellow}

}



