<#
    .SYNOPSIS
    Configures a Windows 10 Computer to the Standards of Root Service AG

    .DESCRIPTION
    This Script completly configures a computer for the Root Service AG.
    Some Features may not currently work or are yet to be implemented.
    For more Detail read the "Readme.md" File

    .PARAMETER PCName
    Specifies the New PC Name that is set by the Script

    .PARAMETER Key
    Specifies the WindowsKey needed to activate Windows

    .PARAMETER MultiDrve
    Tells the Script if there is a second Disk that has to be renamed to
    [SWITCH PARAM]

    .PARAMETER User
    If you add the User Parameter, only cosmetic functions will be runned
    [SWITCH PARAM]

    .PARAMETER Func
    Specifies wich funtions should be run. Takes an Array of Strings as an input (Single string also possible)
    List of Funtions:
    - TestInternet
    - ShowDesktopIcons
    - InstallStandardApps
    - ActivateWindows
    - DelWindowsOld
    - DelDefaultUsers
    - RenameDisk
    - RenameDevice
    - SetEnergyOptions
    - ActivateRDP
    - RemoveOneDrive
    - SetSystemProperties
    - ActivateNumLock
    - RemoveLanguagebar
    - GetTeamViewer
    - RemoveDefaultSoftware
    - RemoveTiles
    - ActivateAdmin
    - EndScript

    .INPUTS
    Only the above specified Parameters can be taken as input.

    .OUTPUTS
    Errors. The Script will only show info about the current function and
    Errors if they occur.

    .EXAMPLE
    PS> .\Setup.ps1 NB50

    .EXAMPLE
    PS> .\Setup.ps1 -PCName NB50

    .EXAMPLE
    PS> .\Setup.ps1 -PCName NB50 -Key "XdhXX-XzuXX-XjhXX-casXX-XX56X" -MultiDrive

    .EXAMPLE
    PS> .\Setup.ps1 -User

    .EXAMPLE
    PS> .\Setup.ps1 -Func RemoveTiles

    .EXAMPLE
    PS> .\Setup.ps1 -Func "RemoveTiles,RemoveLanguagebar"

    .EXAMPLE
    PS> .\Setup.ps1 -Func @{ RemoveTiles, RemoveLanguagebar }

    .LINK
    Github: https://github.com/drmgames5/standard-pc-setup.git
    Note that the Repository is currently Private. You need to ask mdo to get invited.
#>

param (
    [string] $PCName,
    [string] $Key,
    [switch] $MultiDrive,
    [switch] $User,
    [validateset(
        "TestInternet",
        "ShowDesktopIcons",
        "InstallStandardApps",
        "ActivateWindows",
        "DelWindowsOld",
        "DelDefaultUsers",
        "RenameDisk",
        "RenameDevice",
        "SetEnergyOptions",
        "ActivateRDP",
        "RemoveOneDrive",
        "SetSystemProperties",
        "ActivateNumLock",
        "RemoveLanguagebar",
        "GetTeamViewer",
        "RemoveDefaultSoftware",
        "RemoveTiles",
        "ActivateAdmin",
        "EndScript"
    ) ] [string[]] $Func
)

#[] Logging

#Activate Verbose for Transscript
$VerbosePreference = "Continue"

#Log-Directory
$cDate = Get-Date -Format "yyyy-MM-dd_HH-mm"
$logdir = "$pwd\Log_$cDate\"

#log Ordner erstellen
If(!(Test-Path $logdir)){
    try { mkdir $logdir }
    catch {
        #Log Error - cannot use LogWrite cause its initiated later
        Write-Verbose "ERROR | $logdir could not be created."
        Write-Verbose "Error:" $_
    }
}

# Transcript logging
Start-Transcript -OutputDirectory $logdir

# Logging function
Function LogWrite {
    Param ([string]$logstring)

    $Logfile = $logdir + "Setup.log"
    $currentDate = Get-Date -Format "[HH:mm] [yyyy-MM-dd]"
    $ExePolicy = Get-ExecutionPolicy

    Add-content $Logfile -value "$currentDate | $logstring | $ExePolicy" 
    Write-Verbose $logstring;
}

#[] Internet Verbindung testen
function TestInternet() {
    While (!(Test-Connection -computer google.com -count 1 -quiet)) {
        Write-Host -ForegroundColor Red "Keine Internet Verbindung!"
        Start-Sleep -Seconds 2 
    } 
    Write-Host -ForegroundColor Green "Internet Verbunden"
}

#00 Alle Desktopsymbole anzeigen
function ShowDesktopIcons() {
    #Log Func
    LogWrite "FUNC | 00 Alle Desktop Symbole werden angezeigt"

    $ErrorActionPreference = "SilentlyContinue"

    $RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    If (Test-Path $RegistryPath) {
        If (-Not(Get-ItemProperty -Path $RegistryPath -Name "HideIcons")) {
            try { New-ItemProperty -Path $RegistryPath -Name "HideIcons" -Value "0" -PropertyType DWORD -Force }
            catch {
                #Log Warning
                LogWrite "WARN | HideIcons RegistryKey ist nicht vorhanden und konnte nicht erstellt werden"
                Write-Verbose "Error:" $_
            }
        }
        $Check = (Get-ItemProperty -Path $RegistryPath -Name "HideIcons").HideIcons
        If ($Check -NE 0) {
            try { New-ItemProperty -Path $RegistryPath -Name "HideIcons" -Value "0" -PropertyType DWORD -Force }
            catch {
                #Log Warning
                LogWrite "WARN | HideIcons RegistryKey konnte nicht erstellt werden"
                Write-Verbose "Error:" $_
            }
        }
    } else {
        #Log Error
        LogWrite "Error | $RegistryPath nicht vorhanden"
    }
    
    $RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons"
    If (-Not(Test-Path $RegistryPath)) {
        try { New-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer" -Name "HideDesktopIcons" -Force }
        catch {
            #Log Warning
            LogWrite "WARN | '$RegistryPath' RegistryKey konnte nicht erstellt werden"
            Write-Verbose "Error:" $_
        }
    
        try { New-Item -Path $RegistryPath -Name "NewStartPanel" -Force }
        catch {
            #Log Warning
            LogWrite "WARN | '$RegistryPath\NewStartPanel' RegistryKey konnte nicht erstellt werden"
            Write-Verbose "Error:" $_
        }
    } 

    $RegistryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel"

    If (Test-Path $RegistryPath) {
        $iconKeys = @{
            RecycleBin ="{645FF040-5081-101B-9F08-00AA002F954E}"
            MyComputer = "{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
            ControlPanel  = "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}"
            UserFiles = "{59031a47-3f72-44a7-89c5-5595fe6b30ee}"
        }
        foreach ($icon in $iconKeys.values){
            $Res = Get-ItemProperty -Path $RegistryPath -Name $icon
            If (-Not($Res)) {
                try { New-ItemProperty -Path $RegistryPath -Name $icon -Value "0" -PropertyType DWORD -Force }
                catch {
                    #Log Warning
                    LogWrite "WARN | Icon RegistryKey konnte nicht erstellt werden"
                    Write-Verbose "Error:" $_
                }
            }
            $Check = (Get-ItemProperty -Path $RegistryPath -Name $icon).$icon
            If ($Check -NE 0) {
                New-ItemProperty -Path $RegistryPath -Name $icon -Value "0" -PropertyType DWORD -Force
                try { New-ItemProperty -Path $RegistryPath -Name $icon -Value "0" -PropertyType DWORD -Force }
                catch {
                    #Log Warning
                    LogWrite "WARN | Icon RegistryKey konnte nicht erstellt werden"
                    Write-Verbose "Error:" $_
                }
            }
        }
    } else {
        #Log Error
        LogWrite "Error | $RegistryPath nicht vorhanden"
    }
}

#01 Standard Programme installieren
function InstallStandardApps() {
    #Log Func
    LogWrite "FUNC | 01 Standardprogramme werden installiert"

    #Zwischenablage für exe erstellen
    $workdir = "c:\installer\"

    If (Test-Path -Path $workdir -PathType Container)
    { Write-Verbose "$workdir already exists" }
    ELSE
    { 
        try { New-Item -Path $workdir  -ItemType directory }
        catch {
            #Log Error
            LogWrite "ERROR | $workdir could not be created."
            Write-Verbose "Error:"  $_
        }
    }

    #CSV mit Standard-Software angaben einlesen
    try { $StandardApps = Import-Csv ".\StandardApps.csv" }
    catch {
        #Log Error
        LogWrite "ERROR | CSV konnte nicht importiert werden."
        Write-Verbose "Error:" $_
    }
    
    #Log Info
    LogWrite "INFO | Starte Downloads und installationen von StandardApps"

    foreach ($App in $StandardApps) {

        #Log Info
        LogWrite "INFO | Software-Infos aus CSV: " $App

        #DateiPfad nach download
        $destination = $workdir+$App.Name+".exe"

        #Installer herunterladen
        try { Invoke-WebRequest $App.Link -OutFile $destination }
        catch {
            #Log Error
            LogWrite "ERROR | Software konnte nicht heruntergeladen werden"
            Write-Verbose "Variables:" $App.Link " | " $destination
            Write-Verbose "Error:"  $_
        }

        #Installation
        If ([string]::IsNullOrEmpty($App.ArgumentList)){
            try { Start-Process -FilePath $destination -Wait }
            catch {
                #Log Error
                LogWrite "ERROR | " $App.Name " konnte nicht installiert werden."
                Write-Verbose "Variables: " $destination
                Write-Verbose "Error:" $_
            }
        } else {
            try { Start-Process -FilePath $destination -NoNewWindow -Wait -ArgumentList $App.ArgumentList }
            catch {
                #Log Error
                LogWrite "ERROR | " $App.Name " konnte nicht installiert werden."
                Write-Verbose "Variables: " $destination " " $App.ArgumentList
                Write-Verbose "Error:" $_
            }
        }   
    }

    #Zwischenablage löschen
    try { Remove-Item -Recurse -Force $workdir }
    catch {
        #Log Error
        LogWrite "ERROR | $workdir could not be deleted."
        Write-Verbose "Error:" $_
    }

    #DesktopIcons Greenshot und Adobe aus Public Desktop löschen
    If(Test-Path "C:\Users\Public\Desktop\Greenshot.lnk"){
        try { Remove-Item "C:\Users\Public\Desktop\Greenshot.lnk" }
        catch {
            #Log Error
            LogWrite "ERROR | Greenshot.lnk could not be deleted from Public Desktop."
            Write-Verbose "Error:" $_
        }
    }
    else {
        #Log Warning
        LogWrite "WARN | Greenshot.lnk nicht gefunden"
    }
    If(Test-Path "C:\Users\Public\Desktop\Acrobat Reader DC.lnk"){
        try { Remove-Item "C:\Users\Public\Desktop\Acrobat Reader DC.lnk" }
        catch {
            #Log Error
            LogWrite "ERROR | Acrobat Reader DC.lnk could not be deleted from Public Desktop."
            Write-Verbose "Error:" $_
        }
    }
    else {
        #Log Warning
        LogWrite "WARN | Acrobat Reader DC.lnk nicht gefunden"
    }
}

#02 Windows aktivieren
function ActivateWindows() {
    #Log Func
    LogWrite "FUNC | 02 Windows wird aktiviert"

    If ($null -ne $Key) {
        #aktivieren
        $computer = Get-Content env:computername
        $service = get-wmiObject -query "select * from SoftwareLicensingService" -computername $computer
        $service.InstallProductKey($Key)
        $service.RefreshLicenseStatus()

        #Überprüfen ob win Aktiviert ist
        $Status = (Get-CimInstance -ClassName SoftwareLicensingProduct -Filter "Name like 'Windows%'" | Where-Object PartialProductKey).licensestatus
        If ($Status -ne 1) {
            #Log Error -> windows nicht aktiviert
            LogWrite "ERROR | Windows konnte nicht aktiviert werden"
            Write-Verbose "[key=$Key] [computername=$computer] [service=$service]"
        }
        else {
            #Log INFO -> win aktiv
            LogWrite "INFO | Windows ist aktiviert"
        }
    }
    else {
        #Log INFO -> win aktiv
        LogWrite "INFO | Kein Windows Key angegeben"
    } 
}

#03 Delete Windows.old
function DelWindowsOld() {
    #Log Func
    LogWrite "FUNC | 03 Windows.old Ordner wird gelöscht"

    If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {   
        $arguments = "& '" + $myinvocation.mycommand.definition + "'"
        try { Start-Process powershell -Verb runAs -ArgumentList $arguments }
        catch {
            #Log Error
            LogWrite "ERROR | '$arguments' Befehl nicht ausgeführt"
            Write-Verbose "Error:" $_
        }
        Break
    } 

    #Windows.old path 
    $path = $env:HOMEDRIVE + "\windows.old"
    If (Test-Path -Path $path) {
        #create registry value
        $regpath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations"
        try { New-ItemProperty -Path $regpath -Name "StateFlags1221" -PropertyType DWORD  -Value 2 -Force }
        catch {
            #Log Warning
            LogWrite "WARN | '$regpath\StateFlags1221' Registry key nicht erstellt"
            Write-Verbose "Error:" $_
        }

        #start clean application
        try { cleanmgr /SAGERUN:1221 }
        catch {
            #Log Warning
            LogWrite "WARN | cleanmgr wurde nicht durchgeführt."
            Write-Verbose "Error:" $_
        }
    }
    else {   
        Write-Verbose "There is no 'Windows.old' folder in system driver"
        
    }
    
    #Logging
    If (Test-Path -Path $path) {
        #Log Error -> win old existiert noch
        LogWrite "ERROR | Windows.old Ordner existiert noch | [regpath=$regpath] [path=$path]"
    }
    else {
        #Log Info -> nicht existent
        LogWrite "INFO | Windows.old wurde gelöscht oder war nicht vorhanden"
    }
}

#04 Benutzer mit Profil löschen (Benutzer Defaultuser01 und rs1)
function DelDefaultUsers() {
    #Log Func
    LogWrite "FUNC | 04 Unnötige Benutzer werden gelöscht"

    # Array für die User die gelöscht werden sollen
    $uselessUser = @('rs1', 'defaultuser01', 'Defaultuser01')

    # T - cleanup
    # For Each -> user löschen und prüfen ob noch da  
    foreach ($userX in $uselessUser) {

        $Uexist = Get-LocalUser | Where-Object { $_.Name -eq "$userX" } | Select-Object -ExpandProperty Name

        If ($Uexist -eq $userX) {
            (Get-CimInstance -Class Win32_UserProfile | Where-Object { $_.LocalPath.split('\')[-1] -eq '$userX' } | Remove-CimInstance)
            Remove-LocalUser -Name "$userX" 
            $Uexist = Get-LocalUser | Where-Object { $_.Name -eq "$userX" } | Select-Object -ExpandProperty Name

            If ($Uexist -eq $userX) {
                #Log Error
                LogWrite "ERROR | $userX wurde nicht gelöscht | [userX = $userX] [Uexist=$Uexist]"
            } 
            else {
                #Log Info
                LogWrite "INFO | $userX wurde erfolgreich gelöscht"
            } 
        }
        else {
            #Log Info 
            LogWrite "INFO | $userX nicht vorhanden"
        }      
    }
} 

#05 Laufwerke umbenennen
function RenameDisk() {
    #Log Func
    LogWrite "FUNC | 05 Laufwerke werden Umbennant" 

    $Volumes = @{ 
        C = "System"
    }

    if ($MultiDrive) {
        $Volumes.add( 'D', "Daten" )
    }

    foreach ($vol in $Volumes.GetEnumerator()){
        $dL = $vol.Key
        $dN = $vol.Value

        #Festplatte umbennnen
        Set-Volume -DriveLetter $dL -NewFileSystemLabel $dN

        #überprüfen
        $mD = Get-Volume -DriveLetter $dL | Select-Object -ExpandProperty FileSystemLabel
        If ($md -eq $dN) {
            #Log Info -> erfolg
            LogWrite "INFO | $dN Festplatte wurde umbennant"
        }
        else {
            #Log Warning -> nicht oder falsch bennant
            LogWrite "WARN | $dN Festplatte wurde nicht oder falsch bennant | [mD=$mD]"
        } 
    }
}

#06 Computer Umbenennen
function RenameDevice() {
    #Log Func
    LogWrite "FUNC | 06 Computer wird umbennant in: $PCName -> keine Logs" 

    $Computer = Get-Content env:computername

    Rename-Computer -ComputerName $Computer -NewName $PCName -Force

    # T - Prüfen ob Gerät umbennant wurde, regedit newname auslesen
}

#07 Energieoptionen ändern
function SetEnergyOptions() {
    #Log Func
    LogWrite "FUNC | 07 Energieoptionen werden eingestellt -> keine Logs" 

    # T - brauchts das? > zu testzwecken deaktiviert
    #$powerPlan = Get-WmiObject -Namespace root\cimv2\power -Class Win32_PowerPlan -Filter "ElementName = 'High Performance'"
    #$powerPlan.Activate()

    powercfg -change -monitor-timeout-ac 10 
    powercfg -change -standby-timeout-ac 0
    powercfg -change -hibernate-timeout-ac 0

    powercfg -change -monitor-timeout-dc 10 
    powercfg -change -standby-timeout-dc 0
    powercfg -change -hibernate-timeout-dc 0

    #Schnellstart deaktivieren
    powercfg /hibernate off

    #T - prüfen ob einstellungen vorgenommen wurden
}

#08 Remote Dekstop aktivieren
function ActivateRDP() {
    #Log Func
    LogWrite "FUNC | 08 Remote Desktop wird aktiviert -> keine Logs"

    #RegKey
    $RegistryPath = "HKLM:\System\CurrentControlSet\Control\Terminal Server"
    try { Set-ItemProperty -Path $RegistryPath -Name "fDenyTSConnections" –Value 0 }
    catch {
        #Log Error
        LogWrite "ERROR | $RegistryPath konnte nicht angepasst werden"
        Write-Verbose "Error:" $_
    }

    #Firewall Rule
    try { Enable-NetFirewallRule -DisplayGroup "Remotedesktop" }
    catch {
        #Log Error
        LogWrite "ERROR | Firewallregel konnte nicht aktiviert werden"
        Write-Verbose "Error:" $_
    }
}

#09 OneDrive aus Autostart entfernen
function RemoveOneDrive() {
    #Log Func
    LogWrite "FUNC | 09 OneDrive wird aus Autostart entfernt"

    $RegistryPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run"

    if (Test-Path "$RegistryPath\OneDrive") {
        try { Remove-ItemProperty -Name 'OneDrive' -Path $RegistryPath }
        catch {
            #Log Error
            LogWrite "ERROR | One Drive konnte nicht aus Autostart entfernt werden"
            Write-Verbose "Error:" $_
        }
    }
    else {
        # Log Info
        LogWrite "INFO | OneDrive erfolgreich aus Autostart entfernt"
    }   
}

#10 Systemeigenschaften anpassen
function SetSystemProperties() {
    #Log Func
    LogWrite "FUNC | 10 Systemeigenschaften werden angepasst"

    $RegistryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\"

    If (!(Test-Path "$RegistryPath\OEMInformation")) {
        try { New-Item -Path $RegistryPath -Name OEMInformation }
        catch {
            #Log Warning
            LogWrite "WARN | '$RegistryPath\OEMInformation' RegKey konnte nicht erstellt werden "
            Write-Verbose "Error:" $_
        }
    }

    $SysProps = @{
        Manufacturer = "Root Service AG"
        SupportHours = "von 08:00 bis 17:00"
        SupportPhone = "Tel.: +41 71 634 80 40"
        SupportURL = "https://get.teamviewer.com/zaddm3n"
    }

    foreach ($prop in $SysProps.GetEnumerator()){
        $name = $prop.Key
        $value = $prop.Value

        try { New-ItemProperty -Path $RegistryPath -Name $name -Value $value -PropertyType String -Force }
        catch {
            #Log Error
            LogWrite "WARN | Property '$name' konnte nicht erstellt werden"
            Write-Verbose "Error:" $_
        }
    }
}

#11 NumLock automatisch setzen
function ActivateNumLock() {
    #Log Func
    LogWrite "FUNC | 11 NumLock wird Fest eingestellt"

    $RegistryPaths = @(
        "HKCU:\Control Panel\Keyboard",
        "Microsoft.PowerShell.Core\Registry::HKU\.DEFAULT\Control Panel\Keyboard"
    )

    foreach ($RegistryPath in $RegistryPaths){
        If (!(Test-Path $RegistryPath)) {
            try { New-Item -Path $RegistryPath -Name Keyboard }
            catch {
                #Log Warning
                LogWrite "WARN | '$RegistryPath\Keyboard' konnte nicht erstellt werden"
                Write-Verbose "Error:" $_
            }
    
            try { New-ItemProperty -Path $RegistryPath -Name InitialKeyboardIndicators -Value 2 }
            catch {
                #Log Warning
                LogWrite "WARN | '$RegistryPath\InitialKeyboardIndicators' konnte nicht erstellt werden"
                Write-Verbose "Error:" $_
            }
        }
        If ((Get-ItemProperty -Path $RegistryPath -Name InitialKeyboardIndicators).InitialKeyboardIndicators -gt 100) {
            If (!((Get-ItemProperty -Path $RegistryPath -Name InitialKeyboardIndicators).InitialKeyboardIndicators = 2)) {
                try { New-ItemProperty -Path $RegistryPath -Name InitialKeyboardIndicators -Value 2 }
                catch {
                    #Log Warning
                    LogWrite "WARN | '$RegistryPath\InitialKeyboardIndicators' (Value 2) konnte nicht erstellt werden"
                    Write-Verbose "Error:" $_
                }
            } 
        } 
        Else {
            If (!((Get-ItemProperty -Path $RegistryPath -Name InitialKeyboardIndicators).InitialKeyboardIndicators = 2147483650)) {
                try { New-ItemProperty -Path $RegistryPath -Name InitialKeyboardIndicators -Value 2147483650 }
                catch {
                    #Log Warning
                    LogWrite "WARN | '$RegistryPath\InitialKeyboardIndicators' (Value 2147483650) konnte nicht erstellt werden"
                    Write-Verbose "Error:" $_
                }
            } 
        }
    }

    #Numlock aktivieren
    if (-not [console]::NumberLock) { 
        $w = New-Object -ComObject WScript.Shell; 
        $w.SendKeys('{NUMLOCK}') 
    }
}

#12 Sprachleiste ausblenden, nur DE-CH
function RemoveLanguagebar() {
    #Log Func
    LogWrite "FUNC | 12 Sprachleiste wird eingestellt"

    $list = (Get-WinUserLanguageList) | Where-Object LanguageTag -eq de-CH
    Set-WinUserLanguageList $list -Force
    try { Set-WinUserLanguageList $list -Force }
    catch {
        #Log Warning
        LogWrite "WARN | Neue Languagelist konnte nicht gesetzt werden."
        Write-Verbose "Error:" $_
    }

    #logging
    $newList = (Get-WinUserLanguageList) | Select-Object -ExpandProperty LanguageTag
    If ($newList -eq "de-CH") {
        #Log Info
        LogWrite "INFO | Sprachleiste wurde entfernt, nur de-CH übrig"
    } 
    else {
        #Log Error
        LogWrite "ERROR | Sprachleiste wurde NICHT entfernt | [newList=$newList]"
    } 
}

#13 Fernwartung auf Desktops und in C:/root/ reinkopieren
function GetTeamViewer() {
    #Log Func
    LogWrite "FUNC | 13 Fernwartung wird Heruntergeladen" 

    #Sicherhehen das Root dir existiert
    If(-Not(Test-Path "C:\root")){
        mkdir "C:\root" -Force
    }

    #Fernwartung herunterladen
    $source = "https://customdesign.teamviewer.com/download/version_15x/zaddm3n_windows/TeamViewerQS.exe"
    $destinationArr = @("C:\Users\Public\Desktop\Root Service AG - Fernwartung.exe", "C:\root\Root Service AG - Fernwartung.exe")
    foreach ($destination in $destinationArr) {
        try { Invoke-WebRequest $source -OutFile $destination }
        catch {
            #Log Warning
            LogWrite "ERROR | Fernwartung konnte nicht heruntergeladen werden"
            Write-Verbose "Error:" $_
        }
    }

    #Logging
    If (Test-Path "C:\Users\Public\Desktop\Root Service AG - Fernwartung.exe") {
        #Log Info
        LogWrite "INFO | Fernwartung auf Public Desktop gespeichert"
    } 
}

#14 Vorinstallierten Windows Müll entfernen
function RemoveDefaultSoftware() {
    #Log Func
    LogWrite "FUNC | 14 Vorinstallierter Windows Müll wird entfernt -> keine Logs" 

    #CSV mit Standard-Software angaben einlesen
    try { $RemoveApps = Import-Csv ".\RemoveApps.csv" }
    catch {
        #Log Error
        LogWrite "ERROR | CSV konnte nicht importiert werden."
        Write-Verbose "Error:" $_
    }

    foreach ($app in $RemoveApps.Name){
        #Entfernen
        try { Get-appxpackage -allusers *$app* | Remove-AppxPackage }
        catch {
            #Log Warning
            LogWrite "WARN | $app konnte nicht deinstalliert werden (Remove-AppxPackage)"
            Write-Verbose "Error:" $_

            #Get App Object
            $appObj = Get-WmiObject -Class Win32_Product | Where-Object {
                $_.Name -match $app
            }
            #deinstallieren
            try { $appObj.Uninstall() }
            catch {
                #Log Error
                LogWrite "ERROR | $app konnte nicht deinstalliert werden ($appObj.Uninstall())"
                Write-Verbose "Error:" $_
            }
        }
    }

    #Restliche Software loggen um rauszufinden was noch ins "RemoveApps.csv" gehört
    (Get-AppxPackage | Select-Object Name) | Out-File "$logdir\RemainingPackages.csv"
}

#16 Alle Kacheln entfernen
function RemoveTiles() {
    #Log Func
    LogWrite "FUNC | 16 Alle Kacheln werden entfernt" 

    #Delete layout file if it already exists
    If (Test-Path C:\Windows\StartLayout.xml) {
        
        try { Remove-Item C:\Windows\StartLayout.xml }
        catch {
            #Log Warning
            LogWrite "WARN | Startlayout konnte nicht entfernt werden."
            Write-Verbose "Error:" $_
        }
    }

    #Creates the blank layout file
    echo "<LayoutModificationTemplate xmlns:defaultlayout=""http://schemas.microsoft.com/Start/2014/FullDefaultLayout"" xmlns:start=""http://schemas.microsoft.com/Start/2014/StartLayout"" Version=""1"" xmlns=""http://schemas.microsoft.com/Start/2014/LayoutModification"">" >> C:\Windows\StartLayout.xml
    echo "  <LayoutOptions StartTileGroupCellWidth=""6"" />" >> C:\Windows\StartLayout.xml
    echo "  <DefaultLayoutOverride>" >> C:\Windows\StartLayout.xml
    echo "    <StartLayoutCollection>" >> C:\Windows\StartLayout.xml
    echo "      <defaultlayout:StartLayout GroupCellWidth=""6"" />" >> C:\Windows\StartLayout.xml
    echo "    </StartLayoutCollection>" >> C:\Windows\StartLayout.xml
    echo "  </DefaultLayoutOverride>" >> C:\Windows\StartLayout.xml
    echo "</LayoutModificationTemplate>" >> C:\Windows\StartLayout.xml

    $regAliases = @("HKLM", "HKCU")

    #Assign the start layout and force it to apply with "LockedStartLayout" at both the machine and user level
    foreach ($regAlias in $regAliases) {
        $basePath = $regAlias + ":\SOFTWARE\Policies\Microsoft\Windows"
        $keyPath = $basePath + "\Explorer" 
        IF (!(Test-Path -Path $keyPath)) { 
            New-Item -Path $basePath -Name "Explorer" 
        } 
        Set-ItemProperty -Path $keyPath -Name "LockedStartLayout" -Value 1
        Set-ItemProperty -Path $keyPath -Name "StartLayoutFile" -Value "C:\Windows\StartLayout.xml" 
    } 

    #Restart Explorer, open the start menu (necessary to load the new layout), and give it a few seconds to process
    Stop-Process -name explorer 
    Start-Sleep -s 5 
    $wshell = New-Object -ComObject wscript.shell; $wshell.SendKeys('^{ESCAPE}')
    Start-Sleep -s 5 
    #Enable the ability to pin items again by disabling "LockedStartLayout"
    foreach ($regAlias in $regAliases) {
        $basePath = $regAlias + ":\SOFTWARE\Policies\Microsoft\Windows"
        $keyPath = $basePath + "\Explorer" 
        Set-ItemProperty -Path $keyPath -Name "LockedStartLayout" -Value 0 
    } 

    #Restart Explorer and delete the layout file
    Stop-Process -name explorer 
    Remove-Item C:\Windows\StartLayout.xml 
}

#17 Administrator aktivieren
function ActivateAdmin() {
    #Log Func
    LogWrite "FUNC | 17 Administrator aktivieren"

    #Random pw generieren
    Add-Type -AssemblyName System.Web
    $PassComplexCheck = $false
    $pwLength = 12 # Hier kann die gewünschte PW länge angepasst werden
    do {
        $SecurePassword=[System.Web.Security.Membership]::GeneratePassword($pwLength,1)
        If ( ($SecurePassword -cmatch "[A-Z\p{Lu}\s]") `
            -and ($SecurePassword -cmatch "[a-z\p{Ll}\s]") `
            -and ($SecurePassword -match "[\d]") `
            -and ($SecurePassword -match "[^\w]")
        )
        {
            $PassComplexCheck=$True
        }
    } While ($PassComplexCheck -eq $false)

    $SecurePassword = ConvertTo-SecureString -String $SecurePassword

    #PW setzen
    try { Get-LocalUser -Name "Administrator" | Set-LocalUser -Password $SecurePassword }
    catch {
        #Log Warning
        LogWrite "WARN | Konnte PW für Administrator nicht setzen"
        Write-Verbose "Error:" $_
    }
    
    #User aktivieren
    try { Get-LocalUser -Name "Administrator" | Enable-LocalUser }
    catch {
        #Log Warning
        LogWrite "WARN | Konnte Administrator nicht aktivieren"
        Write-Verbose "Error:" $_
    }
}

#[] Script beenden
function EndScript() {
    #Log Func
    LogWrite "FUNC | [] Schlussmeldung"
    
    # Schlussmeldung
    Write-Host "Folgende Einstellungen müssen von Hand vorgenommen werden: " -ForegroundColor Yellow
    Write-Host "- App Update im App Store ausschalten falls store vorhanden" -ForegroundColor Magenta
    Write-Host "- Taskleiste Alle Symbole" -ForegroundColor Cyan
    Write-Host "- Standardbrowser und Adobe für Pdf setzen" -ForegroundColor Magenta
    Write-Host "- Office einrichten und aktvieren" -ForegroundColor Cyan
    Write-Host "- Windows- und Treiberupdates installieren falls noch nicht gemacht" -ForegroundColor Magenta
    Write-Warning "- Admin und Rs PW setzen und in TPM eintragen!!!"
}

# ---------------------------------------------------- Ausführung ----------------------------------------------------

# On every Run:
TestInternet

# install Windows if Key given
if ($null -ne $Key){
    ActivateWindows
}

# Konfigure User only
if ($User) {
    #Nur User einrichten

    #Aestethics
    ShowDesktopIcons
    RemoveTiles
    RemoveLanguagebar

    #Software
    RemoveDefaultSoftware
    RemoveOneDrive

    #Settings
    ActivateNumLock

    #Schlussmeldung
    EndScript
}
# Option to run only one function
elseif ($null -ne $Func) {

    #Falls parameter mitgegeben wurden -> entsprechende funktion ausführen
    foreach ($function in $Func) {
        try {&$function}
        catch {
            #Log Error
            LogWrite "ERROR | $function konnte nicht aufgerufen werden"
            Write-Verbose "Error:" $_
        }
    }
} 
# Standardkonfiguration
elseif ($null -ne $PCName) {
    #Falls Script ohne Parameter gestartet wurde -> normal Benutzer/Gerät einrichten

    #Gerät
    RenameDisk
    RenameDevice
    DelWindowsOld

    #Benutzer
    DelDefaultUsers
    ActivateAdmin

    #Software
    RemoveDefaultSoftware
    RemoveOneDrive
    InstallStandardApps

    #Settings
    SetEnergyOptions
    ActivateRDP
    ActivateNumLock

    #For Root Service AG
    SetSystemProperties
    GetTeamViewer

    #Aestethics
    ShowDesktopIcons
    RemoveTiles
    RemoveLanguagebar

    #Schlussmeldung
    EndScript
}
else {
    #Log Error
    LogWrite "Error | Fehlende Parameter. Für Standard-Installation muss mindestens der PCName angegeben werden"
}

# Execution policy
Set-ExecutionPolicy Restricted