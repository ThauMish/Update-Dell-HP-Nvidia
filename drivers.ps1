# Creer une instance pour  Update Service Manager
$UpdateSvc = New-Object -ComObject Microsoft.Update.ServiceManager

# Fais la liste des services
$Service = (New-Object -ComObject Microsoft.Update.ServiceManager).Services
$UpdateSvc.AddService2("$Service",7,"")

# Creer une update session
$Session = New-Object -ComObject Microsoft.Update.Session
$Searcher = $Session.CreateUpdateSearcher()

# Set le serviceID pour l'update
$Searcher.ServiceID = '$Service'
$Searcher.SearchScope =  1 # MachineOnly
$Searcher.ServerSelection = 3 # Third Party

# Les critères de recherches
$Criteria = "IsInstalled=0 and Type='Driver'"
Write-Host('Recherche update drivers...') -Fore Green
$SearchResult = $Searcher.Search($Criteria)

# List qui match les critèes
$Updates = $SearchResult.Updates

# Affiche les drivers
$Updates | select Title, DriverModel, DriverVerDate, Driverclass, DriverManufacturer | fl

# Create la collection d'update a telecharger
$UpdatesToDownload = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { $UpdatesToDownload.Add($_) | out-null }
Write-Host('Telecharge les drivers...')  -Fore Green

$UpdateSession = New-Object -Com Microsoft.Update.Session

$Downloader = $UpdateSession.CreateUpdateDownloader()

# Set les updates a DL
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download()

# Creer les la collections d'update a installer
$UpdatesToInstall = New-Object -Com Microsoft.Update.UpdateColl
$updates | % { if($_.IsDownloaded) { $UpdatesToInstall.Add($_) | out-null } }
Write-Host('Installation des drivers...')  -Fore Green

# Install les updates
$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToInstall
$InstallationResult = $Installer.Install()

Write-Host('MAJ drivers terminer, passage MAJ Nvidia')

# Ce script permet la detection automatique de la carte nvidia et installe le dernier drivers depuis le site officiel

# Installer options
param (
    [switch]$clean = $false, # efface ancien driver et remplace par le nv
    [string]$folder = "$env:temp" # emplacement driver
)

# Check 7zip

$7zipinstalled = $false 
if ((Test-path HKLM:\SOFTWARE\7-Zip\) -eq $true) {
    $7zpath = Get-ItemProperty -path  HKLM:\SOFTWARE\7-Zip\ -Name Path
    $7zpath = $7zpath.Path
    $7zpathexe = $7zpath + "7z.exe"
    if ((Test-Path $7zpathexe) -eq $true) {
        $archiverProgram = $7zpathexe
        $7zipinstalled = $true 
    }    
}

else {
    while ($choice -notmatch "[o|n]") {
        $choice = read-host "Pour continuer l'installation vous avez besoin de 7-zip, voulez-vous le telecharger ? Tapper "o" pour oui, "n" pour non"
    }
    if ($choice -eq "o") {

        $7zip = "https://www.7-zip.org/a/7z1900-x64.exe"
        $output = "$PSScriptRoot\7Zip.exe"
        (New-Object System.Net.WebClient).DownloadFile($7zip, $output)
       
        Start-Process "7Zip.exe" -Wait -ArgumentList "/S"
        # Delete the installer
        Remove-Item "$PSScriptRoot\7Zip.exe"
        & $MyInvocation.MyCommand.Definition
    }
    else {
        Write-Host "Press any key to exit..."
        $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit
    }
}

# Check la version driver installé
Write-Host "Verification de la version du driver installe"
try {
    $VideoController = Get-WmiObject -ClassName Win32_VideoController | Where-Object { $_.Name -match "NVIDIA" }
    $ins_version = ($VideoController.DriverVersion.Replace('.', '')[-5..-1] -join '').insert(3, '.')
}
catch {
    Write-Host -ForegroundColor Yellow "La carte Nvidia n'as pas pu etre detecter..."
    Write-Host "Appuyez sur une touche pour quitter..."
    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}
Write-Host "Version installe`t$ins_version"

# Check le dernier driver sur nvidia
$link = Invoke-WebRequest -Uri 'https://www.nvidia.com/Download/processFind.aspx?psid=101&pfid=816&osid=57&lid=1&whql=1&lang=en-us&ctk=0&dtcid=1' -Method GET -UseBasicParsing
$link -match '<td class="gridItem">([^<]+?)</td>' | Out-Null
$version = $matches[1]
Write-Host "Derniere version disponible `t`t$version"

# Compare le driver installer et la version nvidia
if (!$clean -and ($version -eq $ins_version)) {
    Write-Host "La meme version est deja installe"
    Write-Host "Appuiez sur une touche pour quitter..."
    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Check la version de windows
if ([Environment]::OSVersion.Version -ge (new-object 'Version' 9, 1)) {
    $windowsVersion = "win10-win11"
}
else {
    Write-Host "Version windows pas compatible"
}

# Check windows 64bit
if ([Environment]::Is64BitOperatingSystem) {
    $windowsArchitecture = "64bit"
}
else {
    Write-Host "Version 32 bits pas dispo"
}

# dossier telechargement nvidia
$nvidiaTempFolder = "$folder\NVIDIA"
New-Item -Path $nvidiaTempFolder -ItemType Directory 2>&1 | Out-Null

# Genere le lien
$url = "https://international.download.nvidia.com/Windows/$version/$version-desktop-$windowsVersion-$windowsArchitecture-international-dch-whql.exe"
$rp_url = "https://international.download.nvidia.com/Windows/$version/$version-desktop-$windowsVersion-$windowsArchitecture-international-dch-whql-rp.exe"

# telecharge l'installer
$dlFile = "$nvidiaTempFolder\$version.exe"
Write-Host "Telechargement de la derniere version $dlFile"
Start-BitsTransfer -Source $url -Destination $dlFile

if ($?) {
    Write-Host "Continuer..."
}
else {
    Write-Host "Le telechargement n'as pas pu aboutir"
    Start-BitsTransfer -Source $rp_url -Destination $dlFile
}

# Extrais les fichiers nvnidia
$extractFolder = "$nvidiaTempFolder\$version"
$filesToExtract = "Display.Driver HDAudio NVI2 PhysX EULA.txt ListDevices.txt setup.cfg setup.exe"
Write-Host "Telechargement fini, extraction des fichiers..."

if ($7zipinstalled) {
    Start-Process -FilePath $archiverProgram -NoNewWindow -ArgumentList "x -bso0 -bsp1 -bse1 -aoa $dlFile $filesToExtract -o""$extractFolder""" -wait
}

else {
    Write-Host "Fail au moment de l'extraction"
    Write-Host "Appuiez sur une touche pour quitter..."
    $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# Enleve les dependances pas necessaires
(Get-Content "$extractFolder\setup.cfg") | Where-Object { $_ -notmatch 'name="\${{(EulaHtmlFile|FunctionalConsentFile|PrivacyPolicyFile)}}' } | Set-Content "$extractFolder\setup.cfg" -Encoding UTF8 -Force

# Install les drivers
Write-Host "Le driver est entrain d'etre installe, patientez 5 mins tous ce passe en arriere plan..."
$install_args = "-passive -noreboot -noeula -nofinish -s"
if ($clean) {
    $install_args = $install_args + " -clean"
}
Start-Process -FilePath "$extractFolder\setup.exe" -ArgumentList $install_args -wait

# Delete les fichiers telechargé
Write-Host "Supprime les fichiers restants"
Remove-Item $nvidiaTempFolder -Recurse -Force

# Driver installé , reboot ?
Write-Host -ForegroundColor Green "Driver installé. reboot pour finir l'installation "
Write-Host "Reboot maintenant ?"
$Readhost = Read-Host "(Y/N) Non par défaut"
Switch ($ReadHost) {
    Y { Write-host "Reboot ..."; Start-Sleep -s 2; Restart-Computer }
    N { Write-Host "Quitte le script dans 5 sec "; Start-Sleep -s 5 }
    Default { Write-Host "Quitte le script dans 5sec"; Start-Sleep -s 5 }
}

exit

