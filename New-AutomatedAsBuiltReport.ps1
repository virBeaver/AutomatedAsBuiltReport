<#
.SYNOPSIS
    New-AutomatedAsBuiltReport.ps1 - A script to run AsBuiltReport in an automated way.
.DESCRIPTION
    This script is used to run the community project "AsBuiltReport" in an automated way.
    The script askes the user for the vCenter and corresponding credentials and then creates a new AsBuiltReport report based on predefined configurations (JSON).
    The predefined configurations are stored in the "Config" directory under the current directory.
.INPUTS
    $VIServer:      Name/IP address of the vCenter Server.
    $VICredentials: Credentials for $VIServer.
.OUTPUTS
    New AsBuiltReport as Microsoft Word file.
.NOTES
    Author:     Tim Neukirchen
    E-Mail:     neukirchentim@gmail.com
    Blog:       https://virbeaver.com
    Twitter:    @virBeaver
    Year:       2020
#>

#Allow execution of unsigned scripts
Set-ExecutionPolicy Unrestricted -Force

#Ignore invalid certificates while connecting to vCenter
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false

#Set predefined variables
$OutputPath = ".\"
$AsBuiltConfigPath = ".\Config\AsBuiltConfig.json"

#Check if NuGet is installed
Write-Host  "Check if NuGet is already installed..." -ForegroundColor Yellow
if (!(Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
    Write-Host "NuGet not installed, installing NuGet now..." -ForegroundColor Yellow -NoNewline
    Install-PackageProvider -Name NuGet -Force | Out-Null
    Write-Host "`rNuGet not installed, installing NuGet now... successful" -ForegroundColor Yellow 
} else {
    Write-Host "NuGet already installed... successful" -ForegroundColor Yellow
}

#Trust repository "PSGallery"
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

#Check if VMware.PowerCLI is installed
Write-Host  "Check if VMware.PowerCLI is already installed..." -ForegroundColor Yellow
if (!(Get-Module -Name VMware.PowerCLI -ListAvailable)) {
    Write-Host "VMware.PowerCLI not installed, installing VMware.PowerCLI now..." -ForegroundColor Yellow -NoNewline
    Install-Module -Name VMware.PowerCLI -Force
    Write-Host "`rVMware.PowerCLI not installed, installing VMware.PowerCLI now... successful" -ForegroundColor Yellow 
} else {
    Write-Host "VMware.PowerCLI already installed, updating VMware.PowerCLI now..." -ForegroundColor Yellow -NoNewline
    Update-Module -Name VMware.PowerCLI
    Write-Host "`rVMware.PowerCLI already installed, updating VMware.PowerCLI now... successful" -ForegroundColor Yellow
}

#Check if AsBuiltReport is installed
Write-Host  "Check if AsBuiltReport is already installed..." -ForegroundColor Yellow
if (!(Get-Module -Name AsBuiltReport -ListAvailable)) {
    Write-Host "AsBuiltReport not installed, installing AsBuiltReport now..." -ForegroundColor Yellow -NoNewline
    Install-Module -Name AsBuiltReport -Force
    Write-Host "`rAsBuiltReport not installed, installing AsBuiltReport now... successful" -ForegroundColor Yellow
} else {
    Write-Host "AsBuiltReport already installed, updating AsBuiltReport now..." -ForegroundColor Yellow -NoNewline
    Update-Module -Name AsBuiltReport 
    Write-Host "`rAsBuiltReport already installed, updating AsBuiltReport now... successful" -ForegroundColor Yellow
}

#Ask user for input (vCenter Server and credentials)
$VIServer = Read-Host "vCenter Server to connect to"
$VICredential = Get-Credential -Message "Please enter the credentials for $VIServer"

#Check if vCenter connection can be established
Write-Host "Check if connection to $VIServer can be established..." -ForegroundColor Yellow -NoNewline
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $true -Confirm:$false | Out-Null
Connect-VIServer -Server $VIServer -Credential $VICredential -ErrorAction:SilentlyContinue | Out-Null
if ($global:DefaultVIServers.count -eq 1 -and $global:DefaultVIServers[0].name -eq $VIServer) {
    Disconnect-VIServer -Server $VIServer -Confirm:$false
    Write-Host "`rCheck if connection to $VIServer can be established... successful" -ForegroundColor Yellow
    #Run AsBuiltReport with predefined configuration
    Write-Host "Running AsBuiltReport now..." -ForegroundColor Yellow
    New-AsBuiltReport -Target $VIServer -Credential $VICredential -Format Word  -Report VMware.vSphere -EnableHealthCheck -OutputPath $OutputPath -AsBuiltConfigPath $AsBuiltConfigPath
    Write-Host "`rRunning AsBuiltReport now... successful" -ForegroundColor Yellow
} else {
    Read-Host "`rCheck if connection to $VIServer can be established... failed [Press ENTER to return]"
    Return
}