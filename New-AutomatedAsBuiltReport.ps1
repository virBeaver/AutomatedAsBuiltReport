<#
.SYNOPSIS
    New-AutomatedAsBuiltReport.ps1 - A script to run AsBuiltReport in an automated way.
.DESCRIPTION
    This script is used to run the community project "AsBuiltReport" in an automated way.
    The script askes the user for the vCenter and corresponding credentials and the creates a new AsBuilt report based on predefined configurations (JSON).
    The predefined configurations are stored in the "Config" directory under the current directory.
.INPUTS
    $VIServer:      Name/IP address of the vCenter Server.
    $VICredentials: Credentials for $VIServer.
    $OutPutPath:    Path to create the AsBuiltReport in.
.OUTPUTS
    New AsBuiltReport as Microsoft Word file.
.NOTES
    Author:     Tim Maier
    E-Mail:     tim.maier@icloud.com
    Blog:       https://virbeaver.com
    Twitter:    @virBeaver
    Year:       2019
#>

#Set variables
$OutputPath = ".\"
$AsBuiltConfigPath = ".\Config\AsBuiltConfig.json"


#Trust repository "PSGallery"
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

#Check if NuGet is installed
Write-Host  "Check if NuGet is already installed..." -ForegroundColor Yellow
if (!(Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
    Write-Host "NuGet not installed, installing NuGet now..." -ForegroundColor Yellow
    Install-PackageProvider -Name NuGet -Force | Out-Null
    Write-Host "NuGet not installed, installing NuGet now... done" -ForegroundColor Yellow 
} else {
    Write-Host "NuGet already installed... done" -ForegroundColor Yellow
}

#Check if VMware.PowerCLI is installed
Write-Host  "Check if VMware.PowerCLI is already installed..." -ForegroundColor Yellow
if (!(Get-Module -Name VMware.PowerCLI -ListAvailable)) {
    Write-Host "VMware.PowerCLI not installed, installing VMware.PowerCLI now..." -ForegroundColor Yellow
    Install-Module -Name VMware.PowerCLI -Force
    Write-Host "VMware.PowerCLI not installed, installing VMware.PowerCLI now... done" -ForegroundColor Yellow 
} else {
    Write-Host "VMware.PowerCLI already installed, updating VMware.PowerCLI now..." -ForegroundColor Yellow
    Update-Module -Name VMware.PowerCLI
    Write-Host "VMware.PowerCLI already installed, updating VMware.PowerCLI now... done" -ForegroundColor Yellow
}

#Check if AsBuiltReport is installed
Write-Host  "Check if AsBuiltReport is already installed..." -ForegroundColor Yellow
if (!(Get-Module -Name AsBuiltReport -ListAvailable)) {
    Write-Host "AsBuiltReport not installed, installing AsBuiltReport now..." -ForegroundColor Yellow
    Install-Module -Name AsBuiltReport -Force
    Write-Host "AsBuiltReport not installed, installing AsBuiltReport now... done" -ForegroundColor Yellow
} else {
    Write-Host "AsBuiltReport already installed, updating AsBuiltReport now..." -ForegroundColor Yellow
    Update-Module -Name AsBuiltReport 
    Write-Host "AsBuiltReport already installed, updating AsBuiltReport now... done" -ForegroundColor Yellow
}

#Ask user for input (vCenter Server and credentials)
$VIServer = Read-Host "vCenter Server to connect to"
$VICredential = Get-Credential -Message "Please enter the credentials for $VIServer"

#Check if vCenter connection can be established
Write-Host "Check if connection to $VIServer can be established..." -ForegroundColor Yellow
Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $true -Confirm:$false | Out-Null
Connect-VIServer -Server $VIServer -Credential $VICredential -ErrorAction:SilentlyContinue 
if ($global:DefaultVIServers.count -eq 1 -and $global:DefaultVIServers[0].name -eq $VIServer) {
    Disconnect-VIServer -Server $VIServer -Confirm:$false
    #Run AsBuiltReport with predefined configuration
    Write-Host "Connected to $VIServer, running AsBuiltReport now..." -ForegroundColor Yellow
    New-AsBuiltReport -Target $VIServer -Credential $VICredential -Format Word  -Report VMware.vSphere -EnableHealthCheck -OutputPath $OutputPath -AsBuiltConfigPath $AsBuiltConfigPath
    Write-Host "Connected to $VIServer, running AsBuiltReport now... done" -ForegroundColor Yellow
} else {
    Read-Host "Unable to connect to $VIServer with given credentials, press any key to return..."
    Return
}