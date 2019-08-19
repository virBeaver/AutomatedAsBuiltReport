<#
.SYNOPSIS
    New-AutomatedAsBuiltReport.ps1 - A script to run AsBuiltReport in an automated way.
.DESCRIPTION
    This script is used to run the community project "AsBuiltReport" in an autmated way.
    The script askes the user for the vCenter and corresponding credentials and the creates a new AsBuilt report based on predefined configurations (JSON).
.INPUTS
    $VIServer:      Name/IP address of the vCenter Server.
    $VICredentials: Credentials for $VIServer.
    $OutPutPath:    Path to create the AsBuiltReport in.
.OUTPUTS
    New AsBuiltReport as Microsoft Word file.
.NOTES
    Author:     Tim Maier
    E-Mail:     tim.maier@icloud.com
    Twitter:    @virBeaver
    Year:       2019
#>

#Set variables
$OutputPath = ".\"
$AsBuiltConfigPath = ".\Config\AsBuiltReport.Config.json"


#Trust repository "PSGallery"
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

#Check if VMware.PowerCLI is installed
if (!(Get-Module -Name VMware.PowerCLI -ListAvailable)) {
    Install-Module -Name VMware.PowerCLI -Force
} else {
    Update-Module -Name VMware.PowerCLI
}

#Check if AsBuiltReport is installed
if (!(Get-Module -Name AsBuiltReport -ListAvailable)) {
    Install-Module -Name AsBuiltReport -Force
} else {
    Update-Module -Name AsBuiltReport 
}

#Asked user for input (vCenter Server and credentials)
$VIServer = Read-Host -Prompt "vCenter Server to connect to"
$VICredential = Get-Credential -Message "Please enter the credentials for $VIServer"

#Run AsBuiltReport with predefined configuration
New-AsBuiltReport -Target $VIServer -Credential $VICredential -Format Word  -Report VMware.vSphere -EnableHealthCheck -OutputPath $OutputPath -AsBuiltConfigPath $AsBuiltConfigPath