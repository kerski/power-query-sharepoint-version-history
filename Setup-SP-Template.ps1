<#
.SYNOPSIS
Performs installation of Version History Template - Bronze; a Power BI Dataflow
.DESCRIPTION
    Description: This script:
        1) Upload sample Power BI Dataflow to the Power BI workspace
        Dependencies: 
            1) Power BI Powershell installed
            2) ImportModel.ps1
         
    Author: John Kerski
.INPUTS
None.
.OUTPUTS
None.
.EXAMPLE
PS> .\Setup-SP-Template.ps1
#>
### GLOBAL VARIABLES
$Branch = "development"
$SharePointURLKeyword = "\|SHAREPOINT_URL\|"
$ListNameKeyword = "\|LIST_NAME\|"
$ListName = $null
$SPSiteURL = $null
$DF_URL = ""
$DFOutput = ".\Version History - Bronze.json"
$DFUtilsURI = "https://raw.githubusercontent.com/kerski/powerbi-powershell/master/examples/dataflows/DFUtils.psm1"
$GraphURI = "https://raw.githubusercontent.com/kerski/powerbi-powershell/master/examples/dataflows/Graph.psm1"
$ImportURI = "https://raw.githubusercontent.com/kerski/powerbi-powershell/master/examples/dataflows/ImportModel.ps1"

### UPDATE VARIABLES HERE thru Read-Host
# Set Workspace Name
$WorkspaceName = Read-Host "Please enter the name of the Power BI Workspace"
# Paste the URL of the SharePoint list
$Location = Read-Host "Please past the URL of the SharePoint list."

$ListNameResults = ($Location | Select-String -Pattern '/Lists/([^/]+)(?:/|$)' -AllMatches)

if (!$ListNameResults)
{
    Throw "We could not extract the list name from the URL $($Location)."} else {
    $ListName = $ListNameResults.Matches.Groups[1].Value
    $SPSiteURL = $Location.SubString(0,$ListNameResults.Matches.Groups[1].Index - 7)
}

Write-Host -ForegroundColor Cyan "Updating the Power BI dataflow template for the $($ListName) list in $($SPSiteURL)"

# Download Dependencies
Write-Host -ForegroundColor Cyan "Downloading scripts and dataflow template from GitHub."
#Download scripts for Graph, DFUtils, and Import-Module.ps1
Invoke-WebRequest -Uri $DFUtilsURI -OutFile "./DFUtils.psm1"
Invoke-WebRequest -Uri $GraphURI -OutFile "./Graph.psm1"
Invoke-WebRequest -Uri $ImportURI -OutFile "./ImportModel.ps1"

#Download file
$FileLocation = ".\SharePoint - Version History Template - Bronze.json"

(Get-Content $FileLocation) -replace $SharePointURLKeyword, $SPSiteURL | Set-Content $DFOutput -Force
(Get-Content $DFOutput) -replace $ListNameKeyword, $ListName | Set-Content $DFOutput -Force

#Install Powershell Module if Needed
if (Get-Module -ListAvailable -Name "MicrosoftPowerBIMgmt") {
    Write-Host "MicrosoftPowerBIMgmt installed moving forward"
} else {
    #Install Power BI Module
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -AllowClobber -Force
}
#Login into Power BI to Get Workspace Information
Login-PowerBI

$WS = Get-PowerBIWorkspace -Name $WorkspaceName

if(!$WS){
    Throw "Unable to retrieve the workspace information for the workspace provided: $($WorkspaceName)"
}

#Upload Bronze Data Flow
.\ImportModel.ps1 -Workspace $WS.name -File $DFOutput
