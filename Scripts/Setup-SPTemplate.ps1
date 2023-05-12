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
PS> .\Setup-SPTemplate.ps1
#>
### GLOBAL VARIABLES
$Branch = "development"
$SharePointURLKeyword = "\|SHAREPOINT_URL\|"
$ListNameKeyword = "\|LIST_NAME\|"
$ListName = $null
$SPSiteURL = $null
$DF_URL = ""
$DFOutput = ".\Version History - Bronze.json"
$DFUtilsURI = "https://raw.githubusercontent.com/kerski/power-query-sharepoint-version-history/$($Branch)/Scripts/DFUtils.psm1"
$GraphURI = "https://raw.githubusercontent.com/kerski/power-query-sharepoint-version-history/$($Branch)/Scripts/Graph.psm1"
$ImportURI = "https://raw.githubusercontent.com/kerski/power-query-sharepoint-version-history/$($Branch)/Scripts/ImportModel.ps1"
$TemplateURI = "https://raw.githubusercontent.com/kerski/power-query-sharepoint-version-history/$($Branch)/SharePoint%20-%20Version%20History%20Template%20-%20Bronze.json"
$FileLocation = "./SharePoint - Version History Template - Bronze.json"
### Install dependencies
#Install Powershell Module if Needed
if (Get-Module -ListAvailable -Name "MicrosoftPowerBIMgmt") {
    Write-Host "MicrosoftPowerBIMgmt installed moving forward"
} else {
    #Install Power BI Module
    Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -AllowClobber -Force
}

if (Get-Module -ListAvailable -Name "PnP.PowerShell") {
    Write-Host "PnP.PowerShell installed moving forward"
} else {
    #Install Power BI Module
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -AllowClobber -Force
}

### UPDATE VARIABLES HERE thru Read-Host
# Set Workspace Name
$WorkspaceName = Read-Host "Please enter the name of the Power BI Workspace"
# Paste the URL of the SharePoint list
$Location = Read-Host "Please past the URL of the SharePoint list"

$ListNameResults = ($Location | Select-String -Pattern '/Lists/([^/]+)(?:/|$)' -AllMatches)
if (!$ListNameResults)
{
    Throw "We could not extract the list name from the URL $($Location)."} else {   
    $SPSiteURL = $Location.SubString(0,$ListNameResults.Matches.Groups[1].Index - 7)
    # Connect to SharePoint
    Connect-PnPOnline $SPSiteURL -Interactive
    $ListNameSearch = Get-PnPList | Where-Object {$_.DefaultViewUrl.ToString() -like "*/lists/$($ListNameResults.Matches.Groups[1].Value)*"}
    
    if(!$ListNameSearch)
    {
        Throw "We could not locate the list name from the URL $($Location)."
    }#end if

    $ListName = $ListNameSearch[0].Title
} #end if


Write-Host -ForegroundColor Cyan "Updating the Power BI dataflow template for the $($ListName) list in $($SPSiteURL)"

# Download Dependencies
Write-Host -ForegroundColor Cyan "Downloading scripts and dataflow template from GitHub."
#Download scripts for Graph, DFUtils, and Import-Module.ps1
Invoke-WebRequest -Uri $DFUtilsURI -OutFile "./DFUtils.psm1"
Invoke-WebRequest -Uri $GraphURI -OutFile "./Graph.psm1"
Invoke-WebRequest -Uri $ImportURI -OutFile "./ImportModel.ps1"
#Download Template
Invoke-WebRequest -Uri $TemplateURI -OutFile $FileLocation

(Get-Content $FileLocation) -replace $SharePointURLKeyword, $SPSiteURL | Set-Content $DFOutput -Force
(Get-Content $DFOutput) -replace $ListNameKeyword, $ListName | Set-Content $DFOutput -Force

#Login into Power BI to Get Workspace Information
Login-PowerBI

$WS = Get-PowerBIWorkspace -Name $WorkspaceName

if(!$WS){
    Throw "Unable to retrieve the workspace information for the workspace provided: $($WorkspaceName)"
}

#Upload Bronze Data Flow
.\ImportModel.ps1 -Workspace $WS.name -File $DFOutput

Write-Host -ForegroundColor Green "Dataflow Template successfully uploaded to $($WS.name)"