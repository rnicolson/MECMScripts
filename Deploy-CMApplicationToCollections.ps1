<#
.SYNOPSIS
    Creates multiple  deployments to a group of collections using parameters imported from a CSV. See the DECRIPTION for an example of the CSV format.

.DESCRIPTION
    Creates multiple deployments to a group of collections imported from a CSV. The CSV needs to have a single Column 
    with the following headers. 
    Name	
    AppName	
    Available	
    Deadline	
    Action	
    Purpose	
    AllowRepair	
    RebootOutsideSW	
    TimeBasedOn	
    OverrideSW	
    Notification

    An example of the data in the rows is below. DO NOT include the quotes. They are just there to make it easier to see the separation for each column
    "Collection Name"	"Application Name"	"2/22/2023 10:00"	"2/22/2025 22:30"	"Install"	"Available"	"FALSE"	"FALSE"	"LocalTime"	"TRUE"	"DisplaySoftwareCenterOnly"

.PARAMETER outputFile
    Specifies the name and path for the CSV-based input file. 
    Required?                    true
    Position?                    0
    Default value
    Accept pipeline input?       false

.INPUTS
    Requires a path for the CSV to set all the required settings for the deployments.

.OUTPUTS
    New Config Manager deployments based on the parameters set in the CSV

.EXAMPLE
    powershell.exe -ExecutionPolicy ByPass -File .\Deploy-CMApplicationToCollectionsParam.ps1 -CollectionListPath "C:\Temp\CollectionsForDeployment.csv"


.NOTES
    Author - Ryan Nicolson
    Last Revised - 02/28/2023
    ChangeLog - Production version release

#>
param (
    [Parameter(Mandatory=$true)]
    [string]$collectionListPath
)
#MECM Site Information
$SiteCode = "XXX"
$SiteServer = "server.domain.com"

#############################Script##################################
#Connect to MECM site
Import-Module -Name "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
if($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
New-PSDrive -name $SiteCode -psprovider CMSite  -Root $SiteServer 
}
Set-Location "$($SiteCode):\"

#Import the CSV file
$collections = Import-Csv -Path $collectionListPath

#Step through the CSV creating deployments for each row
foreach ($collection in $Collections) {
    #Get details about the collection
    $collectionObject = Get-CMCollection -Name $collection.Name
    #Set all boolean values imported as strings to booleans
    $AllowRepairBool = [System.Convert]::ToBoolean($collection.AllowRepair)
    $RebootOutsideSWBool = [System.Convert]::ToBoolean($collection.RebootOutsideSW)
    $OverrideSWBool = [System.Convert]::ToBoolean($collection.OverrideSW)
    #Set the parameters for the deployment required for this row
    $DeploymentParams = @{
        Name                       = $collection.AppName
        AvailableDateTime          = $collection.Available
        DeadlineDateTime           = $collection.Deadline
        Collection                 = $collectionObject
        DeployAction               = $collection.Action
        DeployPurpose              = $collection.Purpose
        AllowUserRepairApplication = $AllowRepairBool
        RebootOutsideServiceWindow = $RebootOutsideSWBool
        TimeBaseOn                 = $collection.TimeBasedOn
        OverrideServiceWindow      = $OverrideSWBool
        UserNotification           = $collection.Notification
    }
    #Create the deployment
    New-CMApplicationDeployment @DeploymentParams
}
