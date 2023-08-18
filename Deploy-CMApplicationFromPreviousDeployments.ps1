<#
.SYNOPSIS
    Creates new deployments for an application based on the existing deployments to a previous application.

.DESCRIPTION
    Creates new deployments for an application based on the existing deployments to a previous application. You will enter the name
    of the previous application and the name of the new application and the deployments will be replicated on the new application.
    If there are some deploymnets that are required instead of available you will need to provide the new deadline date for the deployment. 

.PARAMETER oldAppName
  Specifies the name of the previous application that you want to replicate the deployments for.

.PARAMETER newAppName
  Specifies the name of the new application that you want deployed to the same collections as the previous application.

.PARAMETER daysUntilRequired
  Specifies the number of days from today's date that you want the deadline for any required applications to be. For example if you run the script
  January 1, 2023 at 10:00 AM and specify daysUntilRequired to be 4, the deadline for required deployments will be January 5, 2023 at 10:00 AM.

.PARAMETER sameDeadlineForAll
  If there are multiple deployments that will be set to required you can specify this parameter to set them all to have the same deadline.
  If you don't add this parameter you will be prompted to confirm they won't all have the same deadline. If you specify that they will not then 
  you can enter the deadline for each deployment while the script runs.

.INPUTS
    Nothing can be piped into the script but you must provide the old application name and the new application names. If any of the 
    previous deployments are required you will also need to provide a new deadline date either as a parameter of the number of 
    days from the current date as an offset or by entering the date for each deployment.

.OUTPUTS
    New Config Manager deployments based on the dealine dates provided by offset of for each deployment. Each deployment is
    based off the settings of the previous deployments.

.EXAMPLE
    powershell.exe -ExecutionPolicy ByPass -File "C:\temp\Deploy-CMApplicationFromPreviousDeployments.ps1" -oldAppName '7-Zip 22.01' -newAppName '7-Zip 23.01' -daysUntilRequired '3' -sameDeadlineForAll


#>
param (
    [Parameter(Mandatory = $true)]
    [string]$oldAppName,
    [Parameter(Mandatory = $true)]
    [string]$newAppName,
    [Parameter(Mandatory = $false)]
    [int]$daysUntilRequired,
    [Parameter(Mandatory = $false)]
    [Switch]$sameDeadlineForAll
)

function Get-DeadlineDate {
    param (
        [Parameter(Mandatory = $false)]
        [Switch[]]$sameDeadlineForAll,
        [Parameter(Mandatory = $false)]
        [string[]]$collectionName
    )
    if (!$sameDeadlineForAll) {
        $promptValue = "Please specify the deadline date for the deployment to $collectionName"
    }
    else {
        $promptValue = "Please specify the deadline date for all deployments."
    }
    $validDate = $false
    while (!$validDate) {
        $deadlineDate = Read-Host -Prompt $promptValue
        $deadlineDate = $deadlineDate -as [DateTime]
        if (!$deadlineDate) {  
            write-host -ForegroundColor Red 'You entered an invalid date'
            $validDate = $false
        }
        else {
            $validDate = $true
        }
    }
    Return $deadlineDate
}


#MECM Site Information
$SiteCode = "XX1"
$SiteServer = "sitename.domain.blah"

#############################Config Manager Connection##################################

#Connect to MECM site
Import-Module -Name "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
if ($null -eq (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
    New-PSDrive -name $SiteCode -psprovider CMSite  -Root $SiteServer 
}
Set-Location "$($SiteCode):\"

##############################Variables#####################
$logLocation = "c:\$newAppName-DistLog.txt"
$collections = Get-CMApplicationDeployment -Name $oldAppName
$newAppDetails = Get-CMApplication -Name $newAppName
$numberOfReqCollections = 0
$hasRequiredCollections = $false
Start-Transcript -Path $logLocation
#############################Script##################################

if ($newAppDetails.IsDeployed -eq $false) {
    Write-Host "Application $newAppName was not previously deployed to a distribution point. Distributing to primary site distribution group"
    Write-Host "You may see an error here if the application was recently distributed and it has not fully completed the process."
    Start-CMContentDistribution -ApplicationName $newAppName -DistributionPointGroupName "MECM Primary Site Distribution Group"
}

$date = Get-Date
if ($PSBoundParameters.ContainsKey('sameDeadlineForAll') -ne $false) {
    $sameDeadlineForAll = $true
}
else {
    $sameDeadlineForAll = $false
}

#Setting the dealines for installation for required installs
if ($PSBoundParameters.ContainsKey('daysUntilRequired') -ne $false) {
    $deadlineDate = $date.AddDays($daysUntilRequired)
    $deadlineSetByOffset = $true
}
else {
    $deadlineSetByOffset = $false
}

foreach ($collection in $collections) {
    if ($collection.OfferTypeID -eq 0) {
        $hasRequiredCollections = $true
        $numberOfReqCollections = $numberOfReqCollections + 1
    }
}

if ($PSBoundParameters.ContainsKey('sameDeadlineForAll') -eq $false -and $hasRequiredCollections) {
    $validResponse = $false
    while (!$validResponse) {
        Write-Host -ForegroundColor Red "$numberOfReqCollections deployments are set to required but you did not specify if all required deployments should have the same deadline."
        $sameDeadlineInput = Read-Host -Prompt "Should all deployments have the same deadline? Y or N?"

        if ($sameDeadlineInput -eq "y" -or $sameDeadlineInput -eq "Y" -or $sameDeadlineInput -eq "N" -or $sameDeadlineInput -eq "n") {  
            $validResponse = $true
        }
        else {
            write-host -ForegroundColor Red 'You entered an invalid option'
            $validResponse = $false
        }
    }
    switch ($sameDeadlineInput) {
        "Y" { $sameDeadlineForAll = $true }
        "y" { $sameDeadlineForAll = $true }
        "N" { $sameDeadlineForAll = $false }
        "n" { $sameDeadlineForAll = $false }
    }
}

if ($hasRequiredCollections -and $sameDeadlineForAll -and !($deadlineSetByOffset)) {
    $deadlineDate = Get-DeadlineDate -sameDeadlineForAll $sameDeadlineForAll
}

foreach ($collection in $collections) {
    $collectionName = $collection.CollectionName
    $collectionObject = Get-CMCollection -Name $collection.CollectionName
    switch ($collection.OfferTypeID) {
        0 { $depPurpose = "Required" }
        2 { $depPurpose = "Available" }
        Default { $depPurpose = "Available" }
    }
    if ($depPurpose -eq "Available") {
        $DeploymentParams = @{
            Name                       = $newAppName
            AvailableDateTime          = $date
            Collection                 = $collectionObject
            DeployAction               = "Install"
            DeployPurpose              = $depPurpose
            AllowUserRepairApplication = $False
            TimeBaseOn                 = "LocalTime"
            UserNotification           = "DisplaySoftwareCenterOnly"
        }
    }
    else {
        if (!$sameDeadlineForAll) {
            $deadlineDate = Get-DeadlineDate -collectionName $collectionName
        }
        
        $DeploymentParams = @{
            Name                       = $newAppName
            AvailableDateTime          = $date
            DeadlineDateTime           = $deadlineDate
            Collection                 = $collectionObject
            DeployAction               = "Install"
            DeployPurpose              = $depPurpose
            AllowUserRepairApplication = $False
            RebootOutsideServiceWindow = $collection.RebootOutsideOfServiceWindows
            TimeBaseOn                 = "LocalTime"
            OverrideServiceWindow      = $collection.OverrideServiceWindows
            UserNotification           = "DisplaySoftwareCenterOnly"
        }
    }

    Write-Output "*****************************"
    Write-Output "Deploying to $collectionName"
    Write-Output $DeploymentParams
    New-CMApplicationDeployment @DeploymentParams
}

Stop-Transcript
