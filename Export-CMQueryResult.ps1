<#
.SYNOPSIS
    This script exports results from an MECM query and saves it into a CSV.

.DESCRIPTION
    This takes the output from a query in MECM exports it to a CSV. 

.PARAMETER requiredCollectionName
    Specifies the name of the MECM collection you want to run inventory for. 
    Required?                    true
    Position?                    0
    Default value
    Accept pipeline input?       false

.PARAMETER queryName
    Specifies the name of the MECM query you want to run inventory for. 
    Required?                    true
    Position?                    1
    Default value
    Accept pipeline input?       false

.PARAMETER outputFile
    Specifies the name and path for the CSV-based output file. 
    Required?                    true
    Position?                    2
    Default value
    Accept pipeline input?       false

.INPUTS
    None

.OUTPUTS
    CSV file of the query results

.EXAMPLE
    Export-CMQueryResults.ps1 -requiredCollectionName 'Collection-Name' -queryName 'Query-Name' -outputFile 'C:\Temp\queryresult.csv'
    powershell.exe -ExecutionPolicy ByPass -File "C:\temp\Export-CMQueryResults.ps1" -requiredCollectionName 'Collection-Name' -queryName 'Query-Name' -outputFile 'C:\Temp\queryresult.csv'

.NOTES
    The script will cycle through each object returned from the Invoke-CMQuery command
    and break it down to individual objects which are then returned from the function.
    Finally it will export the results to a CSV. This script may return errors to the console
    depending on the data returned from the query. This is the result of the 
    $propertyValue = [system.String]::Join(" ", $propertyValue)
    line which is fixing the results of certain items returned from queries. Items like MAC addresses,
    IP addresses or OU information is returned as an object and this command breaks the objects apart
    and converts them to a string which can fit into a single cell in the CSV.

    Author - Ryan Nicolson
    Date - January 19, 2023

#>

#######################Parameters###########################
param (
    [Parameter(Mandatory=$true)]
    [string]$requiredCollectionName,
    [Parameter(Mandatory=$true)]
    [string]$queryName,
    [Parameter(Mandatory=$true)]
    [string]$outputFile
)

#######################Functions##########################
function Optimize-CMQueryResult {
    <#
    .SYNOPSIS
    Function which creates a usable object from the results of an Invoke-CMQuery request

    .DESCRIPTION
    Inputs the queryID and collection ID for the query you want to run. This works best for 
    queries that require a collection name

    .PARAMETER QueryID
    The query ID from MECM

    .PARAMETER CollectionID
    The collection ID from MECM

    .EXAMPLE
    Optimize-CMQueryResult -QueryID $CMQueryID -CollectionID $CMCollectionId
    #>
    param (
        [Parameter(Mandatory=$true)]
        [string]$QueryID,
        [Parameter(Mandatory=$true)]
        [string]$CollectionID
    )
    $computersList = New-Object System.Collections.ArrayList($null)
    $queryout = Invoke-CMQuery -Id $QueryID -LimitToCollectionId $CollectionID
    foreach ($computer in $queryout) {
        foreach ($attribute in $computer) {
            $computerObject = New-Object -TypeName psobject
            $members = $attribute | Get-Member -Type Property
            foreach ($member in $members) {
                if ($member.Name.StartsWith("SMS_") ) {
                    $memberNames = $member| Get-Member -Type Properties
                    foreach ($memberName in $memberNames) {                   
                        if (-not ($memberName.Definition.StartsWith("System.")) -and -not ($memberName.Definition.StartsWith("string Type")) -and -not ($memberName.Definition.StartsWith("string Definition"))) {
                            $currentMemberName = $memberName.Name
                            $currentProperties = $computer.($member.($memberName.($currentMemberName)))
                            foreach ($property in $currentProperties) {
                                $currentPropertiestypes = $currentProperties | Get-Member -Type Properties
                                foreach ($propertyType in $currentPropertiestypes) {
                                    if ($propertyType.Definition.StartsWith("Microsoft")) {
                                        $propertyValueName = $propertyType.Name
                                        $propertyValueFullName = $member.Name + " " + $propertyValueName
                                        $propertyValue = $property.($propertyValueName)
                                        $propertyValue = [system.String]::Join(" ", $propertyValue)
                                        $computerObject | Add-Member -MemberType NoteProperty -Name $propertyValueFullName -Value $propertyValue
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        if ($null -ne $computerObject) {
            $computersList.Add($computerObject) > $null
        }
    }
    Return $computersList
}

##############################Variables#############################

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

#Pull the collection and query information for use in calls to the query
$collection = Get-CMCollection -name $requiredCollectionName
$query = Get-CMQuery -name $queryName

#Pull the query reult for the query
$queryOutput = Optimize-CMQueryResult -QueryID $query.QueryID -CollectionID $collection.CollectionId


#Export the ArrayList to CSV using the provided path
$queryOutput | Export-Csv -Path $outputFile -NoTypeInformation
