## Variables
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[string]$scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
[string]$intelligenceReportGeneratorConfigFile = Join-Path -Path $scriptRoot -ChildPath 'BPI_INTELLIGENCE_REPORT_GENERATOR_CONFIG.xml'

## Import variables from XML configuration file
[Xml.XmlDocument]$xmlConfigFile = Get-Content -LiteralPath $intelligenceReportGeneratorConfigFile -Encoding UTF8
[Xml.XmlElement]$xmlConfig = $xmlConfigFile.IntelligenceReportGenerator_Config

## Create Authorization Header Value
[byte[]]$credentialBytes = [System.Text.Encoding]::UTF8.GetBytes("$($xmlConfig.Client_ID):$($xmlConfig.Secret_Key)")
[string]$encodedCredential = [Convert]::ToBase64String($credentialBytes)
[string]$authorizationHeader = "Basic $encodedCredential"

#################################################################################
## STEP 1 - Get Access Token                                                   ##
#################################################################################
[hashtable]$headers = @{
    'authorization' = $authorizationHeader
    'accept' = 'application/json'
    'content-type' = 'application/json'
}
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Post -Uri $xmlConfig.Token_Endpoint -Headers $headers
[string]$accessToken = $apiResponse.access_token

#################################################################################
## STEP 2 - Get Targeted Account Service Details                               ##
#################################################################################
[hashtable]$headers = @{
    'authorization' = "Bearer $accessToken"
    'accept' = 'application/json'
    'content-type' = 'application/json'
}
[PSCustomObject]$body = @{
    search_terms = @(
        @{
            value = 'true'
            fields = @('active')
        }
        @{
            value = $xmlConfig.Client_ID
            fields = @('display_name')
        }
    )
}
$body = $body | ConvertTo-Json -Depth 3
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Post -Uri "$($xmlConfig.Root_URL)/v1/account/search" -Headers $headers -Body $body

If ($apiResponse.data.total_count -eq 1) {
    # API Return Only One Result, Which Is Expected
    If ($apiResponse.data.results.user_descriptor.directory_type -eq 'SERVICE_ACCOUNT') {
        [PSCustomObject]$targetServiceAccountDetails = $apiResponse.data.results
    } Else {
        Throw "An Error Occured When Validating Account Service Details Due To Directory Type Is Not A Service Account (Directory Type: $($apiResponse.data.results.user_descriptor.directory_type))."
    }
} Else {
    Throw "An Error Occured When Searching Account Service Details Due To Incorrect Total Count Number Returned By Workspace ONE Intelligence (Total Count: $($apiResponse.data.total_count))."
}

#################################################################################
## STEP 3 - Ask for Targeted Report Name                                       ##
#################################################################################
[string]$confirmation = $null
While ($confirmation -ne 'y') {
    [string]$reportName = $null
    $reportName = Read-Host -Prompt 'Enter The Name Of The Report You Would Like To Execute And Download'
    $confirmation = Read-Host -Prompt "You entered [$reportName]. Is this correct ? (y/n)"
}

#################################################################################
## STEP 4 - Get Targeted Report Details                                        ##
#################################################################################
[PSCustomObject]$body = @{
    search_terms = @(
        @{
            value = "adobeconnect"
            fields = @('name')
        }
    )
}
# Convert Body To JSON
$body = $body | ConvertTo-Json -Depth 3
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Post -Uri "$($xmlConfig.Root_URL)/v2/reports/search" -Headers $headers -Body $body

If ($apiResponse.data.total_count -eq 1) {
    # API Return Only One Result, Which Is Expected
        [PSCustomObject]$targetReportDetails = $apiResponse.data.results
} Else {
    Throw "An Error Occured When Searching Account Service Details Due To Incorrect Total Count Number Returned By Workspace ONE Intelligence (Total Count: $($apiResponse.data.total_count))."
}

#################################################################################
## STEP 5 - Verify If Targeted Report Is Sharing With Targeted Service Account ##
## This Is Mandatory To be Able To Run Reports                                 ##
#################################################################################
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Get -Uri "$($xmlConfig.Root_URL)/v1/reports/$($targetReportDetails.id)/share/accounts" -Headers $headers
[string]$isReportSharedWithAccount = [string]::IsNullOrEmpty(($apiResponse.data.details.account_id | Where-Object {$_ -eq $($targetServiceAccountDetails.id)}))

If ($isReportSharedWithAccount) {
    # Targeted Report Is Not Shared With Targeted Service Account
    $body = ConvertTo-Json -InputObject @(
        @{
            "user_descriptor" = @{
                "id" = $targetServiceAccountDetails.id
            }
            "account_access_level" = $xmlConfig.Account_Access_Level
        } ) -Depth 3
    [PSCustomObject]$apiResponse = Invoke-RestMethod -Method Put -Uri "$($xmlConfig.Root_URL)/v1/reports/$($targetReportDetails.id)/share" -Headers $headers -Body $body
}

#################################################################################
## STEP 6 - Run The Report                                                     ##
#################################################################################
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method POST -Uri "$($xmlConfig.Root_URL)/v1/reports/$($targetReportDetails.id)/run" -Headers $headers
[PSCustomObject]$targetReportRequestDetails = $apiResponse.data

#################################################################################
## STEP 7 - Tracking Report Generation State                                   ##
#################################################################################
[bool]$continueToSearch = $true
[bool]$errorOnSearch = $false
[int]$offsetValue = 0
[int]$pageSizeValue = $xmlConfig.API_Page_Size


while($continueToSearch){
    [PSCustomObject]$body = @{
        offset = $offsetValue
        page_size = $pageSizeValue
    }
    # Convert Body To JSON
    $body = $body | ConvertTo-Json
    [PSCustomObject]$apiResponse = Invoke-RestMethod -Method POST -Uri "$($xmlConfig.Root_URL)/v1/reports/$($targetReportDetails.id)/downloads/search" -Headers $headers -Body $body
    [PSCustomObject]$targetReportRequestDetails = $apiResponse.data.results | Where-Object {$_.report_schedule_id -eq $($targetReportRequestDetails.id)}
    If (-not [string]::IsNullOrEmpty($targetReportRequestDetails)) {
        # Found The Right Target Report Request Details
        $continueToSearch = $false
    } Elseif ($apiResponse.data.total_count -gt ($apiResponse.data.page_size + $apiResponse.data.offset)) {
        # Need To Read The Next Page Of Results
        $offsetValue += $pageSizeValue
    } Else {
        # No Results Founds
        $continueToSearch = $false
        $errorOnSearch = $true
    }
}

Write-Host $errorOnSearch

<#
#region Variables - YOU HAVE TO MODIFY THIS SECTION!

#You can get authentication info creating a service account in Intelligence ( https://techzone.vmware.com/getting-started-workspace-one-intelligence-apis-workspace-one-operational-tutorial#_1107619 ). 

#Set the auth endpoint according to your tenant's region
$authEndpoint = 'https://auth.eu1.data.vmwservices.com/oauth/token?grant_type=client_credentials'

#The authorization header is the clientId:clientSecret in Base64 format.
$authHeader = @{
    'Content-Type' = 'application/json'
    'Authorization' = 'Basic pasteYourClientId:ClientSecretInBase64FormatHere'
}
#There's no need to modify the body
$body = @{
    "offset" = 0;
    "page_size" = 100;
}
#Choose a name for your report
$reportName = "Asset_Inventory_"
#Where to download the report, select an already existing folder
$downloadFolder = "C:\WS1\Reports\"

#You can easily get $reportGUID from the desired report webpage URL in Intelligence, it's the GUID after "/list/" and before "/overview"
#THIS REPORT HAS TO BE SHARED WITH YOUR SERVICE ACCOUNT IN INTELLIGENCE
$reportGUID = "a51cac58-b7db-48b4-82f5-c46ed82fc89e"

#Set these variables according to your tenant's region
$reportingUri = "https://api.eu1.data.vmwservices.com/v1/reports/"
$reportingDownloadUri = "https://api.eu1.data.vmwservices.com/v1/reports/tracking/"
#endregion

###### DO NOT MODIFY ANTYTHING BELOW THIS LINE ######
#####################################################

#Obtaining the access token
$accessToken = Invoke-RestMethod -Method Post -Uri $authEndpoint -Headers $authHeader -ContentType 'application/x-www-form-urlencoded'

#Building the authorization headers
$headers = @{}
$headers.Add("Authorization","$($accessToken.token_type) "+" "+"$($accessToken.access_token)")

#Getting the report download id
$dataUri = $reportingUri + $reportGUID + "/downloads/search"

$data = Invoke-RestMethod -Method Post -Headers $headers -Uri $dataUri -Body ($body | ConvertTo-Json) -ContentType "application/json"

$lastReport =  $data.data.results.GetValue(0)
$reportId = $lastReport.id

#Getting the report generation date and creating the download destination
$reportDate = $lastReport.created_at.Substring(0,16).Replace(':','-').Replace('T','_')
$reportFile = $downloadFolder + $reportName + $reportDate + ".csv"

#Creating the report download Uri using the reportId
$downloadUri = $reportingDownloadUri + $reportId + '/download'

#Downloading the report in the specified location
Invoke-RestMethod -Method Get -Headers $headers -Uri $downloadUri -OutFile $reportFile



#>