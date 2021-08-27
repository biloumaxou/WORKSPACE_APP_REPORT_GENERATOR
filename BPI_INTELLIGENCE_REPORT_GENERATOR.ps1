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
[array]$targetServiceAccountDetails = $apiResponse.data.results | Where-Object {($_.display_name -eq $xmlConfig.Client_ID) -and ($_.user_descriptor.directory_type -eq 'SERVICE_ACCOUNT')}
If ($targetServiceAccountDetails.Count -ne 1) {
    Throw "An Error Occured When Searching Account Service Details Due To Incorrect Total Count Number Returned By Workspace ONE Intelligence (Total Count: $($targetServiceAccountDetails.Count))."
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
            value = $reportName
            fields = @('name')
        }
    )
}
# Convert Body To JSON
$body = $body | ConvertTo-Json -Depth 3
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Post -Uri "$($xmlConfig.Root_URL)/v2/reports/search" -Headers $headers -Body $body
[array]$targetReportDetails = $apiResponse.data.results | Where-Object {$_.name -eq $reportName}
If ($targetReportDetails.Count -ne 1) {
    Throw "An Error Occured When Searching Report Details Due To Incorrect Total Count Number Returned By Workspace ONE Intelligence (Total Count: $($targetReportDetails.Count))."
}

#################################################################################
## STEP 5 - Verify If Targeted Report Is Sharing With Targeted Service Account ##
## This Is Mandatory To be Able To Run Reports                                 ##
#################################################################################
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Get -Uri "$($xmlConfig.Root_URL)/v1/reports/$($targetReportDetails.id)/share/accounts" -Headers $headers
[bool]$isReportNotSharedWithAccount = [string]::IsNullOrEmpty(($apiResponse.data.details.account_id | Where-Object {$_ -eq $($targetServiceAccountDetails.id)}))

If ($isReportNotSharedWithAccount) {
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
[bool]$isReportRun = $true
[bool]$SupportShortDelay = $true
While ($isReportRun) {
    [bool]$continueToSearch = $true
    [int]$offsetValue = 0

    While ($continueToSearch){
        [PSCustomObject]$body = @{
            offset = $offsetValue
            page_size = $xmlConfig.API_Page_Size
        }
        # Convert Body To JSON
        $body = $body | ConvertTo-Json
        [PSCustomObject]$apiResponse = Invoke-RestMethod -Method POST -Uri "$($xmlConfig.Root_URL)/v1/reports/$($targetReportDetails.id)/downloads/search" -Headers $headers -Body $body
        [PSCustomObject]$targetReportTrackingDetails = $apiResponse.data.results | Where-Object {$_.report_schedule_id -eq $($targetReportRequestDetails.id)}
        If (-not [string]::IsNullOrEmpty($targetReportTrackingDetails)) {
            # Found The Right Target Report Request Details
            $continueToSearch = $false
            Break
        } Elseif ($apiResponse.data.total_count -gt ($apiResponse.data.page_size + $apiResponse.data.offset)) {
            # Need To Read The Next Page Of Results
            $offsetValue += $xmlConfig.API_Page_Size
        } Else {
            # No Results Founds
            $continueToSearch = $false
        }
    }
    If ($targetReportTrackingDetails) {
        If ($targetReportTrackingDetails.status -eq 'COMPLETED') {
            # Report Is Generated
            $isReportRun = $false
        } Else {
            # Report still running
            Start-Sleep -Seconds 15
        }
    } Else {
        If ($SupportShortDelay) {
            $SupportShortDelay = $false
            Start-Sleep -Seconds 30
        } Else {
            Throw 'An Error Occured When Searching Report Completion Status. No Report Were Found'
        }
    }
}

#################################################################################
## STEP 8 - Download Report                                                    ##
#################################################################################
$reportFile = $scriptRoot + '\' + $reportName + ".csv"
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Get -Uri "$($xmlConfig.Root_URL)/v1/reports/tracking/$($targetReportTrackingDetails.id)/download" -Headers $headers -OutFile $reportFile