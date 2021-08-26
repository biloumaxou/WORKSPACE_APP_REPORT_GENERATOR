## Variables
[string]$scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
[string]$intelligenceReportGeneratorConfigFile = Join-Path -Path $scriptRoot -ChildPath 'BPI_INTELLIGENCE_REPORT_GENERATOR_CONFIG.xml'

## Import variables from XML configuration file
[Xml.XmlDocument]$xmlConfigFile = Get-Content -LiteralPath $intelligenceReportGeneratorConfigFile -Encoding UTF8
[Xml.XmlElement]$xmlConfig = $xmlConfigFile.IntelligenceReportGenerator_Config

## Create Authorization Header Value
[byte[]]$credentialBytes = [System.Text.Encoding]::UTF8.GetBytes("$($xmlConfig.Service_Account.Client_ID):$($xmlConfig.Service_Account.Secret_Key)")
[string]$encodedCredential = [Convert]::ToBase64String($credentialBytes)
[string]$authorizationHeader = "Basic ${encodedCredential}"

## Get Access Token
[hashtable]$headers = @{
    'Authorization' = $authorizationHeader
    'Accept' = 'application/json'
    'Content-Type' = 'application/json'
}
[PSCustomObject]$apiResponse = Invoke-RestMethod -Method Post -Uri $xmlConfig.Service_Account.Token_Endpoint -Headers $headers
[string]$accessToken = $apiResponse.access_token

## Run Report
[hashtable]$headers = @{
    'Authorization' = "Bearer $accessToken"
    'Accept' = 'application/json'
    'Content-Type' = 'application/json'
}
Invoke-RestMethod -Method Get -Uri "https://auth.eu1.data.vmwservices.com/v1/meta/integration/airwatch/entity/device/attributes" -Headers $headers