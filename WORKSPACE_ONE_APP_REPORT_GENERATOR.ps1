##*=============================================
##* VARIABLE DECLARATION
##*=============================================
#region VariableDeclaration
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[string]$envUserDesktop = [Environment]::GetFolderPath('DesktopDirectory')
[System.Collections.ArrayList]$logElements = New-Object -TypeName 'System.Collections.ArrayList'
[string]$scriptRoot = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
[string]$wsoneAppReportConfigFile = Join-Path -Path $scriptRoot -ChildPath 'BPI_INTELLIGENCE_REPORT_GENERATOR_CONFIG.xml'
[int]$script:arrayID = 0

## Verify If XML Configuration File Exist
If (-not (Test-Path -Path $wsoneAppReportConfigFile -PathType Leaf)) {
    Throw 'Workspace ONE App Report XML configuration file not found.'
}

## Import variables from XML configuration file
[Xml.XmlDocument]$xmlConfigFile = Get-Content -LiteralPath $wsoneAppReportConfigFile -Encoding UTF8
[Xml.XmlElement]$xmlConfig = $xmlConfigFile.WorkspaceONEAppReport_Config
##  Get UEM Authentication Details
[Xml.XmlElement]$xmlUemAuthenticationDetails = $xmlConfig.UEM_Authentication
[string]$uemTokenEndpoint = $xmlUemAuthenticationDetails.UEM_TokenEndpoint
[string]$uemClientID = $xmlUemAuthenticationDetails.UEM_ClientID
[string]$uemSecretKey = $xmlUemAuthenticationDetails.UEM_SecretKey
[string]$uemRootURL = $xmlUemAuthenticationDetails.UEM_RootURL
##  Get Intelligence Authentication Details
[Xml.XmlElement]$xmlIntelligenceAuthenticationDetails = $xmlConfig.Intelligence_Authentication
[string]$IntelligenceTokenEndpoint = $xmlIntelligenceAuthenticationDetails.Intelligence_TokenEndpoint
[string]$IntelligenceClientID = $xmlIntelligenceAuthenticationDetails.Intelligence_ClientID
[string]$IntelligenceSecretKey = $xmlIntelligenceAuthenticationDetails.Intelligence_SecretKey
[string]$IntelligenceRootURL = $xmlIntelligenceAuthenticationDetails.Intelligence_RootURL
##  Get Tool Configuration Details
[Xml.XmlElement]$xmlToolDetails = $xmlConfig.Tool_Config
[string]$intelligenceAccountAccessLevel = $xmlToolDetails.Intelligence_Account_Access_Level
[int]$apiPageSize = $xmlToolDetails.API_Page_Size
[string]$reportDirectoryLocation = $ExecutionContext.InvokeCommand.ExpandString($xmlToolDetails.Reports_Location)
#endregion
##*=============================================
##* END VARIABLE DECLARATION
##*=============================================

##*=============================================
##* FUNCTION LISTINGS
##*=============================================
#region FunctionListings
#region Function Get-AccessToken
Function Get-AccessToken {
    [CmdletBinding()]
	Param (
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$ClientID,
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$ClientSecret,
        [Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$AccessTokenURL
	)

    Try {
        [hashtable]$headers = @{
            'accept' = 'application/json'
            'content-type' = 'application/x-www-form-urlencoded'
        }

        [hashtable]$body = @{
            'grant_type' = 'client_credentials'
            'client_id' = $ClientID
            'client_secret' = $ClientSecret
        }

        [PSCustomObject]$response = Invoke-RestMethod -Method Post -Uri $AccessTokenURL -Headers $headers -Body $body
        If ([string]::IsNullOrEmpty($response.access_token)) {
            Throw 'Access Token Field Is Empty.'
        } Else {
            Write-Output -InputObject $response.access_token
        }
    } Catch {
        Throw "An Error Occured When Getting Access Token. $($_.Exception.Message)"
    }
}
#endregion
#region Function Invoke-WorkspaceOneApi
Function Invoke-WorkspaceOneApi {
    [CmdletBinding()]
	Param (
        [Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$RootURL,
		[Parameter(Mandatory=$false)][ValidateNotNullorEmpty()][string]$BaseURL,
        [Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$AccessToken,
		[Parameter(Mandatory=$false)][ValidateRange('1','4')][string]$VersionAPI = '1',
        [Parameter(Mandatory=$false)][ValidateSet('Get','Put','Post')][string]$Method = 'Get',
        [Parameter(Mandatory=$false)][ValidateNotNullorEmpty()][hashtable]$Parameters,
        [Parameter(Mandatory=$false)][ValidateNotNullorEmpty()][System.Object]$Body,
        [Parameter(Mandatory=$false)][ValidateNotNullorEmpty()][string]$OutFile
	)

    Try {
        ## Build Header
        [hashtable]$headers = @{
            'Authorization' = "Bearer $AccessToken"
            'Accept' = "application/json;version=$VersionAPI"
            'Content-Type' = 'application/json'
        }

        ## Build API URL By Concatenate $RootURL and $BaseURL
        If ($PSBoundParameters.ContainsKey('BaseURL')) {
            [string]$apiURI = Join-PathURL -Path $RootURL -ChildPath $BaseURL
        }
        ## Build API URL By Adding Parameters
        If ($PSBoundParameters.ContainsKey('Parameters')) {
            [System.Collections.Specialized.NameValueCollection]$apiParams = [System.Web.HttpUtility]::ParseQueryString([string]::Empty)
            ForEach($key in $Parameters.keys){
                $apiParams.Add($key,$Parameters[$key])
            }
            [System.Object]$requestURI = [System.UriBuilder]$apiURI
            $requestURI.Query = $apiParams.ToString()
            $apiURI = $requestURI.Uri
        }

        Switch ($Method) {
            'Get' { Write-Output -InputObject (Invoke-RestMethod -Method 'Get' -Uri $apiURI -Headers $headers -OutFile $OutFile) }
            'Put' { Write-Output -InputObject (Invoke-RestMethod -Method 'Put' -Uri $apiURI -Headers $headers -Body $body) }
            'Post' { Write-Output -InputObject (Invoke-RestMethod -Method 'Post' -Uri $apiURI -Headers $headers -Body $body) }
        }
    } Catch {
        Throw "An Error Occured When invoking REST API. $($_.Exception.Message)"
    }
}
#endregion
#region Function Join-PathURL
Function Join-PathURL {
    [CmdletBinding()]
	Param (
        [Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$Path,
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$ChildPath
	)

    Try {
        ## Verify If $Path Not Finish With Value /
        $Path = $Path.Trim()
        While ($Path.EndsWith('/')) {
            $Path = $Path.Substring(0,$Path.length - 1)
        }

        ## Verify If $ChildPath Not Begin With Value / And Not Finish With Value /
        $ChildPath = $ChildPath.Trim()
        While ($ChildPath.EndsWith('/')) {
            $ChildPath = $ChildPath.Substring(0,$ChildPath.length - 1)
        }
        While ($ChildPath.StartsWith('/')) {
            $ChildPath = $ChildPath.Substring(1)
        }

        ## Concatenate $Path And $ChildPath
        [string]$fullPath = $Path + '/' + $ChildPath

        ## Verify If $fullPath Is A Correct URL
        If ((([System.URI]$fullPath).Scheme -match '[http|https]') -and (-not [string]::IsNullOrEmpty(([System.URI]$fullPath).AbsoluteUri))) {
            Write-Output -InputObject $fullPath
        } Else {
            Throw 'An Error Occured On URL Validation.'
        }
    } Catch {
        Throw "An Error Occured When Joining Path And ChildPath URL. $($_.Exception.Message)"
    }
}
#endregion
#region Function Add-Message
Function Add-Message {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()][string]$Message,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$Color = 'Default',
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][bool]$IsTempMessage = $false,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$RemoveTempMessageByGroupID,
        [Parameter(Mandatory=$false)][switch]$RemoveAllTempMessage,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$TempMessageGroupID = 'None'
    )

    Try {
        If ($RemoveTempMessageByGroupID) {
            Remove-Message -GroupID $RemoveTempMessageByGroupID
        } ElseIf ($RemoveAllTempMessage) {
            Remove-Message -RemoveAll
        }
        $null = $logElements.Add([pscustomobject]@{ID = ($script:arrayID -as [string]); Message = $Message; Color = $Color; IsTempMessage = ($IsTempMessage -as [string]); GroupID = $TempMessageGroupID})
        $script:arrayID++
        Show-Message
    } Catch {
        Throw "An Error Occured When Trying To Add Message. $($_.Exception.Message)"
    }
}
#endregion
#region Function Show-Message
Function Show-Message {
    [CmdletBinding()]
    Param ()

    Try {
        Clear-Host

        [int]$count = 0

        Foreach ($logElement in $logElements) {
            $count++
            If ($count -lt $logElements.Count ) {
                [string]$logOutput = "$($logElement.Message)`n"
            } Else {
                [string]$logOutput = $logElement.Message
            }

            Switch ($logElement.Color) {
                'Default' { Write-Host -Object $logElement.Message }
                default { Write-Host -Object $logElement.Message -ForegroundColor $logElement.Color }
            }
        }
    } Catch {
        Throw "An Error Occured When Trying To Show Message. $($_.Exception.Message)"
    }
}
#endregion
#region Function Remove-Message
Function Remove-Message {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][string]$GroupID,
        [Parameter(Mandatory=$false)][ValidateNotNullOrEmpty()][switch]$RemoveAll
    )

    Try {
        If ($RemoveAll) {
            While ($logElements.IsTempMessage.IndexOf('True') -ne -1) {
                $null = $logElements.RemoveAt($logElements.IsTempMessage.IndexOf('True'))
            }
        } ElseIf ($PSBoundParameters.ContainsKey('GroupID')) {
            [array]$indexElements = @()
            ForEach ($arrayElement in $logElements) {
                If ($arrayElement.IsTempMessage -eq 'True' -and $arrayElement.GroupID -eq $GroupID) {
                    $indexElements += $arrayElement.ID
                }
            }
            If ($indexElements.Count -gt 0) {
                ForEach ($indexElement in $indexElements) {
                    $null = $logElements.RemoveAt($logElements.ID.IndexOf($indexElement))
                }
            }
        } Else {
            Throw 'An Error Occured When Trying To Remove Message.'
        }
    } Catch {
        Throw "An Error Occured When Trying To Remove Message. $($_.Exception.Message)"
    }
}
#endregion
#region Function New-Menu
Function New-Menu {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)][String]$MenuTitle,
        [Parameter(Mandatory=$True)][array]$MenuOptions
    )

    Try {
        [int]$MaxValue = $MenuOptions.count-1
        [int]$Selection = 0
        [bool]$EnterPressed = $False

        While(-not $EnterPressed) {
            Add-Message -Message "$MenuTitle" -IsTempMessage $True -RemoveTempMessageByGroupID 'None'

            For ($i=0; $i -le $MaxValue; $i++){
                If ($i -eq $Selection){
                    Add-Message -Message "[ $($MenuOptions[$i]) ]" -Color 'Cyan' -IsTempMessage $True
                } Else {
                    Add-Message -Message "  $($MenuOptions[$i])  " -IsTempMessage $True
                }
            }

            $KeyInput = $host.ui.rawui.readkey("NoEcho,IncludeKeyDown").virtualkeycode

            Switch($KeyInput){
                13{
                    $EnterPressed = $True
                    Return $Selection
                    break
                }
                38{
                    If ($Selection -eq 0){
                        $Selection = $MaxValue
                    } Else {
                        $Selection -= 1
                    }
                    Remove-Message -GroupID 'None'
                    break
                }
                40{
                    If ($Selection -eq $MaxValue){
                        $Selection = 0
                    } Else {
                        $Selection +=1
                    }
                    Remove-Message -GroupID 'None'
                    break
                }
                Default {
                    Remove-Message -GroupID 'None'
                }
            }
        }
    } Catch {
        Throw "An Error Occured When Trying To Create A New Menu. $($_.Exception.Message)"
    }
}
#endregion
#endregion
##*=============================================
##* END FUNCTION LISTINGS
##*=============================================

##*=============================================
##* SCRIPT BODY
##*=============================================
#region ScriptBody
Add-Message -Message '#########################################################'
Add-Message -Message '####       Welcome To Workspace One App Report       ####'
Add-Message -Message '#########################################################'

################################################
## STEP 1: Get UEM And Intelligence Access Token
################################################
Add-Message -Message "`n<-- Step 1 - Get Access Token -->"
## Get Workspace ONE UEM Access Token API
Add-Message -Message 'Workspace ONE UEM Access Token Retrieval. Please Wait...' -Color 'Yellow' -IsTempMessage $true
[string]$uemAccessToken = Get-AccessToken -ClientID $uemClientID -ClientSecret $uemSecretKey -AccessTokenURL $uemTokenEndpoint
Add-Message -Message 'Workspace ONE UEM Access Token Retrieved' -Color 'Green' -RemoveAllTempMessage
## Get Workspace ONE Intelligence Access Token API
Add-Message -Message 'Workspace ONE Intelligent Access Token Retrieval. Please Wait...' -Color 'Yellow' -IsTempMessage $true
[string]$intelligenceAccessToken = Get-AccessToken -ClientID $IntelligenceClientID -ClientSecret $IntelligenceSecretKey -AccessTokenURL $IntelligenceTokenEndpoint
Add-Message -Message 'Workspace ONE UEM Intelligent Token Retrieved' -Color 'Green' -RemoveAllTempMessage

################################################
## STEP 2: Get App Name
################################################
Add-Message -Message "`n<-- Step 2 - Choose Targeted App With Its Name -->"
[string]$confirmation = '1'
While ($confirmation -ne '0') {
    [string]$appName = $null
    Add-Message -Message 'Please Enter An App Name For Which You Want To Generate A Report' -IsTempMessage $true -RemoveAllTempMessage
    $appName = (Read-Host).Trim()
    $confirmation = New-Menu -MenuTitle "You entered [$appName] As App Name. Is this correct ?" -MenuOptions 'Yes','No'
}
Add-Message -Message "You Choose To Generate A Report For App Containing [$appName] Name" -IsTempMessage $true -TempMessageGroupID 'appdetails' -RemoveAllTempMessage

################################################
## STEP 3: Search App Details Based On App Name
################################################
[hashtable]$Params = @{
    'applicationtype' = 'Internal'
    'type' = 'App'
    'model' = 'Desktop'
    'applicationname' = $appName
}
[PSCustomObject]$uemAppDetails = Invoke-WorkspaceOneApi -RootURL $uemRootURL -BaseURL 'API/mam/apps/search' -AccessToken $uemAccessToken -Parameters $Params
If ($uemAppDetails.Total -le 0) {
    ## No App Is Found Matching Provided App Name
    Add-Message -Message "No App Containing [$appName] Name Has Been Found From Workspace ONE UEM" -Color 'Red'
} Else {
    ## At Least 1 App Is Found Matching Provided App Name
    Add-Message -Message "[$($uemAppDetails.Total)] App(s) Detected Matching With [$appName] App Name Has Been Found From Workspace ONE UEM" -IsTempMessage $true -TempMessageGroupID 'appdetails'
    [array]$options = @()
    [hashtable]$appDetails = @{}
    [int]$keyValue = 0
    ForEach ($uemAppDetail in $uemAppDetails.Application) {
        $appDetails.Add($keyValue, @{Id = $uemAppDetail.Id.Value; Uuid = $uemAppDetail.Uuid; ApplicationName = $uemAppDetail.ApplicationName; BundleId = $uemAppDetail.BundleId; AppVersion = $uemAppDetail.AppVersion})
        $null = $keyValue++
        $options += "Application Name: $($uemAppDetail.ApplicationName) | Version: $($uemAppDetail.AppVersion) | ID: $($uemAppDetail.Id.Value)"
    }
    [int]$confirmation = New-Menu -MenuTitle "Please Select The Desired App" -MenuOptions $options
    $appName = $appDetails[$confirmation].ApplicationName
    Add-Message -Message "You Choose To Generate A Report For [$($appDetails[$confirmation].ApplicationName)] App In Version [$($appDetails[$confirmation].AppVersion)] Identified By The ID [$($appDetails[$confirmation].Id)]" -Color 'Green' -RemoveAllTempMessage
    Add-Message -Message "The Intelligence Report Related With [$($appDetails[$confirmation].ApplicationName)] App Will Be Named [$($appDetails[$confirmation].ApplicationName)_$($appDetails[$confirmation].Id)]" -Color 'Green' -RemoveAllTempMessage
    [string]$reportName = $null
    $reportName = "$($appDetails[$confirmation].ApplicationName)_$($appDetails[$confirmation].Id)".Replace(' ','_')

################################################
## STEP 4: Get Intelligence Targeted Account Service Details
################################################
    Add-Message -Message "`n<-- Step 3 - Get Account Service Details -->"
    [PSCustomObject]$body = @{
        search_terms = @(
            @{
                value = 'true'
                fields = @('active')
            }
            @{
                value = $IntelligenceClientID
                fields = @('display_name')
            }
        )
    }
    [System.Object]$body = $body | ConvertTo-Json -Depth 3
    [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL 'v1/account/search' -AccessToken $intelligenceAccessToken -Method 'Post' -Body $body
    [array]$targetServiceAccountDetails = $apiResponse.data.results | Where-Object {($_.display_name -eq $IntelligenceClientID) -and ($_.user_descriptor.directory_type -eq 'SERVICE_ACCOUNT')}
    If ($targetServiceAccountDetails.Count -ne 1) {
        Throw "An Error Occured When Searching Account Service Details Due To Incorrect Total Count Number Returned By Workspace ONE Intelligence (Total Count: $($targetServiceAccountDetails.Count))."
    }
    Add-Message -Message "Account Service Details Are Retrived Successfully" -Color 'Green'

################################################
## STEP 5: Get intelligence Targeted Report Details Or Create Report
################################################
    Add-Message -Message "`n<-- Step 4 - Get Intelligence Report Details -->"
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
    [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL 'v2/reports/search' -AccessToken $intelligenceAccessToken -Method 'Post' -Body $body
    [PSCustomObject]$targetReportDetails = $apiResponse.data.results | Where-Object {$_.name -eq $reportName}
    If (([array]$targetReportDetails).Count -ne 1) {
        Add-Message -Message "The Report [$reportName] Not Exist yet. Create Report" -Color 'Yellow' -IsTempMessage $true
        ## Generate New Report
        [PSCustomObject]$body = @{
            name = $reportName
            description = "Get All Device Details For Which $appName is installed"
            filter = " airwatch.application.app_id = $($appDetails[$confirmation].Id)  AND  airwatch.application.app_is_installed = true  AND  airwatch.device.device_enrollment_status IN ( 'Enrolled' )  AND  airwatch.device._event_created_time WITHIN 60 days "
            integration = "airwatch"
            entity = "application"
            join_entities_by_integration = @{
                airwatch = @("application","device","user")
            }
            column_names = @("airwatch.application._app_name","airwatch.application._app_version","airwatch.application.app_id","airwatch.application.app_install_date","airwatch.device.device_location_group_name","airwatch.device._device_hostname","airwatch.device.device_id","airwatch.device.device_model","airwatch.device.device_manufacturer_name","airwatch.device._device_os_version","airwatch.user.device_enrollment_user_first_name","airwatch.user.device_enrollment_user_last_name","airwatch.user.device_enrollment_user_name","airwatch.user.device_enrollment_user_email")
        }
        # Convert Body To JSON
        $body = $body | ConvertTo-Json -Depth 3
        [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL 'v2/reports' -AccessToken $intelligenceAccessToken -Method 'Post' -Body $body
        [PSCustomObject]$targetReportDetails = $apiResponse.data
        Add-Message -Message "The Report [$reportName] Is Created Successfully" -Color 'Green' -RemoveAllTempMessage
    } Else {
        Add-Message -Message "The Report [$reportName] Already Exist" -Color 'Green' -RemoveAllTempMessage
    }

################################################
## STEP 6: Verify If Targeted Report Is Sharing With Targeted Service Account
################################################
    Add-Message -Message "Verify If Report [$reportName] Is Created By The Current Intelligence Service Account." -Color 'Yellow' -IsTempMessage $true
    [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL "v1/reports/$($targetReportDetails.id)/share/accounts" -AccessToken $intelligenceAccessToken
    [bool]$isReportNotSharedWithAccount = [string]::IsNullOrEmpty(($apiResponse.data.details.account_id | Where-Object {$_ -eq $($targetServiceAccountDetails.id)}))

    If ($targetReportDetails.created_by -ne $targetServiceAccountDetails.id) {
        Add-Message -Message "Report [$reportName] Has Not Been Created By The Current Intelligence Service Account." -Color 'Yellow' -IsTempMessage $true
        Add-Message -Message "Verify If Report [$reportName] Is Shared With The Current Intelligence Service Account." -Color 'Yellow' -IsTempMessage $true
        If ($isReportNotSharedWithAccount) {
            Add-Message -Message "The Report [$reportName] Is Not Shared With The Current Intelligence Service Account" -Color 'Yellow' -IsTempMessage $true
            # Targeted Report Is Not Shared With Targeted Service Account And Target Report Is Not Created By Targeted Service Account
            $body = ConvertTo-Json -InputObject @(
                @{
                    'user_descriptor' = @{
                        'id' = $targetServiceAccountDetails.id
                    }
                    'account_access_level' = $intelligenceAccountAccessLevel
                } ) -Depth 3
                [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL "v1/reports/$($targetReportDetails.id)/share" -AccessToken $intelligenceAccessToken -Method 'Put' -Body $body
                Add-Message -Message "The Report [$reportName] Is Successfully Shared With The Current Intelligence Service Account" -Color 'Green' -RemoveAllTempMessage
        } Else {
            Add-Message -Message "The Report [$reportName] Is Shared With The Current Intelligence Service Account" -Color 'Green' -RemoveAllTempMessage
        }
    } Else {
        Add-Message -Message "The Report [$reportName] Has Been Created By The Current Intelligence Service Account" -Color 'Green' -RemoveAllTempMessage
    }

################################################
## STEP 7: Run The Report
################################################
    Add-Message -Message "`n<-- Step 5 - Run Intelligence Report -->"
    [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL "v1/reports/$($targetReportDetails.id)/run" -AccessToken $intelligenceAccessToken -Method 'Post'
    [PSCustomObject]$targetReportRequestDetails = $apiResponse.data
    Add-Message -Message "The Report [$reportName] Is Successfully Running" -Color 'Yellow' -IsTempMessage $true

################################################
## STEP 8: Tracking Report Generation State
################################################
    [bool]$isReportRun = $true
    [int]$SupportShortDelay = 5
    While ($isReportRun) {
        [bool]$continueToSearch = $true
        [int]$offsetValue = 0

        While ($continueToSearch){
            [PSCustomObject]$body = @{
                offset = $offsetValue
                page_size = $apiPageSize
            }
            # Convert Body To JSON
            $body = $body | ConvertTo-Json
            [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL "v1/reports/$($targetReportDetails.id)/downloads/search" -AccessToken $intelligenceAccessToken -Method 'Post' -Body $body
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
                Add-Message -Message "The Report [$reportName] Is Still Running. Please Wait..." -Color 'Yellow' -IsTempMessage $true -RemoveAllTempMessage
                Start-Sleep -Seconds 10
            }
        } Else {
            If ($SupportShortDelay -ge 0) {
                $SupportShortDelay--
                Add-Message -Message "The Report [$reportName] Is Still Running. Please Wait..." -Color 'Yellow' -IsTempMessage $true -RemoveAllTempMessage
                Start-Sleep -Seconds 5
            } Else {
                Throw 'An Error Occured When Searching Report Completion Status. No Report Were Found'
            }
        }
    }
    Add-Message -Message "The Report [$reportName] Is Generated. Tracking ID [$($targetReportTrackingDetails.id)]" -Color 'Green' -RemoveAllTempMessage

################################################
## STEP 9: Download Report
################################################
    Add-Message -Message "`n<-- Step 6 - Download CSV -->"
    [string]$currentDate = Get-Date -Format 'dd_MM_yyyy_HH_mm'
    [string]$csvFileName = $reportName + '_' + $currentDate + '.csv'
    Add-Message -Message "The Report [$reportName] Will Be Downloaded Into [$reportDirectoryLocation] Repository And Will Be Named [$csvFileName]" -Color 'Yellow' -IsTempMessage $true -RemoveAllTempMessage

    [string]$reportLocationFile = $reportDirectoryLocation + '\' + $reportName + $currentDate + ".csv"
    [PSCustomObject]$apiResponse = Invoke-WorkspaceOneApi -RootURL $IntelligenceRootURL -BaseURL "v1/reports/tracking/$($targetReportTrackingDetails.id)/download" -AccessToken $intelligenceAccessToken -OutFile $reportLocationFile
    Add-Message -Message "[$csvFileName] File Has Been Successfully Downloaded Into [$reportDirectoryLocation] Repository" -Color 'Green' -RemoveAllTempMessage
}
Add-Message -Message "Thank You For Using Workspace One App Report"
Pause
#endregion
##*=============================================
##* END SCRIPT BODY
##*=============================================