<#PSScriptInfo

.VERSION 2.0.1

.GUID 883af802-165c-4708-b4c1-352686c02f01

.AUTHOR 
https://www.linkedin.com/in/profesorkaz/; Sebastian Zamorano

.COMPANYNAME 
Microsoft Purview Advanced Rich Reports

.TAGS 
#Microsoft365 #M365 #MPARR #MicrosoftPurview #ActivityExplorer

.PROJECTURI 
https://aka.ms/MPARR-YouTube 

.RELEASENOTES
The MIT License (MIT)
Copyright (c) 2015 Microsoft Corporation
Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:
The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.
THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

#>

<# 

.DESCRIPTION 
This script permit to export Activity Explorer activities based in a configuration file

#>

<#
HISTORY
	Script      : ActivityExplorerExport2.ps1
	Author      : S. Zamorano
	Version     : 2.0.1
	Description : Export Activity Explorer activities to CSV or Json format.
	17-04-2024		S. Zamorano		- Public release
	12-08-2024		S. Zamorano		- Version 2 Public release
#>

[CmdletBinding(DefaultParameterSetName = "None")]
param(
	#Export-ActivityExplorerData cmdlet requires a PageSize that can be between 1 to 5000, by default is set to 100, you can change the number below or use the parameter -ChangePageSize to modify during the execution
	[int]$InitialPageSize = 5000,
	[string]$TableName = "ActivityExplorer",
	[Parameter()] 
        [switch]$Help,
	[Parameter(ParameterSetName="CustomParams")] 
    [Parameter(ParameterSetName="CustomParams1")] 
        [switch]$UseCustomParameters,
    [Parameter(ParameterSetName="CustomParams", Mandatory=$true)] 
        [datetime]$AEStartTime,
    [Parameter(ParameterSetName="CustomParams", Mandatory=$true)] 
        [datetime]$AEEndTime,
	[Parameter()] 
        [switch]$ExportToCsv,
	[Parameter()] 
        [switch]$ExportToLogsAnalytics,
	[Parameter()] 
        [switch]$NotApplyFilters,
	[Parameter()] 
        [switch]$WorkloadsFilter,
	[Parameter()] 
        [switch]$FilterOnlyByWorkloads,
	[Parameter()] 
        [int]$PastDays
)

function CheckPowerShellVersion
{
    # Check PowerShell version
    Write-Host "`nChecking PowerShell version... " -NoNewline
    if ($Host.Version.Major -gt 5)
    {
        Write-Host "`t`t`t`tPassed!" -ForegroundColor Green
    }
    else
    {
        Write-Host "Failed" -ForegroundColor Red
        Write-Host "`tCurrent version is $($Host.Version). PowerShell version 7 or newer is required."
        exit(1)
    }
}

function CheckIfElevated
{
    $IsElevated = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (!$IsElevated)
    {
        Write-Host "`nPlease start PowerShell as Administrator.`n" -ForegroundColor Yellow
        exit(1)
    }
}

function CheckRequiredModules 
{
    # Check PowerShell modules
    Write-Host "Checking PowerShell modules..."
    $requiredModules = @(
        @{Name="ExchangeOnlineManagement"; MinVersion="0.0"}
        )

    $modulesToInstall = @()
    foreach ($module in $requiredModules)
    {
        Write-Host "`t$($module.Name) - " -NoNewline
        $installedVersions = Get-Module -ListAvailable $module.Name
        if ($installedVersions)
        {
            if ($installedVersions[0].Version -lt [version]$module.MinVersion)
            {
                Write-Host "`t`t`tNew version required" -ForegroundColor Red
                $modulesToInstall += $module.Name
            }
            else 
            {
                Write-Host "`t`t`tInstalled" -ForegroundColor Green
            }
        }
        else
        {
            Write-Host "`t`t`tNot installed" -ForegroundColor Red
            $modulesToInstall += $module.Name
        }
    }

    if ($modulesToInstall.Count -gt 0)
    {
        CheckIfElevated
		$choices  = '&Yes', '&No'

        $decision = $Host.UI.PromptForChoice("", "Misisng required modules. Proceed with installation?", $choices, 0)
        if ($decision -eq 0) 
        {
            Write-Host "Installing modules..."
            foreach ($module in $modulesToInstall)
            {
                Write-Host "`t$module"
				Install-Module $module -ErrorAction Stop
                
            }
            Write-Host "`nModules installed. Please start the script again."
            exit(0)
        } 
        else 
        {
            Write-Host "`nExiting setup. Please install required modules and re-run the setup."
            exit(1)
        }
    }
}

function CheckPrerequisites
{
    CheckPowerShellVersion
	CheckRequiredModules
}

function connect2service
{
	Write-Host "`nAuthentication is required, please check your browser" -ForegroundColor DarkYellow
	Connect-IPPSSession -UseRPSSession:$false -ShowBanner:$false
}

function DecryptSharedKey 
{
    param(
        [string] $encryptedKey
    )

    try {
        $secureKey = $encryptedKey | ConvertTo-SecureString -ErrorAction Stop  
    }
    catch {
        Write-Error "Workspace key: $($_.Exception.Message)"
        exit(1)
    }
    $BSTR =  [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureKey)
    $plainKey = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
    $plainKey
}

function ReadListOfActivities
{
	$ActivityList = @()
	$configfile = "$PSScriptRoot\ConfigFiles\ActivityExplorerActivitiesSelector.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$activities = ConvertFrom-Json -InputObject $json
	foreach ($activity in $activities.psobject.Properties)
	{
		if ($activities."$($activity.Name)" -eq "True")
		{
			$ActivityList += $activity.Name
		}
	}
	return $ActivityList
}

function ReadListOfWorkloads
{
	$WorkloadList = @()
	$configfile = "$PSScriptRoot\ConfigFiles\ActivityExplorerWorkloadsSelector.json"
	$json = Get-Content -Raw -Path $configfile
	[PSCustomObject]$workloads = ConvertFrom-Json -InputObject $json
	foreach ($workload in $workloads.psobject.Properties)
	{
		if ($workloads."$($workload.Name)" -eq "True")
		{
			$WorkloadList += $workload.Name
		}
	}
	return $WorkloadList
}

function SelectContinuity
{
	$choices  = '&Yes','&No'
	$decision = $Host.UI.PromptForChoice("", "`nDo you want to export more data? ", $choices, 1)
	
	if ($decision -eq 0)
    {
		MainFunction
	}
	if ($decision -eq 1)
	{
		exit
	}
	
}

function Build-Signature ($customerId, $sharedKey, $date, $contentLength, $method, $contentType, $resource) 
{
    # ---------------------------------------------------------------   
    #    Name           : Build-Signature
    #    Value          : Creates the authorization signature used in the REST API call to Log Analytics
    # ---------------------------------------------------------------

	#Original function to Logs Analytics
    $xHeaders = "x-ms-date:" + $date
    $stringToHash = $method + "`n" + $contentLength + "`n" + $contentType + "`n" + $xHeaders + "`n" + $resource

    $bytesToHash = [Text.Encoding]::UTF8.GetBytes($stringToHash)
    $keyBytes = [Convert]::FromBase64String($sharedKey)

    $sha256 = New-Object System.Security.Cryptography.HMACSHA256
    $sha256.Key = $keyBytes
    $calculatedHash = $sha256.ComputeHash($bytesToHash)
    $encodedHash = [Convert]::ToBase64String($calculatedHash)
    $authorization = 'SharedKey {0}:{1}' -f $customerId,$encodedHash
    return $authorization
}

function WriteToLogsAnalytics($body, $LogAnalyticsTableName) 
{
    # ---------------------------------------------------------------   
    #    Name           : Post-LogAnalyticsData
    #    Value          : Writes the data to Log Analytics using a REST API
    #    Input          : 1) PSObject with the data
    #                     2) Table name in Log Analytics
    #    Return         : None
    # ---------------------------------------------------------------
    
	#Read configuration file
	$CONFIGFILE = "$PSScriptRoot\ConfigFiles\ActivityExplorerConfiguration.json"
	$json = Get-Content -Raw -Path $CONFIGFILE
	[PSCustomObject]$config = ConvertFrom-Json -InputObject $json
	
	$EncryptedKeys = $config.EncryptedKeys
	$WLA_CustomerID = $config.Workspace_ID
	$WLA_SharedKey = $config.WorkspacePrimaryKey
	if ($EncryptedKeys -eq "True")
	{
		$WLA_SharedKey = DecryptSharedKey $WLA_SharedKey
	}

	# Your Log Analytics workspace ID
	$LogAnalyticsWorkspaceId = $WLA_CustomerID

	# Use either the primary or the secondary Connected Sources client authentication key   
	$LogAnalyticsPrimaryKey = $WLA_SharedKey
	
	#Step 0: sanity checks
    if($body -isnot [array]) {return}
    if($body.Count -eq 0) {return}
	
	#Step 1: convert the body.ResultData to JSON
	$json_array = @()
	$parse_array = @()
	$parse_array = $body.ResultData | ConvertFrom-Json
	foreach($item in $parse_array) 
	{
		$json_array += $item
	}
	$json = $json_array | ConvertTo-Json -Depth 6
	
	#Step 2: convert the PSObject to JSON
	$bodyJson = $json
	#Step 2.5: sanity checks
	if($bodyJson.Count -eq 0) {return}

    #Step 3: get the UTF8 bytestream for the JSON
    $bodyJsonUTF8 = ([System.Text.Encoding]::UTF8.GetBytes($bodyJson))
	
	#Step 4: build the signature        
    $method = "POST"
    $contentType = "application/json"
    $resource = "/api/logs"
    $rfc1123date = [DateTime]::UtcNow.ToString("r")
    $contentLength = $bodyJsonUTF8.Length    
    $signature = Build-Signature -customerId $LogAnalyticsWorkspaceId -sharedKey $LogAnalyticsPrimaryKey -date $rfc1123date -contentLength $contentLength -method $method -contentType $contentType -resource $resource
    
    #Step 5: create the header
    $headers = @{
        "Authorization" = $signature;
        "Log-Type" = $LogAnalyticsTableName;
        "x-ms-date" = $rfc1123date;
    };

    #Step 6: REST API call
    $uri = 'https://' + $LogAnalyticsWorkspaceId + ".ods.opinsights.azure.com" + $resource + "?api-version=2016-04-01"
    $response = Invoke-WebRequest -Uri $uri -Method Post -Headers $headers -ContentType $contentType -Body $bodyJsonUTF8 -UseBasicParsing

    if ($Response.StatusCode -eq 200) {   
        $rows = $json_array.Count
        Write-Information -MessageData "$rows rows written to Log Analytics workspace $uri" -InformationAction Continue
    }
}

function WriteToJson($results, $ExportFolder, $Activity, $Workload, $date, $cmdlet)
{
	$Activities = $Activity
	$json_array = @() 
	$parse_array = @()
	$parse_array = $results.ResultData | ConvertFrom-Json
	foreach($item in $parse_array) 
	{
		$json_array += $item
	}
	$json = $json_array | ConvertTo-Json -Depth 6
	if($Workload)
	{
		$FileName = "ActivityExplorer export - "+"$Activities in $Workload"+" - "+"$date"+".Json"
	}else
	{
		$FileName = "ActivityExplorer export - "+"$Activities"+" - "+"$date"+".Json"
	}
	$pathJson = $PSScriptRoot+"\"+$ExportFolder+"\"+$FileName
	$path = $pathJson
	$json | Add-Content -Path $path
	Write-Host "`nCmdlet used: " -NoNewLine
	Write-Host "$cmdlet" -ForeGroundColor Green
	Write-Host "`nData exported to... :" -NoNewLine
	Write-Host $pathJson -ForeGroundColor Cyan
	Write-Host "`n----------------------------------------------------------------------------------------`n`n" -ForeGroundColor DarkBlue
}

function WriteToCsv($results, $ExportFolder, $Activity, $Workload, $date, $cmdlet)
{
	$Activities = $Activity
	$parse_array = @()
	$nextpages_array = @()
	$Counter = 0
	foreach($item in $results)
	{
		if($Workload)
		{
			$FileName = "ActivityExplorer export - "+"$Activities in $Workload"+" - "+"$date"+" - "+$Counter+".Csv"
		}else
		{
			$FileName = "ActivityExplorer export - "+"$Activities"+" - "+"$date"+" - "+$Counter+".Csv"
		}
		$pathCsv = $PSScriptRoot+"\"+$ExportFolder+"\"+$FileName
		$path = $pathCsv
		$parse_array = $item| ConvertFrom-Csv
		if($Counter -eq 0)
		{
			$values = $parse_array[0].psobject.properties.name
			$parse_array | Export-Csv -Path $path -NTI -Force -Append | Out-Null
		}else
		{
			
			$nextpages_array | Export-Csv -Path $path -NTI -Force -Append | Out-Null
		}
		$Counter++
	}
	Write-Host "Total results $($results.count)"
	
	Write-Host "`nCmdlet used: " -NoNewLine
	Write-Host "$cmdlet" -ForeGroundColor Green
	Write-Host "`nData exported to..." -NoNewline
	Write-Host "`n$pathCsv" -ForeGroundColor Cyan
}

function ActivityExplorerHelp
{
	cls
	Write-Host "`n"
	Write-Host "################################################################################" -ForegroundColor Green
	Write-Host "`n How to use this script `n" -ForegroundColor Green
	Write-Host "################################################################################" -ForegroundColor Green
	Write-Host "`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "This menu"
	Write-Host ".\ActivityExplorerExport.ps1 -Help" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "Using only the script by default, you'll be able to capture and calculate data from the past 24 hours, based on the exact time of execution."
	Write-Host ".\ActivityExplorerExport.ps1" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "With this next execution, you'll be able to capture and calculate data from the past X days, based on the exact time of execution"
	Write-Host ".\ActivityExplorerExport.ps1 -PastDays 31" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "This script uses a configuration file located in the same folder as the script. The file, called ActivityExplorerActivitiesSelector.json, contains"
	Write-Host "a list of common activities and is used by default with the values set to 'True.' It's possible that some activities are not listed; you can manually add those."
	Write-Host "If the file is not present, the script will take all possible activities. Alternatively, you can use the '-NotApplyFilters' attribute to retrieve all activities."
	Write-Host ".\ActivityExplorerExport.ps1 -NotApplyFilters" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "You can request a specific time frame for data retrieval. For example, to get all the data for June 18, 2024, you can set the desired interval."
	Write-Host "This can be a short interval or an extended period spanning several days."
	Write-Host ".\ActivityExplorerExport.ps1 -UseCustomParameters -AEStartTime '"'06/18/2024 00:00:00'"' -AEEndTime '"'06/18/2024 23:59:59'"'" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "If you are not comfortable working with JSON format, you can use the attribute '-ExportToCsv' to export the data in CSV format."
	Write-Host ".\ActivityExplorerExport.ps1 -ExportToCsv" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "If you want to store your data in Logs Analytics you can use the attribute '-ExportToLogsAnalytics'."
	Write-Host ".\ActivityExplorerExport.ps1 -ExportToLogsAnalytics" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "The PowerShell cmdlet used in this script supports a maximum Page Size of 5,000, which is the default value. However, if you want to change this Page Size,"
	Write-Host "you can use the attribute -InitialPageSize and set a smaller value."
	Write-Host ".\ActivityExplorerExport.ps1 -InitialPageSize 300" -ForeGroundColor DarkYellow
	Write-Host "`n`nDescription: " -ForegroundColor Blue -NoNewLine
	Write-Host "You can combine different attributes available in the script to customize its functionality. For example:"
	Write-Host ".\ActivityExplorerExport.ps1 -NotApplyFilters -PastDays 15" -ForeGroundColor DarkYellow
	Write-Host "`n"
	Write-Host "### You can now proceed using any of the options listed in the Help menu. ###" -ForegroundColor Green
	Write-Host "`n"
	return
}

function GetActivityExplorerData($StartingTime, $EndingTime, $ExportFormat, $ExportFolder, $Activities, $Workloads, $ExportOption)
{
	Write-Host "`nExecuting Export-ActivityExplorerData cmdlet for your selection..." -ForeGroundColor Blue
	
	$date = (Get-Date).ToString("yyyy-MM-dd HHmm")
	$ExportExtension = $ExportFormat
	if($ExportFormat -eq "LA")
	{
		$ExportExtension="Json"
	}
	if($ExportOption -eq "All")
	{
		$results = New-Object PSObject
		$TotalResults = @()
		$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize -OutputFormat $ExportExtension
		$TotalResults += $results
		if($results.TotalResultCount -eq "0")
			{
				Write-Host "The previous combination does not return any values, try extending the period used"
				Write-Host "Exiting...`n"
			}else
			{
				Write-Host "`nCollecting data..." -ForegroundColor DarkBlue -NoNewLine
				Write-Host $results.TotalResultCount -ForegroundColor Blue -NoNewLine
				Write-Host " records returned"
				#Run the below steps in loop until all results are fetched
				while ($results.LastPage -ne $true)
				{
					Write-Host "Please wait until the script finish, the data is added to the same file."
					$pageCookie = $results.WaterMark
					$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize -OutputFormat $ExportExtension -PageCookie $pageCookie
					$TotalResults += $results
				}

				if($ExportFormat -eq "Csv")
				{
					$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize-OutputFormat $ExportExtension"
					$CSVresults = $TotalResults.ResultData
					WriteToCsv -results $CSVresults -ExportFolder $ExportFolder -Activity $Activity -date $date -cmdlet $cmdletused
					Write-Host $CSVresults -ForeGroundColor Red
				}elseif($ExportFormat -eq "LA")
				{
					WriteToLogsAnalytics -LogAnalyticsTableName $TableName -body $TotalResults
				}else
				{
					$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize -OutputFormat $ExportExtension"
					WriteToJson -results $TotalResults -ExportFolder $ExportFolder -Activity $Activities -date $date -cmdlet $cmdletused
				}
			}
	}
	elseif($ExportOption -eq "FilteredByActivities")
	{
		foreach($Activity in $Activities)
		{
			$results = New-Object PSObject
			$TotalResults = @()
			Write-Host "`n$Activity`n" -ForeGroundColor Green
			$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Activity', $Activity) -PageSize $InitialPageSize -OutputFormat $ExportExtension

			$TotalResults += $results
			if($results.TotalResultCount -eq "0")
			{
				Write-Host "The previous Activity does not return any values"
				Write-Host "Continuing..."
			}else
			{
				Write-Host "Collecting data..." -ForegroundColor DarkBlue -NoNewLine
				Write-Host $results.TotalResultCount -ForegroundColor Blue -NoNewLine
				Write-Host " records returned"
				#Run the below steps in loop until all results are fetched
				while ($results.LastPage -ne $true)
				{
					Write-Host "Please wait until the script finish, the data is added to the same file."
					$pageCookie = $results.WaterMark
					$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Activity', $Activity) -PageSize $InitialPageSize -OutputFormat $ExportExtension -PageCookie $pageCookie
					$TotalResults += $results
				}
			
				if($ExportFormat -eq "Csv")
				{
					$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize -Filter1 @('Activity', '$Activity') -OutputFormat $ExportExtension"
					$CSVresults = $TotalResults.ResultData
					WriteToCsv -results $CSVresults -ExportFolder $ExportFolder -Activity $Activity -date $date -cmdlet $cmdletused
				}elseif($ExportFormat -eq "LA")
				{
					WriteToLogsAnalytics -LogAnalyticsTableName $TableName -body $TotalResults
				}else
				{
					$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize -Filter1 @('Activity', '$Activity') -OutputFormat $ExportExtension"
					WriteToJson -results $TotalResults -ExportFolder $ExportFolder -Activity $Activity -date $date -cmdlet $cmdletused
				}
			}
		}
	}
	elseif($ExportOption -eq "FilteredByWorkload")
	{
		foreach($workload in $Workloads)
		{
			$results = New-Object PSObject
			$TotalResults = @()
			Write-Host "`n$workload`n" -ForeGroundColor Green
			$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Workload', $workload) -PageSize $InitialPageSize -OutputFormat $ExportExtension

			$TotalResults += $results
			if($results.TotalResultCount -eq "0")
			{
				Write-Host "The previous Workload does not return any values"
				Write-Host "Continuing..."
			}else
			{
				Write-Host "Collecting data..." -ForegroundColor DarkBlue -NoNewLine
				Write-Host $results.TotalResultCount -ForegroundColor Blue -NoNewLine
				Write-Host " records returned"
				#Run the below steps in loop until all results are fetched
				while ($results.LastPage -ne $true)
				{
					Write-Host "Please wait until the script finish, the data is added to the same file."
					$pageCookie = $results.WaterMark
					$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Workload', $workload) -PageSize $InitialPageSize -OutputFormat $ExportExtension -PageCookie $pageCookie
					$TotalResults += $results
				}
			
				if($ExportFormat -eq "Csv")
				{
					$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize -Filter1 @('Workload', '$workload') -OutputFormat $ExportExtension"
					$CSVresults = $TotalResults.ResultData
					WriteToCsv -results $CSVresults -ExportFolder $ExportFolder -Activity $workload -date $date -cmdlet $cmdletused
				}elseif($ExportFormat -eq "LA")
				{
					WriteToLogsAnalytics -LogAnalyticsTableName $TableName -body $TotalResults
				}else
				{
					$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize -Filter1 @('Workload', '$workload') -OutputFormat $ExportExtension"
					WriteToJson -results $TotalResults -ExportFolder $ExportFolder -Activity $workload -date $date -cmdlet $cmdletused
				}
			}
		}
	}
	else
	{
		foreach($workload in $Workloads)
		{
			foreach($Activity in $Activities)
			{
				$results = New-Object PSObject
				$TotalResults = @()
				Write-Host "`n$Activity" -ForeGroundColor Green -NoNewline
				Write-Host " in " -NoNewline
				Write-Host "$workload" -ForeGroundColor Green
				$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Activity', $Activity) -Filter2 @('Workload', $workload) -PageSize $InitialPageSize -OutputFormat $ExportExtension
				
				$TotalResults += $results
				if($results.TotalResultCount -eq "0")
				{
					Write-Host "The previous combination does not return any values"
					Write-Host "Continuing..."
				}else
				{
					Write-Host "Collecting data..." -ForegroundColor DarkBlue -NoNewLine
					Write-Host $results.TotalResultCount -ForegroundColor Blue -NoNewLine
					Write-Host " returned"
					#Run the below steps in loop until all results are fetched
					while ($results.LastPage -ne $true)
					{
						Write-Host "Please wait until the script finish, the data is added to the same file."
						$pageCookie = $results.WaterMark
						$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Activity', $Activity) -PageSize $InitialPageSize -OutputFormat $ExportExtension -PageCookie $pageCookie
						$TotalResults += $results
					}
					
					if($ExportFormat -eq "Csv")
					{
						$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize -Filter1 @('Activity', '$Activity') -OutputFormat $ExportExtension"
						$CSVresults = $TotalResults.ResultData
						WriteToCsv -results $CSVresults -ExportFolder $ExportFolder -Activity $Activity -Workload $workload -date $date -cmdlet $cmdletused
					}elseif($ExportFormat -eq "LA")
					{
						WriteToLogsAnalytics -LogAnalyticsTableName $TableName -body $TotalResults
					}else
					{
						$cmdletused = "Export-ActivityExplorerData -StartTime '$StartingTime' -EndTime '$EndingTime' -PageSize $InitialPageSize -Filter1 @('Activity', '$Activity') -OutputFormat $ExportExtension"
						WriteToJson -results $TotalResults -ExportFolder $ExportFolder -Activity $Activity -Workload $workload -date $date -cmdlet $cmdletused
					}
				}
			}
		}
	}
}

function MainFunction
{
	#Welcome header
	cls
	Clear-Host
	
	Write-Host "`n`n----------------------------------------------------------------------------------------"
	Write-Host "`nWelcome to Activity Explorer Export script!" -ForegroundColor Green
	Write-Host "This script will permit to collect data from Activity Explorer."
	Write-Host "`n----------------------------------------------------------------------------------------" 
	
	
	#Initiate variables
	
	## Time variable
	if(-Not $PastDays)
	{
		$PastData = 1
	}else
	{
		$PastData = $PastDays
	}
	
	if ($UseCustomParameters)
	{
		$strt = $AEStartTime.ToString("MM/dd/yyyy HH:mm:ss")
		$end  = $AEEndTime.ToString("MM/dd/yyyy HH:mm:ss")
	}else
	{
		$end = [DateTime]::UtcNow.ToString("MM/dd/yyyy HH:mm:ss") 
		$strt = [DateTime]::UtcNow.AddDays(-($PastData)).ToString("MM/dd/yyyy HH:mm:ss") 
	}
	
	$ExportOption = "All"
	
	##List of activities
	$ListOfActivities = "All"
	if(-Not $NotApplyFilters)
	{
		$ListOfActivities = ReadListOfActivities
		$ExportOption = "FilteredByActivities"
	}
	if($FilterOnlyByWorkloads)
	{
		$ListOfActivities = "All"
	}
	
	##List each activitiy by Workloads
	$ListOfWorkloads = "All"
	if($WorkloadsFilter)
	{
		if(-Not $NotApplyFilters)
		{
			$ListOfWorkloads = ReadListOfWorkloads
			$ExportOption = "FilteredActivitiesByWorkload"
		}else
		{
			$ListOfWorkloads = "All"
		}
	}
	
	##List all activities by Workload
	if($FilterOnlyByWorkloads)
	{
		$ExportOption = "FilteredByWorkload"
		$ListOfWorkloads = ReadListOfWorkloads
	}
	
	##Export format
	$ExportFormat = "Json"
	if($ExportToCsv)
	{
		$ExportFormat = "Csv"
	}
	if($ExportToLogsAnalytics)
	{
		$ExportFormat = "LA"
		$LogsAnalyticsConfigurationFile = "$PSScriptRoot\ConfigFiles\ActivityExplorerConfiguration.json"
		if(-not (Test-Path -Path $LogsAnalyticsConfigurationFile))
		{
			Write-Host "`nConfiguration file is not present" -ForegroundColor DarkYellow
			Write-Host "Please download the configuration file from http://activityexplorer.kaznets.com and save inside of the ConfigFiles folder.`n"
			Write-Host "Press any key to continue..."
			$key = ([System.Console]::ReadKey($true))
			exit
		}	
	}
	
	##Export folder Name
	$ExportFolderName = "ExportedData"
	$ExportPath = "$PSScriptRoot\$ExportFolderName"
	if(-Not (Test-Path $ExportPath))
	{
		New-Item -ItemType Directory -Force -Path "$PSScriptRoot\$ExportFolderName" | Out-Null
		$StatusFolder = "Created"
	}else
	{
		$StatusFolder = "Available"
	}
	
	##Show variables set
	Write-Host "Start time set to`t:" -NoNewline
	Write-Host "`t$strt" -ForegroundColor Green
	Write-Host "End time set to`t`t:" -NoNewline
	Write-Host "`t$end" -ForegroundColor Green
	Write-Host "Export format set to`t:" -NoNewline
	Write-Host "`t$ExportFormat" -ForegroundColor Green
	Write-Host "Export folder set to`t:" -NoNewline
	Write-Host "`t$ExportFolderName ($StatusFolder)" -ForegroundColor Green
	Write-Host "Page size set to`t:" -NoNewline
	Write-Host "`t$InitialPageSize" -ForegroundColor Green
	Write-Host "Activities selected`t:"
	foreach($activity in $ListOfActivities)
	{
		Write-Host "`t`t`t`t$activity" -ForegroundColor Green
	}
	Write-Host "Workloads selected`t:"
	foreach($workload in $ListOfWorkloads)
	{
		Write-Host "`t`t`t`t$workload" -ForegroundColor Green
	}
	Write-Host "`n`nYou will be prompted for your credentials, remember that you need Compliance Administrator role"
	Write-Host "Press any key to continue..."
    $key = ([System.Console]::ReadKey($true))
	#connect2service
	
	Write-Host "Calling script..."
	
	#Call function to export data from Activity Explorer
	GetActivityExplorerData -StartingTime $strt -EndingTime $end -ExportFormat $ExportFormat -ExportFolder $ExportFolderName -Activities $ListOfActivities -Workloads $ListOfWorkloads -ExportOption $ExportOption
	
	#SelectContinuity
}

if($Help)
{
	ActivityExplorerHelp
	exit
}

CheckPrerequisites
MainFunction