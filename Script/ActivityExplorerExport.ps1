<#PSScriptInfo

.VERSION 1.0.0

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
	Script      : ActivityExplorerExport.ps1
	Author      : S. Zamorano
	Version     : 1.0.0
	Description : Export Activity Explorer activities to CSV or Json format.
	17-04-2024		S. Zamorano		- Public release
#>

[CmdletBinding(DefaultParameterSetName = "None")]
param(
	#Export-ActivityExplorerData cmdlet requires a PageSize that can be between 1 to 5000, by default is set to 100, you can change the number below or use the parameter -ChangePageSize to modify during the execution
	[int]$InitialPageSize = 5000,
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
        [switch]$NotApplyFilters,
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
	Connect-IPPSSession -UseRPSSession:$false
}

function ReadListOfActivities
{
	$ActivityList = @()
	$configfile = "$PSScriptRoot\ActivityExplorerActivitiesSelector.json"
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

function WriteToJson($results, $ExportFolder, $Activity, $date, $cmdlet)
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
	$FileName = "ActivityExplorer export - "+"$Activities"+" - "+"$date"+".Json"
	$pathJson = $PSScriptRoot+"\"+$ExportFolder+"\"+$FileName
	$path = $pathJson
	$json | Add-Content -Path $path
	Write-Host "`nCmdlet used: " -NoNewLine
	Write-Host "$cmdlet" -ForeGroundColor Green
	Write-Host "`nData exported to... :" -NoNewLine
	Write-Host $pathJson -ForeGroundColor Cyan
	Write-Host "`n----------------------------------------------------------------------------------------`n`n" -ForeGroundColor DarkBlue
}

function WriteToCsv($results, $ExportFolder, $Activity, $date, $cmdlet)
{
	$Activities = $Activity
	$parse_array = @()
	$nextpages_array = @()
	$Counter = 0
	foreach($item in $results)
	{
		$FileName = "ActivityExplorer export - "+"$Activities"+" - "+"$date"+" - "+$Counter+".Csv"
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

function GetActivityExplorerData($StartingTime, $EndingTime, $ExportFormat, $ExportFolder, $Activities)
{
	Write-Host "`nExecuting Export-ActivityExplorerData cmdlet..." -ForeGroundColor Blue
	
	$date = (Get-Date).ToString("yyyy-MM-dd HHmm")
	if($Activities -eq "All")
	{
		$results = New-Object PSObject
		$TotalResults = @()
		$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize -OutputFormat $ExportFormat
		$TotalResults += $results
		Write-Host "Collecting data..." -ForegroundColor DarkBlue
		#Run the below steps in loop until all results are fetched
		while ($results.LastPage -ne $true)
		{
			Write-Host "Please wait until the script finish, the data is added to the same file."
			$pageCookie = $results.WaterMark
			$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize -OutputFormat $ExportFormat -PageCookie $pageCookie
			$TotalResults += $results
		}

		if($ExportFormat -eq "Csv")
		{
			$cmdletused = "Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize-OutputFormat $ExportFormat"
			$CSVresults = $TotalResults.ResultData
			WriteToCsv -results $CSVresults -ExportFolder $ExportFolder -Activity $Activity -date $date -cmdlet $cmdletused
			Write-Host $CSVresults -ForeGroundColor Red
		}else
		{
			$cmdletused = "Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize -OutputFormat $ExportFormat"
			WriteToJson -results $TotalResults -ExportFolder $ExportFolder -Activity $Activities -date $date -cmdlet $cmdletused
		}
	}else
	{
		foreach($Activity in $Activities)
		{
			Write-Host "`n$Activity`n" -ForeGroundColor Green
			$results = New-Object PSObject
			$TotalResults = @()
			$results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Activity', $Activity) -PageSize $InitialPageSize -OutputFormat $ExportFormat
			$TotalResults += $results
			Write-Host "Collecting data..." -ForegroundColor DarkBlue
			#Run the below steps in loop until all results are fetched
			while ($results.LastPage -ne $true)
			{
			  Write-Host "Please wait until the script finish, the data is added to the same file."
			  $pageCookie = $results.WaterMark
			  $results = Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -Filter1 @('Activity', $Activity) -PageSize $InitialPageSize -OutputFormat $ExportFormat -PageCookie $pageCookie
			  $TotalResults += $results
			}
			
			if($ExportFormat -eq "Csv")
			{
				$cmdletused = "Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize -Filter1 @('Activity', $Activity) -OutputFormat $ExportFormat"
				$CSVresults = $TotalResults.ResultData
				WriteToCsv -results $CSVresults -ExportFolder $ExportFolder -Activity $Activity -date $date -cmdlet $cmdletused
			}else
			{
				$cmdletused = "Export-ActivityExplorerData -StartTime $StartingTime -EndTime $EndingTime -PageSize $InitialPageSize -Filter1 @('Activity', $Activity) -OutputFormat $ExportFormat"
				WriteToJson -results $TotalResults -ExportFolder $ExportFolder -Activity $Activity -date $date -cmdlet $cmdletused
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
	
	
	##List of activities
	if(-Not $NotApplyFilters)
	{
		$ListOfActivities = ReadListOfActivities
	}else
	{
		$ListOfActivities = "All"
	}
	
	##Export format
	$ExportFormat = "Json"
	if($ExportToCsv)
	{
		$ExportFormat = "Csv"
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
	
	Write-Host "`n`nYou will be prompted for your credentials, remember that you need Compliance Administrator role"
	Write-Host "Press any key to continue..."
    $key = ([System.Console]::ReadKey($true))
	#connect2service
	
	#Call function to export data from Activity Explorer
	GetActivityExplorerData -StartingTime $strt -EndingTime $end -ExportFormat $ExportFormat -ExportFolder $ExportFolderName -Activities $ListOfActivities
	
	#SelectContinuity
}

CheckPrerequisites
MainFunction