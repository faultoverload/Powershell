<# 
This script will look for the most recently edited file in the C:\Reports\SCCM Deployments\
folder and try to use that to append any existing comments or review dates to the new CSV
generated in the same folder with a file name like "All Deployments12019.csv" where "1" is the 
week of the year and "2019" is the year.

This is designed to be run as a scheduled task once a week on any computer the has the SCCM console
or ConfigurationManager.psd1 module installed on it. This module can be found in the 
"C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
folder on any machine that has the console currently installed.

If you wish to use the optional Microsoft Teams output then the PSTeams module should be installed as 
well using the following command.
	Install-Module PSTeams
Additional information about this module can be found at the github link below. They have an excelent
tutorial if you need assistance setting up a custom webhook input as well.
https://github.com/EvotecIT/PSTeams



#>

param(
	[switch]$Teams,# Does the user want output to teams. Default to False
	[string]$Folder="C:\Reports\SCCM Deployments\",#Root folder for CSV files
	[string]$SearchScope="All",#Security scope to use when searching for deployments. Default is All
	[string]$IgnoreScope="",#Use this if you want to ignore deployments made by users in specific security scopes
	[string]$TeamsImage,#Insert gif or image url here
	[string]$TeamsURL#Inset teams webhook url here
)



#Global Vars
$week=Get-Date -UFormat "%V"#Used to create new file names.
$year=Get-Date -UFormat "%Y"
$newFileName= $folder+"All Deployments" +$week+$year+".csv"
$oldFileName= $folder+(Get-ChildItem $folder| Sort-Object lastwritetime -Descending | Select-Object -first 1).name
#Vars for teams reporting
$jonGIF = $TeamsImage
$webhookURL=$TeamsURL
#Import SCCM Modules and setup run location
Import-Module "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
$smsClient = New-Object -ComObject Microsoft.SMS.Client -Strict
$siteCode = $smsclient.GetAssignedSite()+":"
Set-Location $sitecode
#End SCCM site import
$oldFile=Import-Csv $oldFileName
$currentDate = Get-Date
$newDeployment=""
$newDeploymentCount=0
$oldDeployment=""
$oldDeploymentCount=0
$unknownCount = 0
$reviewCount = 0
#DeploymentIntent ::1=Required,2=Available
$deployIntent=@{
	1="Required";
	2="Available"
}
#FeatureType :: 7=Task Sequence,6=Baseline,5=Software Update,2=Program,1=Application
$featureType=@{
	1="Applications";
	2="Program";
	5="Software Update";
	6="Baseline";
	7="Task Sequence"
}
#Using this to try and fix divide by zero error
$nfi = New-Object System.Globalization.CultureInfo -ArgumentList "en-us",$false
#Data Collection
$deployments=Get-CMDeployment | Select-Object ApplicationName,CollectionName,DeploymentIntent,FeatureType,NumberSuccess,NumberTargeted,DeploymentTime,EnforcementDeadline

function generateObject($data){
	$compliance=$null#Null out old value
	$compliance= ($d.NumberSuccess / $d.NumberTargeted).ToString("P",$nfi)#needs catch for divide by zero errors
	[int]$i=$data.DeploymentIntent
	[int]$t=$data.FeatureType
	$reviewDate="Unknown"
	$reason=""

	#Import Old CSV data into new one
	foreach($o in $oldFile){
		if($o.software -like $data.ApplicationName){
			#Copies existing review dates and notes into new csv
			$reviewDate=$o.ReviewDate
			if($reviewDate -eq ""){
				$reviewDate="Unknown"
			}
			$reason=$o.Reason
		}
	}
	$object = [PSCustomObject]@{
		Software=$data.ApplicationName
		ReviewDate=$reviewDate
		Reason=$reason
		Collection=$data.CollectionName
		Purpose=$deployIntent.$i
		Type=$featureType.$t
		Compliant=$compliance
		AssetsTargeted= $data.NumberTargeted
		DeploymentStart=$data.DeploymentTime
		Deadline=$data.EnforcementDeadline
	}
	$object | Export-Csv $newFileName -Append -NoTypeInformation -Force
}


#Data parsing to generate new spreadsheet
foreach($d in $deployments){

	if($d.featuretype -eq "2"){#Program/Package Logic
		$packName= $d.ApplicationName.split("(",2)[0].trim()#Gets package name
		$package = get-cmpackage -name $packName
		if($package.SecuredScopeNames -notcontains $ignoreScope -and $package.SecuredScopeNames -contains $searchScope){
			$col=GET-CMCollection -name $d.CollectionName
			if($col.LimitToCollectionName -notlike "All Servers"){
				generateObject -data $d
			}	
		}
	}elseif($d.featuretype -eq "1"){#Application Logic
		$appName=$d.ApplicationName
		$app=get-cmapplication -name $appName
		if($app.SecuredScopeNames -contains $searchScope -and $app.SecuredScopeNames -notcontains $ignoreScope){
			generateObject -data $d
		}
	}elseif($d.featuretype -eq "5"){#Software Updates
		#Ignoring Software Updates
	}elseif($d.featuretype -eq "6"){#Baseline
		$baseName=$d.ApplicationName
		$base=Get-CMBaseline -Name $baseName
		if($base.IsAssigned){
			generateObject -data $d
		}
	}elseif($d.featuretype -eq "7"){#Task Sequence
		$scopeName = $d.ApplicationName
		$task=get-cmtasksequence -name $scopeName
		if($task.SecuredScopeNames -contains $searchScope -and $task.SecuredScopeNames -notcontains $ignoreScope){
			generateObject -data $d
		}

	}
}

#Start Teams Message output.
if($Teams){
#Notify users about upcoming review items and new deployments.(Now Jenkins/Teams)
	$newFile=Import-Csv $newFileName #This must be here so its imported after being fully created
<#Generate strings for data I care about
List of new deployment names added
List of deployments deleted
Count of how many have unknown or blank review dates
Count of how many have past due review dates
#>
	foreach($item in $newFile){
		if($oldFile.software -like $item.software){}else{
			$newDeployment+=("- "+$item.software+"`n")#Generates list of new deployments
			$newDeploymentCount++
		}
		if($item.ReviewDate -eq "Unknown" -or $item.ReviewDate -eq ""){
			$unknownCount++
		}else{
			$formatedDate = [datetime]$item.ReviewDate
			if($formatedDate -lt $currentDate){
				$reviewCount++
			}
		}
	}
	foreach($item in $oldFile){#Old items
		if($newFile.software -like $item.software){}else{
			$oldDeployment+=("- "+$item.software+"`n")#Generates list of old deployments
			$oldDeploymentCount++
		}
	}
	$oldDeploymentFact = New-TeamsFact -Name "Deleted Deployments" -Value $oldDeployment
	$newDeploymentFact = New-TeamsFact -Name "Added Deployments" -Value $newDeployment
	$summaryFact = New-TeamsFact -Name "Summary" -Value "$unknownCount deployment(s) currently `
	have an unknown review date. `n`
	$reviewCount deployment(s) are past their review date.`n`
	$newDeploymentCount deployment(s) are newly detected this past week.`n`
	$oldDeploymentCount deployment(s) were deleted this past week"
	#Summary of review and unknown review deployments
	$Section1 = New-TeamsSection `
    	-ActivityTitle "Weekly Deployment Status" `
    	-ActivitySubtitle $CurrentDate `
    	-ActivityImageLink $jonGIF `
    	-ActivityText "" `
    	-ActivityDetails $summaryFact, $newDeploymentFact, $oldDeploymentFact
	
	Send-TeamsMessage `
    	-URI $webhookURL `
    	-MessageTitle 'Weekly deployment update' `
    	-MessageText "This text won't show up" `
    	-Color DodgerBlue `
    	-Sections $Section1

}
