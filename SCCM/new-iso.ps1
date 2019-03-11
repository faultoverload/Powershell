<#
Should generate a new ISO file.

Uses either New-CMBootableMedia or New-CMStandaloneMedia for core functionality

Can take command line input or use defaults for destination and media type
Optional message in teams or email when image is finished.

#>


param(
    [Parameter(Mandatory=$true)][string]$Folder="",#Root folder for ISO to be saved to
    [Parameter(Mandatory=$true)][ValidateSet('WinPE','TaskSequence')][string]$mediaType="",#Select WinPE or Full image type
    [Parameter(Mandatory=$true)][string]$DP, #Point at a single DP or Group
    [string]$BackupSourceFolder,#Location of backup XML files
    [string]$ISOName, #Set ISO Name
    [string]$TSName, #Sets Task sequence name. Supports wildcards
    [string]$WinPEName,#Sets WinPE name. Supports wildcards
    [switch]$Teams,# Does the user want output to teams. Default to False
    [string]$TeamsImage="",#Insert gif or image url here
    [string]$TeamsURL,#Inset teams webhook url here
    [switch]$Email, #Email on completition 
    [string]$EmailTo, #Who should get the email
    [string]$EmailSMTP, #Internal smtp server to bounce email off
    [switch]$Force #Allows bypass of file check
)

#Import SCCM Modules and setup run location
Import-Module "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
$smsClient = New-Object -ComObject Microsoft.SMS.Client -Strict
$siteCode = $smsclient.GetAssignedSite()+":"
Set-Location $sitecode
#End SCCM site import
$currentDate = Get-Date
#Vars for teams reporting
$jonGIF = $TeamsImage
$webhookURL=$TeamsURL


<#
##Detect if new media creation is required

Requires the setup of this script to create backups of your task sequences.
Otherwise just pass the -Force flag to ignore this
https://gallery.technet.microsoft.com/SCCM-Task-Sequence-Backup-559868bd
#>
function checkBackup($isoPath){
    if($Force){
        return $true
    }else{
        $taskSequenceName = $TSName
        if($null -eq $backupSourceFolder){
            throw "Please declare a non-null backup source folder or use the -Force flag"
        }
        $backupsPath = Get-childItem  $backupSourceFolder | Where-Object name -like $taskSequenceName
        $backupsPath = ($backupsPath |Sort-Object -Descending).name
        $backupsPath = $backupSourceFolder + $backupsPath[0]
        $backupFiles = Get-ChildItem $backupsPath
        $backupDate = $backupFiles[0].LastWriteTime
        $isoDate = (Get-Item -Path $isoPath).LastWriteTime
        if($backupDate -gt $isoDate){
            return $true
        }else{
            return $false
        }
    }
}

#Generate basic data needed for creation
try{
    $bootDistro = Get-CMDistributionPointGroup -Name $DP | Get-CMDistributionPoint
}catch{
    $bootDistro = Get-CMDistributionPoint -SiteSystemServerName $DP
}
if($null -eq $bootDistro){
    throw "Unable to find valid Distribution Point or Group"
}

switch ($mediaType) {
    "WinPE" {
        if($null -eq $isoname){
            $imagePath = $Folder + "\WinPE-"+$currentDate.DayOfYear+"-"+$currentDate.Year +".iso"
        }else{
            $imagePath = $Folder + "\" + $isoname +".iso"
        }
        $managementGroup=Get-CMManagementPoint
        $bootImage = Get-CMBootImage -Name $WinPEName
        if($bootImage.count -gt 1){
            throw "Found multiple boot images, pleae refine your WinPEName"
        }
        New-CMBootableMedia -MediaMode Dynamic `
            -AllowUnknownMachine `
            -MediaType CdDvd `
            -BootImage $bootimage `
            -DistributionPoints $bootDistro `
            -ManagementPoint $managementGroup `
            -Path $imagePath `
            -Force
        $endTime = new-timespan -end $(Get-Date) -Start $currentDate
        $fact1= New-TeamsFact -Name "Boot ISO" -Value "Created a new boot ISO at $imagePath in $endTime" -ErrorAction Ignore
    }
    "TaskSequence" {
        $taskSequence = Get-CMTaskSequence $TSName
        if($taskSequence.count -gt 1){
            throw "Found multiple task sequences, please refine your TSName"
        }
        if($null -eq $isoname){
            $imagePath = $Folder + "\"+ $taskSequence.name +"-"+$currentDate.DayOfYear+"-"+$currentDate.Year +".iso"
        }else{
            $imagePath = $Folder + "\" + $isoname +".iso"
        }
        if(checkBackup -isoPath $imagePath){
            New-CMStandaloneMedia -MediaType CdDvd `
                -MediaSize SizeUnlimited `
                -Path $imagePath `
                -TaskSequence $taskSequence `
                -DistributionPoints $bootDistro `
                -Force
            $endTime = new-timespan -end $(Get-Date) -Start $currentDate
            $fact1= New-TeamsFact -Name "Full ISO" -Value "Created a new Full ISO at $imagePath in $endTime" -ErrorAction Ignore
        }else{
            throw "Existing image does not need updated"
        }
    }
}

#Teams message block
if($Teams){
    $Section1 = New-TeamsSection `
        -ActivityTitle "ISO Created" `
        -ActivitySubtitle $CurrentDate `
        -ActivityImageLink $jonGIF `
        -ActivityText "" `
        -ActivityDetails $fact1
	
    Send-TeamsMessage `
        -URI $webhookURL `
        -MessageTitle 'ISO Created' `
        -MessageText "This text won't show up" `
        -Color DodgerBlue `
        -Sections $Section1
}

#Email message block
if($Email){
    $body = "A new ISO has been created with the following file path( $imagePath )<br> It took  $endTime to run."
    $dataBlock=@{
        Body = $body
        Subject = "ISO Creation Script - Complete"
        BodyAsHtml = $True
        From = "ISOMaker@example.com"
        To = $EmailTo
        SmtpServer = $EmailSMTP
    }
    Send-MailMessage @dataBlock
}