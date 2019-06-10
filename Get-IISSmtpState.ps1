
<#PSScriptInfo

.VERSION 2.0

.GUID 35b14c0b-4e9a-4111-9d9a-cfe6cf038219

.AUTHOR June Castillote

.COPYRIGHT june.castillote@gmail.com

.TAGS

.LICENSEURI

.PROJECTURI https://github.com/junecastillote/Get-IISSMTPState

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES


.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 IIS SMTP Server Status Report

#> 

param (
	[cmdletbinding()]

	# list of IIS SMTP Servers, accepts array.
	[Parameter(Mandatory=$true,Position=0)]
	[string[]]
	$computerName,

	#path to the report directory (eg. c:\scripts\report)
	[Parameter(Mandatory=$true,Position=1)]
	[string]
	$reportDirectory,

	#Threshold for Queue
	[Parameter()]
	[int]
	$queueThreshold,

	#Threshold for Pickup
	[Parameter()]
	[int]
	$pickupThreshold,

	#Threshold for Drop
	[Parameter()]
	[int]
	$dropThreshold,

	#Threshold for BadMail
	[Parameter()]
	[int]
	$badMailThreshold,

	#path to the log directory (eg. c:\scripts\logs)
	[Parameter()]
	[string]
	$logDirectory,

	#prefix string for the report (ex. COMPANY)
	[Parameter()]
	[string]
	$orgName,
	
	#Switch to enable email report
	[Parameter()]
    [ValidateSet("ErrorOnly","Always")]
    [string]
	$sendEmail,

	#Sender Email Address
	[Parameter()]
    [string]
	$From,

	#Recipient Email Addresses - separate with comma
	[Parameter()]
	[string[]]
	$To,

	#smtpServer
	[Parameter()]
	[string]
	$smtpServer,

	#Port
	[Parameter()]
	[int]
	$Port,

	#switch to indicate whether SMTP Authentication is required
	[Parameter()]
	[switch]
	$smtpAuthRequired,

	#credential for SMTP server (if applicable)
	[Parameter()]
	[pscredential]
	$smtpCredential,

	#switch to indicate if SSL will be used for SMTP relay
	[Parameter()]
	[switch]
    $useSSL
)

[string]$today = Get-Date -Format F

#...................................
#Region CSS
#...................................
$css_string = @'
<style type="text/css">
#HeadingInfo 
	{
		font-family:"Consolas";
		width:100%;
		border-collapse:collapse;
	} 
#HeadingInfo td, #HeadingInfo th 
	{
		font-size:0.8em;
		padding:3px 7px 2px 7px;
	} 
#HeadingInfo th  
	{ 
		font-size:2.0em;
		font-weight:normal;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#604767;
		color:#fff;
	} 
#SectionLabels
	{ 
		font-family:"Calibri";
		width:100%;
		border-collapse:collapse;
	}
#SectionLabels th.data
	{
		font-size:2.0em;
		text-align:left;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#fff;
		color:#000; 
	} 
#data 
	{
		font-family:"Calibri";
		width:100%;
		border-collapse:collapse;
	} 
#data td, #data th
	{ 
		font-size:0.8em;
		border:1px solid #DDD;
		padding:3px 7px 2px 7px; 
	} 
#data th  
	{
		font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
		background-color:#00B388;
		color:#fff; text-align:left;
	} 
#data td 
	{ 	font-size:0.8em;
		padding-top:5px;
		padding-bottom:4px;
		text-align:left;
	} 
#data td.bad
	{ 	font-size:0.8em;
		font-weight: bold;
		padding-top:5px;
		padding-bottom:4px;
		color:#f04953;
	} 
#data td.good
	{ 	font-size:0.8em;
		font-weight: bold;
		padding-top:5px;
		padding-bottom:4px;
		color:#01a982;
	}

.status {
	width: 10px;
	height: 10px;
	margin-right: 7px;
	margin-bottom: 0px;
	background-color: #CCC;
	background-position: center;
	opacity: 0.8;
	display: inline-block;
}
.green {
	background: #01a982;
}
.purple {
	background: #604767;
}
.orange {
	background: #ffd144;
}
.red {
	background: #f04953;
}
</style>
'@
#...................................
#Region CSS
#...................................

#...................................
#Region FUNCTIONS
#...................................
Function Stop-TxnLogging
{
	$txnLog=""
	Do {
		try {
			Stop-Transcript | Out-Null
		} 
		catch [System.InvalidOperationException]{
			$txnLog="stopped"
		}
    } While ($txnLog -ne "stopped")
}

#Function to Start Transaction Logging
Function Start-TxnLogging 
{
    param 
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$logDirectory
    )
	Stop-TxnLogging
    Start-Transcript $logDirectory -Append
}

#Function to get Script Version and ProjectURI for PSv4
Function Get-ScriptInfo
{
    param 
    (
        [Parameter(Mandatory=$true,Position=0)]
        [string]$Path
	)
	
	$props = @{
		Version = (Select-String -Pattern ".VERSION" -Path $Path)[0].ToString().split(" ")[1]
		ProjectURI = (Select-String -Pattern ".PROJECTURI" -Path $Path)[0].ToString().split(" ")[1]
	}
	$scriptInfo = New-Object PSObject -Property $props
    Return $scriptInfo
}
#...................................
#EndRegion
#...................................


Stop-TxnLogging
Clear-Host

#Get Script Information
if ($PSVersionTable.PSVersion.Major -lt 5)
{
	$scriptInfo = Get-ScriptInfo -Path $MyInvocation.MyCommand.Definition
}
else
{
	$scriptInfo = Test-ScriptFileInfo -Path $MyInvocation.MyCommand.Definition	
}

#...................................
#Region PARAMETER CHECK
#...................................
$isAllGood = $true

if ($sendEmail)
{
    if (!$From)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: A valid sender email address is not specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$To)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No recipients specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$smtpServer )
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No SMTP Server specified." -ForegroundColor Yellow
        $isAllGood = $false
    }

    if (!$Port )
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: No SMTP Port specified." -ForegroundColor Yellow
        $isAllGood = $false
	}
	
	if ($smtpAuthRequired)
	{
		if (!$smtpCredential)
		{
			Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: SMTP Server requires authentication, but no credential was specified. Please specify using the -smtpCredential parameter." -ForegroundColor Yellow
        	$isAllGood = $false
		}
	}
}

if ($isAllGood -eq $false)
{
    Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": ERROR: Exiting Script." -ForegroundColor Yellow
    EXIT
}
#...................................
#EndRegion PARAMETER CHECK
#...................................

#...................................
#Region PATHS
#...................................
$logFile = $logDirectory +"\Log_$((get-date).tostring("yyyy_MMM_dd")).log"
$outputHTMLFile = $reportDirectory +"\IIS_SMTPServer_Report_$((get-date).tostring("yyyy_MMM_dd")).html"

#Create folders if not found
if ($logDirectory)
{
    if (!(Test-Path $logDirectory)) 
    {
        New-Item -ItemType Directory -Path $logDirectory | Out-Null
        #start transcribing
        Start-TxnLogging $logFile
        
    }
	else
	{
		Start-TxnLogging $logFile
	}
}

if (!(Test-Path $reportDirectory))
{
	New-Item -ItemType Directory -Path $reportDirectory | Out-Null
}
#...................................
#EndRegion PATHS
#...................................

#...................................
#Region COLLECT IIS SMTP SERVER DETAILS
#...................................
$serverCollection = @()
foreach ($computer in $computerName)
{
	Write-host (Get-Date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Processing $($computer)" -ForegroundColor Yellow
	$temp = "" | Select-Object Computer,QueueDirectory,PickupDirectory,BadMailDirectory,DropDirectory,Service,QueueCount,PickupCount,BadMailCount,DropCount,QueueSize,PickupSize,BadMailSize,DropSize,QueueStatus,PickupStatus,BadMailStatus,DropStatus,ServiceStatus,ServerStatus,CheckItems

	#NOTE: This script will only check the instance smtpsvc/1 - which is the default smtp virtual server
	$n = Get-WmiObject -ComputerName $computer -Namespace "root\MicrosoftIISV2" -Class IISSMTPSERVERSETTING -ErrorAction SilentlyContinue -ErrorVariable wmiErr | Where-Object {$_.name -eq 'smtpsvc/1'}
	$svcStatus = Get-Service -ComputerName $computer -Name SMTPSVC -ErrorAction SilentlyContinue -ErrorVariable svcErr
	$temp.Computer = $computer

	#all status start off as PASSED
	$temp.ServerStatus = "Passed"
	$temp.QueueStatus = "Passed"
	$temp.PickupStatus = "Passed"
	$temp.BadMailStatus = "Passed"
	$temp.DropStatus = "Passed"
	$temp.ServiceStatus = "Passed"
	$temp.CheckItems = @()

	#computer Status
	if (!$wmiErr)
	{
		$temp.QueueDirectory = "\\$($computer)\" + ($n.QueueDirectory -replace ":","$")
		$temp.PickupDirectory = "\\$($computer)\" + ($n.PickupDirectory -replace ":","$")
		$temp.BadMailDirectory = "\\$($computer)\" + ($n.BadMailDirectory -replace ":","$")
		$temp.DropDirectory = "\\$($computer)\" + ($n.DropDirectory -replace ":","$")

		$queue = Get-ChildItem $temp.QueueDirectory -ErrorAction SilentlyContinue -ErrorVariable queueVar | Measure-Object -property length -sum
		$pickup = Get-ChildItem $temp.PickupDirectory -ErrorAction SilentlyContinue -ErrorVariable pickupVar | Measure-Object -property length -sum
		$badmail = Get-ChildItem $temp.BadMailDirectory -ErrorAction SilentlyContinue -ErrorVariable badmailVar | Measure-Object -property length -sum
		$drop = Get-ChildItem $temp.DropDirectory -ErrorAction SilentlyContinue -ErrorVariable dropVar | Measure-Object -property length -sum


		#error checks
		if ($queueVar)
		{
			$temp.CheckItems += $queueVar.Exception.Message
			$temp.QueueDirectory = $queueVar.Exception.Message
			$temp.QueueCount = 0
			$temp.QueueSize = 0
		}
		else {
			$temp.QueueCount = $queue.count
			$temp.QueueSize = ($queue.sum) / 1KB
		}

		if ($pickupVar)
		{
			$temp.CheckItems += $pickupVar.Exception.Message
			$temp.PickupDirectory = $pickupVar.Exception.Message
			$temp.pickupCount = 0
			$temp.pickupSize = 0
		}
		else {
			$temp.pickupCount = $pickup.count
			$temp.pickupSize = ($pickup.sum) / 1KB
		}

		if ($badmailVar)
		{
			$temp.CheckItems += $badmailVar.Exception.Message
			$temp.BadMailDirectory = $badmailVar.Exception.Message
			$temp.badmailCount = 0
			$temp.badmailSize = 0
		}
		else {
			$temp.badmailCount = $badMail.count
			$temp.badmailSize = ($badMail.sum) / 1KB
		}

		if ($dropVar)
		{
			$temp.CheckItems += $dropVar.Exception.Message
			$temp.DropDirectory = $dropVar.Exception.Message
			$temp.DropCount = "0"
			$temp.dropSize = 0
		}
		else {
			$temp.DropCount = $drop.count
			$temp.dropSize = ($drop.sum) / 1KB
		}
		
		#queue threshold tripped
		if ($queueThreshold -and ($queue.Count) -gt $queueThreshold)
		{
			$temp.QueueStatus = "Failed"
			$temp.ServerStatus = "Failed"
			$temp.CheckItems += "Queue Count is $($queue.Count) which is over the theshold of $($queueThreshold)"
		}

		#pickup threshold tripped
		if ($pickupThreshold -and ($pickup.count) -gt $pickupThreshold)
		{
			$temp.PickupStatus = "Failed"
			$temp.ServerStatus = "Failed"
			$temp.CheckItems += "Pickup Count is $($pickup.Count) which is over the theshold of $($pickupThreshold)"
		}

		#badmail threshold tripped
		if ($badMailThreshold -and ($badmail.count) -gt $badMailThreshold)
		{
			$temp.BadMailStatus = "Failed"
			$temp.ServerStatus = "Failed"
			$temp.CheckItems += "BadMail Count is $($badmail.Count) which is over the theshold of $($badMailThreshold)"
		}

		#drop threshold tripped
		if ($dropThreshold -and ($drop.Count) -gt $dropThreshold)
		{
			$temp.DropStatus = "Failed"
			$temp.ServerStatus = "Failed"
			$temp.CheckItems += "Drop Count is $($drop.Count) which is over the theshold of $($dropThreshold)"
		}
	}
	else {
		$temp.QueueDirectory = $wmiErr.Exception.Message
		$temp.PickupDirectory = $wmiErr.Exception.Message
		$temp.BadMailDirectory = $wmiErr.Exception.Message
		$temp.DropDirectory = $wmiErr.Exception.Message
		$temp.QueueStatus = "Failed"
		$temp.PickupStatus = "Failed"
		$temp.BadMailStatus = "Failed"
		$temp.DropStatus = "Failed"
		$temp.ServerStatus = "Failed"
		$temp.QueueCount = 0
		$temp.QueueSize = 0
		$temp.PickupCount = 0
		$temp.PickupSize = 0
		$temp.BadMailCount = 0
		$temp.BadMailSize = 0
		$temp.DropCount = 0
		$temp.DropSize = 0
		$temp.CheckItems += ("WMI Error: " +$wmiErr.Exception.Message)
	}
	
	#Service Status
	if (!$svcErr)
	{
		$temp.Service = $svcStatus.Status	
		if ($temp.Service -ne 'Running')
		{
			$temp.ServiceStatus = "Failed"
			$temp.ServerStatus = "Failed"
			$temp.CheckItems += "SMTP Service is not in 'Running' state"			
		}
	}
	else {
		$temp.Service = $svcErr.Exception.Message
		$temp.ServiceStatus = "Failed"
		$temp.ServerStatus = "Failed"
		$temp.CheckItems += ("Service Error: " + $svcErr.Exception.Message)
		
	}
	
	$serverCollection += $temp
	
}

$serverCollection
#...................................
#EndRegion COLLECT IIS SMTP SERVER DETAILS
#...................................

#...................................
#Region WRITE REPORT
#...................................

$failedServerCount = ($serverCollection | Where-Object {$_.ServerStatus -eq 'Failed'}).count
$mailSubject = "IIS SMTP Service Report | $($today)"
if ($orgName)
{
	if ($failedServerCount -gt 1)
	{
		$mailSubject = "ALERT!!! - [$($orgName)] | IIS SMTP Service Report | $($today)"
	}
	else
	{
		$mailSubject = "[$($orgName)] | IIS SMTP Service Report | $($today)"	
	}
}


$htmlBody = "<html><head><title>"
if ($orgName) 
{
	$header = "[$($orgName)]<br />IIS SMTP Service Report<br />$($today)"
}
else 
{
	$header = "IIS SMTP Service Report<br />$($today)"	
}
$htmlBody += "</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
$htmlBody += $css_string
$htmlBody += '</head><body>'
$htmlBody += '<table id="HeadingInfo">'
$htmlBody += '<tr><th>'+$header+'</th></tr>'
$htmlBody += '</table><hr />'

$htmlBody += '<table id="SectionLabels">'
$htmlBody += '<tr><th class="data">----Issues Summary----</th></tr></table>'
$htmlBody += '<table id="data"><tr><th>Check Item</th><th>Details</th></tr>'

foreach ($s in ($serverCollection | Where-Object {$_.ServerStatus -eq 'Failed'}))
{
	
	$htmlBody += '<tr><td>' + $s.Computer + '</td><td><ol type="1"><li>'+ ($s.CheckItems -join "</li><li>")+'</li></ol></td>'
}
$htmlBody += '</table><hr />'


$htmlBody += '<table id="SectionLabels">'
$htmlBody += '<tr><th class="data">----Iis Smtp Server Details----</th></tr></table>'
$htmlBody += '<table id="data">'
$htmlBody += '<tr><th>Computer</th><th>Service</th><th>Queue [Count/Size (KB)]</th><th>Pickup [Count/Size (KB)]</th><th>BadMail [Count/Size (KB)]</th><th>Drop [Count/Size (KB)]</th></tr>'
foreach ($s in $serverCollection)
{	
	
	$htmlBody += '<tr><td>'+$s.Computer+'</td>'

	if ($s.ServiceStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+ $s.ServiceStatus + '</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">'+ $s.ServiceStatus + '</td>'
	}

	if ($s.QueueStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+$s.QueueCount+'/'+$s.QueueSize+'</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">Failed</td>'
	}

	if ($s.PickupStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+$s.PickupCount+'/'+$s.PickupSize+'</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">Failed</td>'
	}

	if ($s.BadMailStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+$s.BadMailCount+'/'+$s.BadMailSize+'</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">Failed</td>'
	}

	if ($s.DropStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+$s.DropCount+'/'+$s.DropSize+'</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">Failed</td>'
	}	
}
$htmlBody += '</table><hr />'

#$htmlBody += '<table id="SectionLabels">'
#$htmlBody += '<tr><th class = "data">----END of REPORT----</th></tr></table><hr />'
$htmlBody += '<p><font size="2" face="Tahoma"><b><center>----END of REPORT----</center></b><br />'
$htmlBody += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
$htmlBody += '<b>[THRESHOLD]</b><br />'
$htmlBody += 'Queue: ' +  $queueThreshold + '<br />'
$htmlBody += 'Pickup: ' + $pickupThreshold + '<br />'
$htmlBody += 'BadMail: ' + $badMailThreshold + '<br />'
$htmlBody += 'Drop: ' + $dropThreshold + '<br />'
$htmlBody += '<br /><b>[MAIL]</b><br />'
$htmlBody += 'SMTP Server: ' + $smtpServer + '<br />'
$htmlBody += 'Port: ' + $Port + '<br />'
$htmlBody += 'SSL: ' + $useSSL + '<br />'
$htmlBody += 'Authentication: ' + $smtpAuthRequired + '<br />'
$htmlBody += '<br /><b>[REPORT]</b><br />'
$htmlBody += "Generated from Server: $($env:COMPUTERNAME)<br />"
$htmlBody += 'Script Path: ' + $MyInvocation.MyCommand.Definition
$htmlBody += '<p>'
$htmlBody += '<a href="'+ $scriptInfo.ProjectURI +'">IIS Smtp State '+ $scriptInfo.Version +'</a>'

$htmlBody += '</body></html>'
$htmlBody | Out-File $outputHTMLFile
Write-host (Get-Date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Report saved in $($outputHTMLFile)" -ForegroundColor Yellow
#...................................
#EndRegion WRITE REPORT
#...................................

#...................................
#Region SEND REPORT
#...................................
if ($sendEmail)
{
	[string]$mailBody = Get-Content $outputHTMLFile -Raw
	$mailParams = @{
        From = $From
        To = $To
        smtpServer = $smtpServer
        Port = $Port
        useSSL = $useSSL
        body = $mailBody
		bodyashtml = $true
		subject = $mailSubject
	}
	
	if ($failedServerCount -gt 1)
	{
		$mailParams += @{priority = "HIGH"}
	}
	else 
	{
		$mailParams += @{priority = "LOW"}
	}

    
    if ($smtpAuthRequired)
    {
        $mailParams += @{credential = $smtpCredential}
    }

    #Always
    if ($sendEmail -eq 'Always')
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending email to" ($To -join ", ") -ForegroundColor Yellow
        Send-MailMessage @mailParams
    }

    #ErrorOnly AND failedServerCount
    if ($sendEmail -eq 'ErrorOnly' -and $failedServerCount -gt 1)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending email to" ($To -join ", ") -ForegroundColor Yellow
        Send-MailMessage @mailParams
    }
}
#...................................
#EndRegion SEND REPORT
#...................................
Stop-TxnLogging