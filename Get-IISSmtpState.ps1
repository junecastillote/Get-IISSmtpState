
<#PSScriptInfo

.VERSION 2.1

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
2.0 (June 10, 2019)
	- Re-Code from scratch
2.1 (June 11, 2019)
	- Fixed bugs
	- Code cleanup
	- Added error handling logic


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




#...................................
#Region CSS
#...................................
$css_string = @'
<style type="text/css">
#HeadingInfo 
	{
		font-family:"Segoe UI";
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
		font-family:"Segoe UI";
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
		font-family:"Segoe UI";
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
Function Get-TimeZoneInfo
{  
	$tzName = ([System.TimeZone]::CurrentTimeZone).StandardName
	$tzInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($tzName)
	Return $tzInfo	
}
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

#Get TimeZone Information
$timeZoneInfo = Get-TimeZoneInfo

[string]$today = Get-Date -Format F
$today = "$($today) $($timeZoneInfo.DisplayName.ToString().Split(" ")[0])"

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
	#$prop = "" | Select-Object Computer,QueueDirectory,PickupDirectory,BadMailDirectory,DropDirectory,Service,QueueCount,PickupCount,BadMailCount,DropCount,QueueSize,PickupSize,BadMailSize,DropSize,QueueStatus,PickupStatus,BadMailStatus,DropStatus,ServiceStatus,ServerStatus,CheckItems

	$prop = [ordered]@{
		Computer=""
		QueueDirectory=""
		PickupDirectory=""
		BadMailDirectory=""
		DropDirectory=""
		Service=""
		QueueCount=0
		PickupCount=0
		BadMailCount=0
		DropCount=0
		QueueSize=0
		PickupSize=0
		BadMailSize=0
		DropSize=0
		QueueStatus=""
		PickupStatus=""
		BadMailStatus=""
		DropStatus=""
		ServiceStatus=""
		ServerStatus=""
		CheckItems=@()
	}
	#NOTE: This script will only check the default folders
	$svcStatus = Get-Service -ComputerName $computer -Name SMTPSVC -ErrorAction SilentlyContinue -ErrorVariable svcErr
	$prop.Computer = $computer

	#all status start off as PASSED
	$prop.ServerStatus = "Passed"
	$prop.QueueStatus = "Passed"
	$prop.PickupStatus = "Passed"
	$prop.BadMailStatus = "Passed"
	$prop.DropStatus = "Passed"
	$prop.ServiceStatus = "Passed"
	$prop.CheckItems = @()

	#computer Status
	if (!$wmiErr)
	{
		$prop.QueueDirectory = "\\$($computer)\c$\inetpub\mailroot\queue"
		$prop.PickupDirectory = "\\$($computer)\c$\inetpub\mailroot\pickup"
		$prop.BadMailDirectory = "\\$($computer)\c$\inetpub\mailroot\badmail"
		$prop.DropDirectory = "\\$($computer)\c$\inetpub\mailroot\drop"

		$queue = Get-ChildItem $prop.QueueDirectory -ErrorAction SilentlyContinue -ErrorVariable queueVar | Measure-Object -property length -sum
		$pickup = Get-ChildItem $prop.PickupDirectory -ErrorAction SilentlyContinue -ErrorVariable pickupVar | Measure-Object -property length -sum
		$badmail = Get-ChildItem $prop.BadMailDirectory -ErrorAction SilentlyContinue -ErrorVariable badmailVar | Measure-Object -property length -sum
		$drop = Get-ChildItem $prop.DropDirectory -ErrorAction SilentlyContinue -ErrorVariable dropVar | Measure-Object -property length -sum


		#error checks
		if ($queueVar)
		{
			$prop.CheckItems += $queueVar.Exception.Message
			$prop.CheckItems += "Error retrieving queue"
			$prop.QueueDirectory = $queueVar.Exception.Message
			$prop.QueueCount = 0
			$prop.QueueSize = 0
			$prop.QueueStatus = "Failed"
			$prop.ServerStatus = "Failed"
		}
		else {
			$prop.QueueCount = [math]::Round($queue.count)
			$prop.QueueSize = [math]::Round(($queue.sum) / 1KB)
		}

		if ($pickupVar)
		{
			$prop.CheckItems += $pickupVar.Exception.Message
			$prop.CheckItems += "Error retrieving pickup"
			$prop.PickupDirectory = $pickupVar.Exception.Message
			$prop.pickupCount = 0
			$prop.pickupSize = 0
			$prop.PickupStatus = "Failed"
			$prop.ServerStatus = "Failed"
		}
		else {
			$prop.pickupCount = [math]::Round($pickup.count)
			$prop.pickupSize = [math]::Round(($pickup.sum) / 1KB)
		}

		if ($badmailVar)
		{
			$prop.CheckItems += $badmailVar.Exception.Message
			$prop.CheckItems += "Error retrieving badmail"
			$prop.BadMailDirectory = $badmailVar.Exception.Message
			$prop.badmailCount = 0
			$prop.badmailSize = 0
			$prop.BadMailStatus = "Failed"
			$prop.ServerStatus = "Failed"
		}
		else {
			$prop.badmailCount = [math]::Round($badMail.count)
			$prop.badmailSize = [math]::Round(($badMail.sum) / 1KB)
		}

		if ($dropVar)
		{
			$prop.CheckItems += $dropVar.Exception.Message
			$prop.CheckItems += "Error retrieving drop"
			$prop.DropDirectory = $dropVar.Exception.Message
			$prop.DropCount = 0
			$prop.dropSize = 0
			$prop.DropStatus = "Failed"
			$prop.ServerStatus = "Failed"
		}
		else {
			$prop.DropCount = [math]::Round($drop.count)
			$prop.dropSize = [math]::Round(($drop.sum) / 1KB)
		}
		
		#queue threshold tripped
		if ($queueThreshold -and ($queue.Count) -gt $queueThreshold)
		{
			$prop.QueueStatus = "Failed"
			$prop.ServerStatus = "Failed"
			$prop.CheckItems += "Queue Count is $($queue.Count) which is over the theshold of $($queueThreshold)"
		}

		#pickup threshold tripped
		if ($pickupThreshold -and ($pickup.count) -gt $pickupThreshold)
		{
			$prop.PickupStatus = "Failed"
			$prop.ServerStatus = "Failed"
			$prop.CheckItems += "Pickup Count is $($pickup.Count) which is over the theshold of $($pickupThreshold)"
		}

		#badmail threshold tripped
		if ($badMailThreshold -and ($badmail.count) -gt $badMailThreshold)
		{
			$prop.BadMailStatus = "Failed"
			$prop.ServerStatus = "Failed"
			$prop.CheckItems += "BadMail Count is $($badmail.Count) which is over the theshold of $($badMailThreshold)"
		}

		#drop threshold tripped
		if ($dropThreshold -and ($drop.Count) -gt $dropThreshold)
		{
			$prop.DropStatus = "Failed"
			$prop.ServerStatus = "Failed"
			$prop.CheckItems += "Drop Count is $($drop.Count) which is over the theshold of $($dropThreshold)"
		}
	}
	else 
	{
		$prop.QueueDirectory = $wmiErr.Exception.Message
		$prop.PickupDirectory = $wmiErr.Exception.Message
		$prop.BadMailDirectory = $wmiErr.Exception.Message
		$prop.DropDirectory = $wmiErr.Exception.Message
		$prop.QueueStatus = "Failed"
		$prop.PickupStatus = "Failed"
		$prop.BadMailStatus = "Failed"
		$prop.DropStatus = "Failed"
		$prop.ServerStatus = "Failed"
		$prop.QueueCount = 0
		$prop.QueueSize = 0
		$prop.PickupCount = 0
		$prop.PickupSize = 0
		$prop.BadMailCount = 0
		$prop.BadMailSize = 0
		$prop.DropCount = 0
		$prop.DropSize = 0
		$prop.CheckItems += ("WMI Error: " +$wmiErr.Exception.Message)
	}
	
	#Service Status
	if (!$svcErr)
	{
		$prop.Service = $svcStatus.Status	
		if ($prop.Service -ne 'Running')
		{
			$prop.ServiceStatus = "Failed"
			$prop.ServerStatus = "Failed"
			$prop.CheckItems += "SMTP Service is not in 'Running' state"			
		}
	}
	else {
		$prop.Service = $svcErr.Exception.Message
		$prop.ServiceStatus = "Failed"
		$prop.ServerStatus = "Failed"
		$prop.CheckItems += ("Service Error: " + $svcErr.Exception.Message)
		
	}

	$obj = New-Object PSObject -Property $prop
	
	$serverCollection += $obj
	
}

$serverCollection
#...................................
#EndRegion COLLECT IIS SMTP SERVER DETAILS
#...................................

#...................................
#Region WRITE REPORT
#...................................


$failedServers = $serverCollection | Where-Object {$_.ServerStatus -ne 'Passed'}
#Write-Host ($failedServers | Measure-Object).Count
$mailSubject = "IIS Virtual SMTP Service Report | $($today)"
if ($orgName)
{
	if ($failedServers)
	{
		$mailSubject = "ALERT!!! - [$($orgName)] | IIS Virtual SMTP Service Report | $($today)"
	}
	else
	{
		$mailSubject = "[$($orgName)] | IIS Virtual SMTP Service Report | $($today)"	
	}
}
else 
{
	if ($failedServers)
	{
		$mailSubject = "ALERT!!! - IIS Virtual SMTP Service Report | $($today)"
	}
	else
	{
		$mailSubject = "IIS Virtual SMTP Service Report | $($today)"	
	}
}


$htmlBody = "<html><head><title>"
if ($orgName) 
{
	$header = "$($orgName)<br />IIS Virtual SMTP Service Report<br />$($today)"
}
else 
{
	$header = "IIS Virtual SMTP Service Report<br />$($today)"	
}
$htmlBody += "</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
$htmlBody += $css_string
$htmlBody += '</head><body>'
$htmlBody += '<table id="HeadingInfo">'
$htmlBody += '<tr><th>'+$header+'</th></tr>'
$htmlBody += '</table><hr />'

$htmlBody += '<table id="SectionLabels">'
$htmlBody += '<tr><th class="data">Issue Summary</th></tr></table>'
$htmlBody += '<table id="data"><tr><th>Computer</th><th>Details</th></tr>'

if ($failedServers)
{
	foreach ($s in $failedServers)
	{
		$htmlBody += '<tr><td class="bad">' + $s.Computer + '</td><td class="bad"><ol type="1"><li>'+ ($s.CheckItems -join "</li><li>")+'</li></ol></td>'
	}
}
else 
{
	$htmlBody += '<tr><td class="good">NO ISSUES</td><td class="good">NO ISSUES</td>'
}
$htmlBody += '</table><hr />'


$htmlBody += '<table id="SectionLabels">'
$htmlBody += '<tr><th class="data">Server Details</th></tr></table>'
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
		$htmlBody += '<td class="good">'+($s.QueueCount.ToString('N0'))+' / '+($s.QueueSize.ToString('N0'))+' KB</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">'+($s.QueueCount.ToString('N0'))+' / '+($s.QueueSize.ToString('N0'))+' KB</td>'
	}

	if ($s.PickupStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+($s.PickupCount.ToString('N0'))+' / '+($s.PickupSize.ToString('N0'))+' KB</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">'+($s.PickupCount.ToString('N0'))+' / '+($s.PickupSize.ToString('N0'))+' KB</td>'
	}

	if ($s.BadMailStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+($s.BadMailCount.ToString('N0'))+' / '+($s.BadMailSize.ToString('N0'))+' KB</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">'+($s.BadMailCount.ToString('N0'))+' / '+($s.BadMailSize.ToString('N0'))+' KB</td>'
	}

	if ($s.DropStatus -eq 'Passed')
	{
		$htmlBody += '<td class="good">'+($s.DropCount.ToString('N0'))+' / '+($s.DropSize.ToString('N0'))+' KB</td>'
	}
	else 
	{
		$htmlBody += '<td class="bad">'+($s.DropCount.ToString('N0'))+' / '+($s.DropSize.ToString('N0'))+' KB</td>'
	}	
}
$htmlBody += '</table><hr />'

$htmlBody += '<p><font size="2" face="Segoe UI"><b><center>----END of REPORT----</center></b><br />'
$htmlBody += '<p><font size="2" face="Segoe UI"><u>Report Paremeters</u><br />'
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
	
	if ($failedServers)
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
    if ($sendEmail -eq 'ErrorOnly' -and $failedServers)
    {
        Write-Host (get-date -Format "dd-MMM-yyyy hh:mm:ss tt") ": Sending email to" ($To -join ", ") -ForegroundColor Yellow
        Send-MailMessage @mailParams
    }
}
#...................................
#EndRegion SEND REPORT
#...................................
Stop-TxnLogging