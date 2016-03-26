$scriptVersion = "1.0"
Write-Host "=================================================" -ForegroundColor Yellow
Write-Host "           Get-IISSmtpState v$scriptVersion  " -ForegroundColor Yellow
Write-Host "         june.castillote@gmail.com           " -ForegroundColor Yellow
Write-Host "=================================================" -ForegroundColor Yellow
#http://shaking-off-the-cobwebs.blogspot.com/
Write-Host ''
Write-Host (Get-Date) ': Begin' -ForegroundColor Green
Write-Host (Get-Date) ': Setting Paths and Variables' -ForegroundColor Yellow
#$ErrorActionPreference="SilentlyContinue"
$WarningPreference="SilentlyContinue";

#Server names to be checked, seperate with comma ","
$SmtpServers = "smtp1,smtp2"

#>>Define Variables---------------------------------------------------------------
$errSummary = ""
$today = '{0:dd-MMM-yyyy hh:mm tt}' -f (Get-Date)
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$css_string = '<style type="text/css"> #HeadingInfo { font-family:Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #HeadingInfo td, #HeadingInfo th { font-size:0.9em; padding:3px 7px 2px 7px; } #HeadingInfo th  { font-size:1.0em; font-weight:bold; text-align:center; padding-top:5px; padding-bottom:4px; background-color:#CC3300; color:#fff; } #SectionLabels { font-family:Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #SectionLabels th.data { font-size:0.8em; text-align:center; padding-top:5px; padding-bottom:4px; background-color:#A7C942; color:#fff; } #data { font-family:Consolas,Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #data td, #data th  { font-size:0.8em; border:1px solid #98bf21; padding:3px 7px 2px 7px; } #data th  { font-size:0.8em; padding-top:5px; padding-bottom:4px; background-color:#A7C942; color:#fff; text-align:left; } #data td { font-size:0.8em; padding-top:5px; padding-bottom:4px; text-align:left; } #data td.bad { font-size:0.8em; font-weight: bold; padding-top:5px; padding-bottom:4px; background-color:red; } #data td.good { font-size:0.8em; font-weight: bold; padding-top:5px; padding-bottom:4px; color:green; }</style> </head> <body> <hr />'
$reportfile = $script_root + "\IISSmtpReport_" + ('{0:dd_MMM_yyyy}' -f (Get-Date)) + ".html"
#>>------------------------------------------------------------------------------


#>>Thresholds--------------------------------------------------------------------
[int]$Remote_Queue = 20
[int]$Local_Queue = 0
#>>------------------------------------------------------------------------------


#>>Options, set to $false if you do not want to send the report------------------
$SendReportViaEmail = $true
#>>------------------------------------------------------------------------------
#>>Mail
$CompanyName = 'ABC'
$MailSubject = 'IIS Smtp Server Report '
$MailServer = 'smtp.relay.here'
$MailSender = 'ABC postmaster <postmaster@abc.com>'
$MailTo = 'administrator@abc.com'
$MailCC = '' #if you specify a CC address, make sure to uncomment the CC line in the $params variable block
$MailBCC = '' #if you specify a BCC address, make sure to uncomment the BCC line in the $params variable block
#>>------------------------------------------------------------------------------

#convert $SmtpServers string to array
$SmtpServers = $SmtpServers.Split(",")

Function Get-SmtpStats {
Write-Host (Get-Date) ': Getting Status... ' -ForegroundColor Yellow -NoNewLine
$stats_collection = @()
	foreach ($smtpserver in $SmtpServers)
	{
		$x = Get-Service SMTPSVC -ComputerName $smtpserver
		
		if ($x -ne $null) {
			$smtpObj = "" | Select Server,ServiceStatus,Received,Sent,LocalQueue,RemoteQueue
			$smtpObj.Server = $smtpserver
			$smtpObj.ServiceStatus = $x.Status
			$smtpObj.Received = (get-counter -counter '\SMTP Server(_Total)\Messages Received Total' -maxsamples 1 -ComputerName $smtpserver | Select-Object -ExpandProperty countersamples).cookedValue
			$smtpObj.Sent = (get-counter -counter '\SMTP Server(_Total)\Messages Sent Total' -maxsamples 1 -ComputerName $smtpserver | Select-Object -ExpandProperty countersamples).cookedValue
			$smtpObj.LocalQueue = (get-counter -counter '\SMTP Server(_Total)\Local Queue Length' -maxsamples 1 -ComputerName $smtpserver | Select-Object -ExpandProperty countersamples).cookedValue
			$smtpObj.RemoteQueue = (get-counter -counter '\SMTP Server(_Total)\Remote Queue Length' -maxsamples 1 -ComputerName $smtpserver | Select-Object -ExpandProperty countersamples).cookedValue
			
		}
		elseif ($x -eq $null){
			$smtpObj = "" | Select Server,ServiceStatus,Received,Sent,LocalQueue,RemoteQueue
			$smtpObj.Server = $smtpserver
			$smtpObj.ServiceStatus = "Could not get SMTPSVC Status. Check if the server $smtpServer is up"
			$smtpObj.Received = "-"
			$smtpObj.Sent = "-"
			$smtpObj.LocalQueue = "-"
			$smtpObj.RemoteQueue = "-"
			
		}
	$stats_collection + $smtpObj
	}
Write-Host "Done" -ForegroundColor Green
return $stats_collection
}

Function Create-SmtpStatsReport ($smtpStats) {
Write-Host (Get-Date) ': Creating Report... ' -ForegroundColor Yellow -NoNewLine
$mbody = @()
$errString = @()
$currentServer = ""
$mbody += '<table id="SectionLabels"><tr><th class="data">SMTP Server Service Status</th></tr></table>'
$mbody += '<table id="data">'
$mbody += '<tr><th>Server</th><th>Service Status</th><th>Messages Received</th><th>Messages Sent</th><th>Local Queue</th><th>Remote Queue</th></tr>'
	foreach ($smtpData in $smtpStats) {
		$mbody += "<tr><td>$($smtpData.Server)</td>"
		
		if ($smtpData.ServiceStatus -ne 'Running') {
			$errString += "<tr><td>Service Status</td></td><td>$($smtpData.Server) - $($smtpData.ServiceStatus)</td></tr>"
			$mbody += "<td class = ""bad"">$($smtpData.ServiceStatus)</td>"
		}
		elseif ($smtpData.ServiceStatus -eq 'Running') {
			$mbody += "<td class = ""good"">$($smtpData.ServiceStatus)</td>"
		}
		
		if ($smtpData.Received -eq $null){
			$errString += "<tr><td>Performance Counter Object</td></td><td>$($smtpData.Server) - There was an error getting the counter object. Please make sure the SMTP Virtual Server instance is running</td></tr>"
			$mbody += "<td class = ""bad"">There was an error getting the counter object. Please make sure the SMTP Virtual Server instance is running</td>"
			$mbody += "<td class = ""bad"">There was an error getting the counter object. Please make sure the SMTP Virtual Server instance is running</td>"
			$mbody += "<td class = ""bad"">There was an error getting the counter object. Please make sure the SMTP Virtual Server instance is running</td>"
			$mbody += "<td class = ""bad"">There was an error getting the counter object. Please make sure the SMTP Virtual Server instance is running</td>"
		}
		else {
			$mbody += "<td>$($smtpData.Received)</td>"
			$mbody += "<td>$($smtpData.Sent)</td>"
			
			if ($smtpData.LocalQueue -ge $Local_Queue) {
			$errString += "<tr><td>Local Queue Lenght</td></td><td>$($smtpData.LocalQueue) is >= $Local_Queue </td></tr>"
			$mbody += "<td class = ""bad"">$($smtpData.LocalQueue)</td>"					
			}
			else{
				$mbody += "<td class = ""good"">$($smtpData.LocalQueue)</td>"
			}
			
			if ($smtpData.RemoteQueue -ge $Remote_Queue) {
				$errString += "<tr><td>Remote Queue Lenght</td></td><td>$($smtpData.RemoteQueue) is >= $Remote_Queue </td></tr>"
				$mbody += "<td class = ""bad"">$($smtpData.RemoteQueue)</td>"		
			}
			else{
				$mbody += "<td class = ""good"">$($smtpData.RemoteQueue)</td>"
			}
		}
		
		
		$mbody += '</tr>'
	}
Write-Host "Done" -ForegroundColor Green
return $mbody,$errString
}

#>>----------------------------------------------------------------------------
Write-Host '==================================================================' -ForegroundColor Green
Write-Host (Get-Date) ': Begin' -ForegroundColor Yellow
$smtpHealthData = Get-SmtpStats($SmtpServers)
$smtpHealthReport,$sError = Create-SmtpStatsReport ($smtpHealthData) ; $errSummary += $sError

if ($errSummary -eq "") {
	$finalSubject = "[$($CompanyName)] $($MailSubject) $($today)"
}
else {
	$finalSubject = "ALERT!!! [$($CompanyName)] $($MailSubject) $($today)"
}

$mail_body = "<html><head><title>$finalSubject</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
Write-Host (Get-Date) ': Applying CSS to HTML Report' -ForegroundColor Yellow
$mail_body += $css_string
$mail_body += '<table id="HeadingInfo">'
$mail_body += '<tr><th>' + $CompanyName + '<br />' + $finalSubject + '<br />' + $today + '</th></tr>'
$mail_body += '</table><hr />'
$mail_body += '<table id="SectionLabels">'
$mail_body += '<tr><th class="data">----SUMMARY----</th></tr></table>'
$mail_body += '<table id="data"><tr><th>Check Item</th><th>Details</th></tr>'
$mail_body += $errSummary
$mail_body += '</table><hr />'
$mail_body += $smtpHealthReport ; $mail_body += '</table><hr />'
$mail_body += '<p>'
$mail_body += '<table id="SectionLabels">'
$mail_body += '<tr><th>----END of REPORT----</th></tr></table><hr />'
$mail_body += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
$mail_body += '<b>[THRESHOLD]</b><br />'
$mail_body += 'Local Queue: ' +  $Local_Queue + ' hours<br />'
$mail_body += 'Remote Queue: ' + $Remote_Queue + ' hours<br />'
$mail_body += '<b>[MAIL]</b><br />'
$mail_body += 'SMTP Server: ' + $MailServer + '<br /><br />'
$mail_body += '<b>[REPORT]</b><br />'
$mail_body += 'Generated from Server: ' + (gc env:computername) + '<br />'
$mail_body += 'Script Path: ' + $script_root
$mail_body += '<p>'
$mail_body += "<a href='http://shaking-off-the-cobwebs.blogspot.com/'>IIS Smtp State v$scriptVersion</a>"
$mbody = $mbox -replace "&lt;","<"
$mbody = $mbox -replace "&gt;",">"
$mail_body | Out-File $reportfile
Write-Host (Get-Date) ': HTML Report saved to file -' $reportfile -ForegroundColor Yellow
#>>----------------------------------------------------------------------------
#>> Mail Parameters------------------------------------------------------------
#>> Add CC= and/or BCC= lines if you want to add recipients for CC and BCC



$params = @{
    Body = $mail_body
    BodyAsHtml = $true
    Subject = $finalSubject
    From = $MailSender
	To = $MailTo.Split(",")
    SmtpServer = $MailServer
	#Cc = $MailCC.Split(",")
	#Bcc = $MailBCC.Split(",")
}
#>>----------------------------------------------------------------------------
#>> Send Report----------------------------------------------------------------
if ($SendReportViaEmail -eq $true) {Write-Host (Get-Date) ': Sending Report' -ForegroundColor Yellow ; Send-MailMessage @params}
#>>----------------------------------------------------------------------------
Write-Host (Get-Date) ': End' -ForegroundColor Green
#>>SCRIPT END------------------------------------------------------------------