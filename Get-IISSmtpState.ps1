$scriptVersion = "1.1"
Write-Host "=================================================" -ForegroundColor Yellow
Write-Host "           Get-IISSmtpState v$scriptVersion  " -ForegroundColor Yellow
Write-Host "         june.castillote@gmail.com           " -ForegroundColor Yellow
Write-Host "=================================================" -ForegroundColor Yellow
#https://www.lazyexchangeadmin.com/2016/03/iis-smtp-server-status-check-powershell.html
Write-Host ''
Write-Host (Get-Date) ': Begin' -ForegroundColor Green
Write-Host (Get-Date) ': Setting Paths and Variables' -ForegroundColor Yellow
$ErrorActionPreference="SilentlyContinue"
$WarningPreference="SilentlyContinue"

#Server names to be checked, seperate with comma ","
$SmtpServers = "smtp1,smtp2"

#>>Define Variables---------------------------------------------------------------
$errSummary = ""
$today = '{0:dd-MMM-yyyy hh:mm tt}' -f (Get-Date)
$script_root = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$css_string = '<style type="text/css"> #HeadingInfo { font-family:Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #HeadingInfo td, #HeadingInfo th { font-size:0.9em; padding:3px 7px 2px 7px; } #HeadingInfo th  { font-size:1.0em; font-weight:bold; text-align:left; padding-top:5px; padding-bottom:4px; background-color:#fff; color:#808080; } #SectionLabels { font-family:Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #SectionLabels th.data { font-size:0.8em; text-align:center; padding-top:5px; padding-bottom:4px; background-color:#fff; color:#808080; } #data { font-family:Consolas,Tahoma, "Trebuchet MS", Arial, Helvetica, sans-serif; width:100%; border-collapse:collapse; } #data td, #data th  { font-size:0.8em; border:1px solid #808080; padding:3px 7px 2px 7px; } #data th  { font-size:0.8em; padding-top:5px; padding-bottom:4px; background-color:#808080; color:#fff; text-align:left; } #data td { font-size:0.8em; padding-top:5px; padding-bottom:4px; text-align:left; } #data td.bad { font-size:0.8em; font-weight: bold; padding-top:5px; padding-bottom:4px; background-color:#808080; color:#fff } #data td.good { font-size:0.8em; font-weight: bold; padding-top:5px; padding-bottom:4px; color:#808080; }</style> </head> <body> <hr />'
$reportfile = $script_root + "\IISSmtpReport_" + ('{0:dd_MMM_yyyy}' -f (Get-Date)) + ".html"
#>>------------------------------------------------------------------------------


#>>Thresholds--------------------------------------------------------------------
$queue_count = 1
$pickup_count = 1
#$drop_count = 5
#$badmail_count = 5
#>>------------------------------------------------------------------------------


#>>Options, set to $false if you do not want to send the report------------------
$SendReportViaEmail = $true
#>>------------------------------------------------------------------------------
#>>Mail
$CompanyName = 'LazyExchangeAdmin.com'
$MailSubject = 'IIS Smtp Server Report '
$MailServer = 'smtp1'
$MailSender = 'mailer <mailer@LazyExchangeAdmin.com>'
$MailTo = 'administrator@LazyExchangeAdmin.com'
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
		$queue = Get-ChildItem "\\$($smtpserver)\C$\inetpub\mailroot\Queue" | Measure-Object -property length -sum
		$pickup = Get-ChildItem "\\$($smtpserver)\C$\inetpub\mailroot\Pickup" | Measure-Object -property length -sum
		$badmail = Get-ChildItem "\\$($smtpserver)\C$\inetpub\mailroot\Badmail" | Measure-Object -property length -sum
		$drop = Get-ChildItem "\\$($smtpserver)\C$\inetpub\mailroot\Drop" | Measure-Object -property length -sum
		
		$smtpObj = "" | Select Server,ServiceStatus,Queue,Pickup,Drop,BadMail
		if ($x -ne $null) {
			
			$smtpObj.Server = $smtpserver
			$smtpObj.ServiceStatus = $x.Status
			$smtpObj.Queue = $queue
			$smtpObj.Pickup = $pickup
			$smtpObj.Drop = $drop
			$smtpObj.BadMail = $badmail
			
		}
		elseif ($x -eq $null){
			$smtpObj.Server = $smtpserver
			$smtpObj.ServiceStatus = "Could not get SMTPSVC Status. Check if the server $smtpServer is up and the service is running"
			$smtpObj.Queue = "-"
			$smtpObj.Pickup = "-"
			$smtpObj.Drop = "-"
			$smtpObj.BadMail = "-"
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
#$mbody += '<table id="SectionLabels"><tr><th class="data">SMTP Server Service Status</th></tr></table>'
$mbody += '<table id="data">'
$mbody += '<tr><th>Server</th><th>Service Status</th><th>Queue</th><th>Pickup</th><th>Drop</th><th>BadMail</th></tr>'
	foreach ($smtpData in $smtpStats) {
		$mbody += "<tr><td class = ""good"">$($smtpData.Server)</td>"
		
		if ($smtpData.ServiceStatus -ne 'Running') {
			$errString += "<tr><td>Service Status</td></td><td>$($smtpData.Server) - $($smtpData.ServiceStatus)</td></tr>"
			$mbody += "<td class = ""bad"">$($smtpData.ServiceStatus)</td>"
			$mbody += "<td>-</td>"
			$mbody += "<td>-</td>"
			$mbody += "<td>-</td>"
			$mbody += "<td>-</td>"
		}
		elseif ($smtpData.ServiceStatus -eq 'Running') {
			$mbody += "<td class = ""good"">$($smtpData.ServiceStatus)</td>"
			
			if ($smtpData.Queue.count -gt $queue_count) {
				$errString += "<tr><td>Queue Count</td></td><td>$($smtpData.Server) - Items in Queue has breached the threshold of [$($queue_count)]</td></tr>"
				$mbody += "<td class = ""bad"">" + [int]$smtpData.Queue.count + " [" + [int]($smtpData.Queue.sum / 1KB) + " KB]</td>"
			}
			else {
				$mbody += "<td class = ""good"">" + [int]$smtpData.Queue.count + " [" + [int]($smtpData.Queue.sum / 1KB) + " KB]</td>"
			}
			
			if ($smtpData.Pickup.count -gt $pickup_count) {
				$errString += "<tr><td>Queue Count</td></td><td>$($smtpData.Server) - Items in Pickup has breached the threshold of [$($pickup_count)]</td></tr>"
				$mbody += "<td class = ""bad"">" + [int]$smtpData.Pickup.count + " [" + [int]($smtpData.Pickup.sum / 1KB) + " KB]</td>"
			}
			else {
				$mbody += "<td class = ""good"">" + [int]$smtpData.Pickup.count + " [" + [int]($smtpData.Pickup.sum / 1KB) + " KB]</td>"
			}
	
			$mbody += "<td class = ""good"">" + [int]$smtpData.Drop.count + " [" + [int]($smtpData.Drop.sum / 1KB) + " KB]</td>"
			$mbody += "<td class = ""good"">" + [int]$smtpData.BadMail.count + " [" + [int]($smtpData.BadMail.sum / 1KB) + " KB]</td>"
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
	$finalSubject = "$($MailSubject) $($today)"
}
else {
	$finalSubject = "ALERT!!! $($MailSubject) $($today)"
}

$mail_body = "<html><head><title>$finalSubject</title><meta http-equiv=""Content-Type"" content=""text/html; charset=ISO-8859-1"" />"
Write-Host (Get-Date) ': Applying CSS to HTML Report' -ForegroundColor Yellow
$mail_body += $css_string
$mail_body += '<table id="HeadingInfo">'
$mail_body += '<tr><th>' + $CompanyName + '<br />' + $MailSubject + '<br />' + $today + '</th></tr>'
$mail_body += '</table><hr />'

if ($errSummary -ne "") {
$mail_body += '<table id="SectionLabels">'
$mail_body += '<tr><th class="data">----Issues Summary----</th></tr></table>'
$mail_body += '<table id="data"><tr><th>Check Item</th><th>Details</th></tr>'
$mail_body += $errSummary
$mail_body += '</table><hr />'	
}

$mail_body += $smtpHealthReport ; $mail_body += '</table><hr />'
$mail_body += '<table id="SectionLabels">'
$mail_body += '<tr><th class = "data">----END of REPORT----</th></tr></table><hr />'
$mail_body += '<p><font size="2" face="Tahoma"><u>Report Paremeters</u><br />'
$mail_body += '<b>[THRESHOLD]</b><br />'
$mail_body += 'Queue: ' +  $queue_count + '<br />'
$mail_body += 'Pickup: ' + $pickup_count + '<br />'
$mail_body += '<b>[MAIL]</b><br />'
$mail_body += 'SMTP Server: ' + $MailServer + '<br />'
$mail_body += '<b>[REPORT]</b><br />'
$mail_body += 'Generated from Server: ' + (gc env:computername) + '<br />'
$mail_body += 'Script Path: ' + $script_root
$mail_body += '<p>'
$mail_body += "<a href='https://www.lazyexchangeadmin.com/2016/03/iis-smtp-server-status-check-powershell.html'>IIS Smtp State v$scriptVersion</a>"
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