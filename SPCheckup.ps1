<#
.DESCRIPTION
	Provide daily SharePoint farm health checkup report.  Push email with HTML tables to SharePoint adminstrator 
	with high level summary of key featuress to ensure service online and available for users.  
	Monitors SharePoint sepcific features to keep high level pulse for system health.

	Daily push notification proviudes positive feedback loop to ensure farm well maintained over time.
	
	Comments and suggestions always welcome.

.NOTES  
	File Name:  SPCheckup.ps1
	Author   :  Jeff Jones  - @spjeff
	Version  :  1.0
	Modified :  2020-01-20

.LINK
	https://github.com/spjeff/spcheckup
#>
[CmdletBinding()]
param (
    [Alias("i")]
	[switch]$install,
	[string]$smtpServer,
	[string]$smtpFrom,
	[string]$smtpTo
)

Function GetRegKey($name, $value) {
	# Read registry key (if exists) and create if missing
	$key = 'HKLM:\SOFTWARE\SPCheckup'
	New-Item $key -ErrorAction SilentlyContinue | Out-Null
	$read = Get-ItemProperty -Path $key
	if (![string]::IsNullOrEmpty($value)) {
		Get-Item $key | New-ItemProperty -Name $name -Value $value -Force | Out-Null
		return $value
	} else {
		return $read."$name"
	}
}

Function SendMail($html) {
	# Send email message
	$smtp = GetRegKey "smtpServer" $smtpServer
	$from = GetRegKey "smtpFrom" $smtpFrom
	$to =  GetRegKey "smtpTo" $smtpTo
	Send-MailMessage -To $to -From $from -SmtpServer $smtp -Body $html -BodyAsHtml -Subject "SPCheckup" -Attachments $file
}

Function FormatTable($coll) {
	$temp = $coll | ConvertTo-Html -Fragment
	$temp = $temp.Replace('<table','<table border=1 class="spc-table"')
	$temp += "</br></br>"
	return ($temp -join " ")
}

Function FormatHTML() {
	# Header
	$html = @"
	<html>
	<head>
		<title>SPCheckup</title>
	</head>
	<body>
	
	<style>
		.spc-red {
			background-color: red;
		}
		.spc-yellow {
			background-color: yellow;
		}
		.spc-green {
			background-color: limegreen;
		}
		.spc-table td {
			padding:5px;
		}
		.spc-table tr:first-child th {
			font-weight: bold;
		}
	</style>
"@

	# Collection
	$s = Get-SPServer
	$wa = Get-SPWebApplication
	$sa = Get-SPServiceApplication
	$cdb = Get-SPContentDatabase
	$wp = Get-SPSolution

	# Merge HTML
	$html += "<h2>Farm</h2>"
	$html += FormatTable (CollectFarm)
	$html += "<h2>Server  $($s.Count)</h2>"
	$html += FormatTable (CollectServer)
	$html += "<h2>Web App  $($wa.Count)</h2>"
	$html += FormatTable (CollectWebApp)
	$html += "<h2>Service App  $($sa.Count)</h2>"
	$html += FormatTable (CollectSvcApp)
	$html += "<h2>Content DB $($cdb.Count)</h2>"
	$html += FormatTable (CollectContentDatabase)
	$html += "<h2>Solution  $($wsp.Count)</h2>"
	$html += FormatTable (CollectSolution)
	$html += "<hr>Generated at $(Get-Date)"

	# TBD - Project Web Access (PWA)
	# TBE - Search
	# TBD - User Profiles
	# TBD - web.config cutomerrors

	# Footer
	$html += "</body></html>"
	return $html
}

function CollectSvcApp() {
	# Input
	$svcapp = Get-SPServiceApplication | Sort Name

	$coll = @()
	foreach ($sa in $svcapp) {

		# Collect
		$obj = New-Object -Type PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $sa.DisplayName
		$obj | Add-Member -MemberType NoteProperty -Name "Status" -Value $sa.Status
		$obj | Add-Member -MemberType NoteProperty -Name "NeedsUpgrade" -Value $sa.NeedsUpgrade
		$coll += $obj
	}
	return $coll
}

function CollectContentDatabase() {
	# Input
	$cdb = Get-SPContentDatabase | Sort Name

	$coll = @()
	foreach ($c in $cdb) {

		# Collect
		$obj = New-Object -Type PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $c.Name
		$obj | Add-Member -MemberType NoteProperty -Name "SQL" -Value $c.NormalizedDataSource
		$obj | Add-Member -MemberType NoteProperty -Name "NeedsUpgrade" -Value $c.NeedsUpgrade
		$obj | Add-Member -MemberType NoteProperty -Name "Sites" -Value $c.Sites
		$obj | Add-Member -MemberType NoteProperty -Name "WarnSites" -Value $c.WarnSites
		$obj | Add-Member -MemberType NoteProperty -Name "MaxSites" -Value $c.MaxSites
		$coll += $obj
	}
	return $coll
}

function CollectSolution() {
	# Input
	$wsp = Get-SPSolution | Sort Name

	$coll = @()
	foreach ($w in $wsp) {
		# Collect
		$obj = New-Object -Type PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $w.Name
		$obj | Add-Member -MemberType NoteProperty -Name "DeploymentState" -Value $w.DeploymentState
		$obj | Add-Member -MemberType NoteProperty -Name "ContainsWebApplicationResource" -Value $w.ContainsWebApplicationResource
		$obj | Add-Member -MemberType NoteProperty -Name "LastOperationResult" -Value $w.LastOperationResult
		$obj | Add-Member -MemberType NoteProperty -Name "LastOperationEndTime" -Value $w.LastOperationEndTime
		$coll += $obj
	}
	return $coll
}

function CollectWebApp() {
	# Input
	$webapps = Get-SPWebApplication | Sort Url

	$coll = @()
	foreach ($wa in $webapps) {
		$meas = Measure-Command {Invoke-WebRequest -UseDefaultCredentials -Uri $wa.Url }
		$resp = Invoke-WebRequest -UseDefaultCredentials -Uri $wa.Url

		# Collect
		$obj = New-Object -Type PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "URL" -Value $wa.Url
		$obj | Add-Member -MemberType NoteProperty -Name "HTTP" -Value $resp.StatusCode
		$obj | Add-Member -MemberType NoteProperty -Name "HTTP (Sec)" -Value ([math]::round($meas.TotalSeconds,2))
		$coll += $obj
	}
	return $coll
}

function CollectFarm() {
	# Input
	$f = Get-SPFarm
	$disabledTimers = $f.TimerService.Instances |? {$_.Status -ne "Online"}

	$log = Get-SPLogLevel |? {$_.TraceSev -eq "Verbose" -or $_.EventSev -eq "Verbose"}
	$log |% {$verboseULS += "$($_.Area) - $($_.Name)"}

	# Collect
	$coll = @()
	$obj = New-Object -Type PSObject
	$obj | Add-Member -MemberType NoteProperty -Name "ConfigDB" -Value $f.Name
	$obj | Add-Member -MemberType NoteProperty -Name "BuildVersion" -Value $f.BuildVersion
	$obj | Add-Member -MemberType NoteProperty -Name "NeedsUpgrade" -Value $f.NeedsUpgrade
	$obj | Add-Member -MemberType NoteProperty -Name "DisabledTimers" -Value $f.DisabledTimers
	$obj | Add-Member -MemberType NoteProperty -Name "CEIPEnabled" -Value $f.CEIPEnabled	
	$obj | Add-Member -MemberType NoteProperty -Name "VerboseULS" -Value $verboseULS
	$coll += $obj
	return $coll
}

function CollectServer() {
	# Input
	$coll = @()
	$servers = Get-SPServer | Sort Name
	foreach ($s in $servers) {
		# WMI Uptime
		$uptime=""
		$wmi = Get-WmiObject -ComputerName $s.Name -Class Win32_OperatingSystem -ErrorAction SilentlyContinue
		if ($wmi) {
			$t = $wmi.ConvertToDateTime($wmi.LocalDateTime) - $wmi.ConvertToDateTime($wmi.LastBootUpTime)
			$uptime = "$($t.Days) : $($t.Hours) : $($t.Minutes)"
		}

		# RAM
		$ram = $wmi | Select-Object {$f}, TotalVisibleMemorySize, @{Name="Prct";Expression={[math]::round(100-(($_.FreePhysicalMemory/$_.TotalVisibleMemorySize)*100),0)}}

		# Drives
		$drives = ""
		try {
			$wmidisk = Get-WmiObject -ComputerName $s.Name Win32_LogicalDisk -ErrorAction SilentlyContinue |? {$_.DeviceID -ne 'A:'} |? {$_.DeviceID -ne 'V:'} | Select DeviceID,VolumeName,@{n='FreeSpace';e={[math]::round($_.FreeSpace/1GB,2)}},@{n='Size';e={[math]::round($_.Size/1GB,2)}},@{n='Prct';e={[math]::round((($_.FreeSpace/$_.Size)*100),0)}}
		} catch {}
		$wmidisk |% {
			$drives += $_.DeviceID + "  ("+ $_.Prct +" %)"
		}

		# IIS Pools Started
		try {
		$pool = Get-WmiObject -ComputerName $s.Name IIsApplicationPoolSetting -namespace 'root\microsoftiisv2' -ErrorAction SilentlyContinue
		} catch {}
		$poolTotal = $pool.Count
		$poolStarted = ($pool |? {$_.AppPoolState -ne "Started"}).Count

		# IIS Sites Started
		try {
			$sites = Get-WmiObject -ComputerName $s.Name IIsWebServer -namespace 'root\microsoftiisv2' -ErrorAction SilentlyContinue
		} catch {}
		$sitesTotal = $sites.Count
		$sitesStarted = ($sites |? {$_.ServerState -ne "Started"}).Count

		# NeedsUpgrade
		$NeedsUpgrade = $s.$NeedsUpgrade
		if ($s.Role -eq "Invalid") {
			$NeedsUpgrade = ""
		}

		# Collect
		$obj = New-Object -Type PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Name" -Value $s.Name
		$obj | Add-Member -MemberType NoteProperty -Name "Role" -Value $s.Role.ToString().Replace("Invalid","")
		$obj | Add-Member -MemberType NoteProperty -Name "Status" -Value $s.Status
		$obj | Add-Member -MemberType NoteProperty -Name "NeedsUpgrade" -Value $NeedsUpgrade

		$obj | Add-Member -MemberType NoteProperty -Name "UpTime (DD:HH:MM:)" -Value $uptime
		if ($ram) {
			$obj | Add-Member -MemberType NoteProperty -Name "RAM" -Value "$($ram.TotalVisibleMemorySize/(1024*1024)) GB ($($ram.Prct) %)"
		}
		$obj | Add-Member -MemberType NoteProperty -Name "Disk" -Value $drives
		$obj | Add-Member -MemberType NoteProperty -Name "IIS-Websites" -Value "$sitesStarted/$sitesTotal"
		$obj | Add-Member -MemberType NoteProperty -Name "IIS-Pool" -Value "$poolStarted/$poolTotal"
		$coll += $obj
	}
	return $coll
}

function CleanHTML($html) {
	$html = $html.Replace(">Online</td>"," bgcolor='limegreen'>Online</a>").Replace(">True</td>"," bgcolor='limegreen'>True</a>").Replace(">NotDeployed</td>"," bgcolor='yellow'>NotDeployed</a>").Replace("<th>"," <th bgcolor='lightblue'>").Replace(">200"," bgcolor='limegreen'>200").Replace(">WebApplicationDeployed"," bgcolor='limegreen'>WebApplicationDeployed").Replace(">GlobalDeployed"," bgcolor='limegreen'>GlobalDeployed").Replace(">GlobalAndWebApplicationDeployed"," bgcolor='limegreen'>GlobalAndWebApplicationDeployed")
	return $html
}

Function Main() {
    # Load plugins
	Add-PSSnapIn Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null
	Import-Module WebAdministration -ErrorAction SilentlyContinue | Out-Null

	# Attachment
	$farmid = (Get-SPFarm).Id
	$stamp = (Get-Date).ToString().Replace('/','-').Replace(':','').Replace(' ','')
	$file = "SPCheckup-$farmid-$stamp.html"
	
	# Format HTML
	$html = FormatHTML
	$html = CleanHTML $html
	$html | Out-File $file -Force

	# Send mail
	SendMail $html
	Remove-Item $file -Confirm:$false -Force
}
Main
