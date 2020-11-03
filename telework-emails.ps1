#Define parameters and configuration
param($ConfigPath = 'telework-emails.config.json')
$config = Get-Content -Path $ConfigPath | ConvertFrom-Json

#---------------------- Function Declarations ----------------------
<#
.SYNOPSIS
Logs a message to a file

.DESCRIPTION
Logs the given string message to a file in the same directory as the executed file called 'telework-emails.log'

.PARAMETER log_msg
String message to log

.EXAMPLE
Log('Starting execution')
#>
function Log {
    param ($log_msg)
    if (!(Test-Path $config.LogPath)) { Set-Content -Path $config.LogPath -Value '' }
    Add-Content -Path $config.LogPath -Value $log_msg
}
<#
.SYNOPSIS
Exits script with a given message

.DESCRIPTION
Logs a given message and then exits with the given exit code.
Default message is blank and default exit code is 0 (success).

.PARAMETER exit_msg
String message to log at exit

.PARAMETER code
Exit code

.EXAMPLE
ExitWithMessage('execution complete', 0)
#>
function ExitWithMessage {
    param ($exit_msg = '', $code = 0)
    Log($exit_msg)
    exit $code
}

#---------------------- Script starts here ----------------------
Log('----------------------')
Log(('Start Log at ' + $t))
Log(('parameters are as follows:\nStartTemplate:\t'+$config.StartTemplate+'\nEndTemplate:\t'+$config.EndTemplate+'\nSSIDs:\t'+$config.SSIDs+'\nDomain:\t'+$config.Domain))
# Determine which template to use, by hour of the day
$t = Get-Date
if (($t - (Get-Date 08:00)) -lt ((Get-Date 17:00) - $t)) {
    $template = $config.StartTemplate
} else {
    $template = $config.EndTemplate
}
Log(('Template file: ' + $template))
#get SSIDs for wifi
$a = netsh wlan show interfaces | Select-String '\sSSID\s+:\s(.*)'
# Find matches
if ($a.Matches.Count -eq 0) { ExitWithMessage('No Matches', 1) }
# Get last match
$m = $a.Matches[$a.Matches.Count - 1]
# Get groups
if ($m.Groups.Count -eq 0) { ExitWithMessage('No groups in the match', 1) }
# Get last group
$g = $m.Groups[$m.Groups.Count - 1].Value

# Write it out
Log(('Wi-Fi Network: ' + $g))
$c = Get-NetAdapter | Where-Object Status -eq 'Up' | Select-Object -exp Name
Log(('Adapter: ' + $c))
Log(('Domain: ' + $env:USERDOMAIN))

# Check to see if working from home
if (($g -notin $config.SSIDs) -or $env:USERDOMAIN -ne $config.Domain) { 
    ExitWithMessage('Not working from home')
}

# If the program hasn't exited, open the outlook item
$ol = New-Object -ComObject Outlook.Application
$msg = $ol.CreateItemFromTemplate($template)
$msg.GetInspector.Display()
ExitWithMessage('Done', 0)