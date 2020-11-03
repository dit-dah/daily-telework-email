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
    $log = '.\telework-emails.log'
    if (!(Test-Path $log)) { Set-Content -Path $log -Value '' }
    Add-Content -Path $log -Value $log_msg
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
$config = Get-Content -Path '.\telework-emails.config.json' | ConvertFrom-Json
$t = Get-Date
Log('----------------------')
Log(('Start Log at ' + $t))
# Determine which template to use, by hour of the day
Log(('parameters are as follows:\nStartTemplate:\t'+$config.StartTemplate+'\nEndTemplate:\t'+$config.EndTemplate+'\nSSID:\t'+$config.SSID))
# if($null -eq $StartTemplate -or $null -eq $EndTemplate -or $null -eq $SSIDs) {
#     ExitWithMessage('Insufficient parameters', 1)
# }
if (($t - (Get-Date 08:00)) -lt ((Get-Date 17:00) - $t)) {
    $template = $config.StartTemplate
} else {
    $template = $config.EndTemplate
}
Log(('Template file: ' + $template))
#get SSID for wifi
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
if (($g -notin $config.SSIDs) -or $env:USERDOMAIN -ne 'ESD1') { 
    ExitWithMessage('Not working from home')
}

# If the program hasn't exited, open the outlook item
$ol = New-Object -ComObject Outlook.Application
$msg = $ol.CreateItemFromTemplate($template)
$msg.GetInspector.Display()
ExitWithMessage('Done', 0)