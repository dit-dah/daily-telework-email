# daily-telework-email
Powershell script to open daily telework start and end email templates

# Installation
1.	Create a folder for the files (default is “C:\src\Telework Emails”)
2.	Clone the repository from GitHub into the folder
3.	If you haven’t yet, create telework start and a telework end email templates. Save them wherever you choose.
    * TODO: implement HTML file and string options, configure recipient
4.	Edit the .json configuration file:
    1. StartTemplate: URI for the Outlook template file for telework start
    2. EndTemplate: URI for the Outlook template file for telework end
    3. SSIDs: Array of potential wifi SSIDs that your laptop connects to when you log in. (For example, ["Home Wifi", "Home Wifi 2"])
    4. Domain: Domain name for your VPN. Can be found by running `echo $env:USERDOMAIN` in Powershell when connected to VPN
    5. LogPath: Optional configuration - alternate path for logging
5.	Open up task scheduler and create the following scheduled task: -- NOTE: a sample .xml file has been provided, but will NOT work if imported into Task Scheduler
    1. Under [General > Security options] change the account to your account and run only when you’re logged on
    2. Under [Triggers], create two triggers:
        * Weekly at 15 minutes prior to your scheduled logoff time (e.g. for a standard 8:00 - 5:00 day, set to 4:45 PM)
        * On Event - Log: System, Source: Rasman, Event ID: 20267. This is the windows Remote Access Connection Manager event for successfully connecting to a VPN. Verify in Event Viewer at [Windows Logs > System] and searching for “RasMan” (Remote Access Connection Manager).
    3. Under [Actions], create an action set to the following:  -- NOTE: the double curly braces below are to denote configuration and should not be in the final strings
        * Action: Start a program
        * Command: `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`
        * Arguments: `-NoProfile -WindowStyle Hidden -File {{"path\to\telework-emails.ps1"}}` --NOTE: if you want to specify a non-default config file, add `-ConfigPath {{path\to\config}}` to the end of this string
        * WorkingDirectory: `{{telework\emails\folder}}`
6.	Save the scheduled task and run it manually to confirm it works. You should see an email based off your template pop up.