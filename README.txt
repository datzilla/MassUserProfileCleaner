The MassUserProfilesCleaner App is located here:
N:\IT\Desktop Services\Helpdesk Scripts\MassUserProfilesCleaner

You will need to execute: N:\IT\Desktop Services\Helpdesk Scripts\MassUserProfilesCleaner\MassUserProfilesCleaner.vbs
As usual, you control which computer to execute by commenting/uncommenting computers names in PC_List.txt

You will need to copy this to your computer and share the Logs folder to everyone+writable. All executions on the remote computers will be saved centrally in that shared location.

The Batch files  that gets executed on the remote computer will run the profile cleaner VBScript. It’s located here:
N:\IT\Desktop Services\Helpdesk Scripts\MassUserProfilesCleaner\UserProfilesCleaner\run.cmd.

You can see the command in run.cmd append the screen output to \\FEB3025\LogShare\\%COMPUTERNAME%.log. Change the computer name and LogShare to reflect the correct configuration on your computer.
CSCRIPT /B /E:Cscript /NoLogo "%~DP0UserProfilesCleaner.vbs" removeall cleartemp defrag reboot >>\\FEB3025\LogShare\%COMPUTERNAME%.log

The script
-	removes all user profiles
-	clears temp
-	defrags
-	reboots at the end
-	excludes bloombergtraining