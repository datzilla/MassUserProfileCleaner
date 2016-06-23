Mass User Profiles Cleaner is an administrative tool written in VBScript & Batch Script to safely delete user profiles on remote computers. This tool has been tested to run on Windows 7 Enterprise. The tool utilises System Internal's 'PsEXEC.exe' to execute Batch and VBScript on the remote computer.

Requirements:
- Windows 7
- Files and Print Sharing and Admin Share enabled on remote computer
- Remote Management must be enabled on remote computer
- Administrative user account on remote computer


PC_List.txt contains all computers which you want to remove user profiles. The MassUserProfileCleaner.vbs VBScript will read PC_LIST.txt to obtain and cycle through the computer names to perform the profile cleaning task. The MassUserProfileCleaner.vbs VBScript will omit any computers that has ' character at the beginning of its name.

MassUserProfileCleaner.vbs VBScript copies the UserProfileCleaner folder to the remote computer's C:\ drive and execute run.cmd.
Run.cmd will execute UserProfileCleaner.vbs  with customised parameters. Make changes to run.cmd if you want to change the custom paramters.

UserProfileCleaner.vbs VBScript has the following input paramters:
- removeall : This paramter will remove all user profiles on the computer except for the account that the script is running under.
- exclude [username] : This paramter excludes the specified username
- defrag : This paramter tells the VBScript to system windows defragger to defrag the computer
- reboot : This paramter tells the VBScript to reboot the system after all tasks have been performed
