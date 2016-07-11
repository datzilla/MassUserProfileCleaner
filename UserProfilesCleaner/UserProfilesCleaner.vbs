' Script Name: UserProfileCleaners.vbs
' Version: 0.05
' Desc: This script removes and clean up user profiles on Windows 7 x64 systems
' Company: University of Sydney
' Author: Quoc Dat Nguyen
'==============================================================================
' CHANGE LOG
'==============================================================================
' Date			Ver		Description
' -----------|-------|---------------------------------------------------------
' 19/10/2015	0.05	Script created
' 23/06/2016    0.10    Made Profile Delete more efficient for removing all users
' 11/07/2016 	0.15	Added ability to exclude more than one profile
'==============================================================================

option Explicit

' Declare and create objects
' --------------------------
Dim objShell, objFSO, objArgs
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments

' Declare and set global variables
' --------------------------------
Dim blnVerbosity
Dim blnLastModified
Dim blnRemoveAll
Dim blnClearTemp
Dim blnSingleUser
Dim blnDefrag
Dim blnExcludeUser
Dim blnReboot
Dim blnBackupUser
Dim blnRestoreUser
Dim intExitCode
Dim intShowWindow

Dim strArg
Dim strLastModifiedDate
Dim strExcludeUser
Dim strTargetUser

blnVerbosity = True ' default is true
intShowWindow = 1
intExitCode = 0
blnRemoveAll = False
blnClearTemp = False
blnSingleUser = False
blnDefrag = False
blnLastModified = False
blnReboot = False
blnBackupUser = False
blnRestoreUser = False

Const HKEY_LOCAL_MACHINE = &H80000002

Dim strSourcePath : strSourcePath = objFSO.GetParentFolderName (WScript.ScriptFullName)
IncludeFile strSourcePath & "\Win7DeployCommon.vbs"

Dim strWindowsUninstallRegPath : strWindowsUninstallRegPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"	
Const HKLM = &H80000002	
Dim arrProfileKeys : arrProfileKeys = Array ()
Dim objReg, objExcludeUsers
' You can get to the profile list either from 64bit path or 32bit path. Both will get you the list on either platform.
Set objReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
Set objExcludeUsers = CreateObject ("Scripting.Dictionary")
Dim strProfileKey, strPath, objProfileFolder, objNTUserFile
Dim strReg32 : strReg32 = GetRegCmd (32)
Dim strReg64 : strReg64 = GetRegCmd (64)

' Check for quiet parameter input from commandline
' ------------------------------------------------
Dim pos
For pos = 0 To objArgs.Count - 1
	If (LCase (objArgs.Item (pos)) = "quiet") Then
		blnVerbosity = False
		intShowWindow = 0
	End If

	If (LCase (objArgs.Item (pos)) = "lastmodified") Then
		If pos = objArgs.Count - 1 Then WScript.Quit 1
		On Error Resume Next
			strLastModifiedDate = CDate (objArgs.Item (pos + 1))
			If Err.Number <> 0 Then
				WScript.Quit 1
			End If
		On Error Goto 0
		blnLastModified = True
	End If
	
	If (LCase (objArgs.Item (pos)) = "user") Then
		If pos = objArgs.Count - 1 Then WScript.Quit 1
		strTargetUser = objArgs.Item (pos + 1)
		blnSingleUser = True
	End If
	
	If (LCase (objArgs.Item (pos)) = "restore") Then
		If pos = objArgs.Count - 1 Then WScript.Quit 1
		dctExcludeUsers = ProcessExcludeUsers (objArgs.Item (pos + 1))
		blnRestoreUser = True
	End If
	
	If (LCase (objArgs.Item (pos)) = "backup") Then
		If pos = objArgs.Count - 1 Then WScript.Quit 1
		strTargetUser = objArgs.Item (pos + 1)
		blnBackupUser = True
	End If
		
	If (LCase (objArgs.Item (pos)) = "exclude") Then
		If pos = objArgs.Count - 1 Then WScript.Quit 1
		'strExcludeUser = objArgs.Item (pos + 1)
		'Modified to handle more than one exclusive profile
		objExcludeUsers.Add LCase (objArgs.Item (pos + 1)), Null
		blnExcludeUser = True
	End If
	
	If (LCase (objArgs.Item (pos)) = "removeall") Then
		blnRemoveAll = True
	End If
	
	If (LCase (objArgs.Item (pos)) = "cleartemp") Then
		blnClearTemp = True
	End If
	
	If (LCase (objArgs.Item (pos)) = "reboot") Then
		blnReboot = True
	End If
	
	If (LCase (objArgs.Item (pos)) = "defrag") Then
		blnDefrag = True
	End If		
Next

' It isn't possible to have these combinations of  switches at the same time.
If blnSingleUser = True And blnRemoveAll = True Then
	WScript.Echo "Can't delete all profiles and also delete a single profile"
	WScript.Quit 1
ElseIf blnSingleUser = True And blnLastModified = True Then 
	WScript.Echo "Can't delete last modified profiles and also delete a single profile"
	WScript.Quit 1
ElseIf blnRemoveAll = True And  blnLastModified = True Then
	WScript.Echo "You can't specify delete all and also last modified date"
	WScript.Quit 1
ElseIf blnSingleUser = True And objArgs.Count < 2 Then
	WScript.Echo "You didn't specify the user profile to delete"
	WScript.Quit 1
ElseIf blnLastModified = True And objArgs.Count < 2 Then
	WScript.Echo "You didn't specify the last modified date to delete"
	WScript.Quit 1
ElseIf blnSingleUser = True And blnExcludeUser = True Then
	WScript.Echo "Can't exclude user and also target a single user"
	WScript.Quit 1
End If

' DELETE SINGLE USER
' ------------------
If blnSingleUser = True And blnRemoveAll = False Then
	WScript.Echo "Deleting single"
	objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrProfileKeys
	For Each strProfileKey In arrProfileKeys
		If Len (strProfileKey) > 8 Then  
			objReg.GetExpandedStringValue HKLM, strWindowsUninstallRegPath & "\" & strProfileKey, "ProfileImagePath", strPath
			strPath = objShell.ExpandEnvironmentStrings (strPath)
			If objFSO.FolderExists (strPath) = True Then
				Set objProfileFolder = objFSO.GetFolder (strPath)
				If StrComp (LCase (objProfileFolder.Name), LCase (strTargetUser)) = 0 Then
					objShell.Run """" & strReg32 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					objShell.Run """" & strReg64 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					WScript.Echo "deleted: " & strWindowsUninstallRegPath & "\" & strProfileKey 
					'objReg.DeleteKey HKLM, strWindowsUninstallRegPath & "\" & strProfileKey
					objShell.Run "%COMSPEC% /C Rd /S /Q """ & strPath & """", 0, True
					WScript.Echo "deleted: " & strPath
				End If
			Else			
				objShell.Run """" & strReg32 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
				objShell.Run """" & strReg64 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
				WScript.Echo "Deleted: " & strWindowsUninstallRegPath & "\" & strProfileKey
			End If
		End If
	Next
	
	' If there's folder but not registry extry for the user name, we can delete it.
	strPath = objShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%") & "\Users\" & strTargetUser
	If objFSO.FolderExists (strPath) = True Then
		objShell.Run "%COMSPEC% /C Rd /S /Q """ & strPath & """", 0, True
		WScript.Echo "deleted: " & strPath
	End If
End If

' BACKUP PROFILE
' --------------
If blnBackupUser = True Then
	WScript.Echo "Backing up profile"
	objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrProfileKeys
	For Each strProfileKey In arrProfileKeys
		If Len (strProfileKey) > 8 Then  
			objReg.GetExpandedStringValue HKLM, strWindowsUninstallRegPath & "\" & strProfileKey, "ProfileImagePath", strPath
			strPath = objShell.ExpandEnvironmentStrings (strPath)
			If objFSO.FolderExists (strPath) = True Then
				Set objProfileFolder = objFSO.GetFolder (strPath)
				If StrComp (LCase (objProfileFolder.Name), LCase (strTargetUser)) = 0 Then
					
					' Copy registry to a new registry
					objShell.Run """" & strReg32 & """ copy ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & ".bak "" /s /f", 0, True
					objShell.Run """" & strReg64 & """ copy ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & ".bak"" /s /f", 0, True
					WScript.Echo "RENAMED KEY: " & strWindowsUninstallRegPath & "\" & strProfileKey & ".bak"
					
					' Change ProfileImagePath
					objShell.Run """" & strReg32 & """ add ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & ".bak"" /v ProfileImagePath /t REG_EXPAND_SZ /d """ & strPath &".bak"" /f"
					objShell.Run """" & strReg64 & """ add ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & ".bak"" /v ProfileImagePath /t REG_EXPAND_SZ /d """ & strPath &".bak"" /f"
					
					' Delete the old name
					objShell.Run """" & strReg32 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					objShell.Run """" & strReg64 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					
					' rename the user folder
					objShell.Run "%COMSPEC% /C RENAME """ & strPath & """ """ & objProfileFolder.Name & ".bak""", 0, True
					WScript.Echo "RENAMED FOLDER: " & strPath & ".bak"
				End If
			End If
		End If
	Next
End If

' RESTORE PROFILE
' ---------------
If blnRestoreUser = True Then
	WScript.Echo "Restoring profile"
	objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrProfileKeys
	For Each strProfileKey In arrProfileKeys
		If Len (strProfileKey) > 8 Then  
			objReg.GetExpandedStringValue HKLM, strWindowsUninstallRegPath & "\" & strProfileKey, "ProfileImagePath", strPath
			strPath = objShell.ExpandEnvironmentStrings (strPath)
			If objFSO.FolderExists (strPath) = True Then
				Set objProfileFolder = objFSO.GetFolder (strPath)
				If StrComp (LCase (objProfileFolder.Name), LCase (strTargetUser & ".bak")) = 0 Then
					' Copy registry to a new registry
					objShell.Run """" & strReg32 & """ copy ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ ""HKLM\" & strWindowsUninstallRegPath & "\" & Replace (strProfileKey, ".bak","") & """ /s /f", 0, True
					objShell.Run """" & strReg64 & """ copy ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ ""HKLM\" & strWindowsUninstallRegPath & "\" & Replace (strProfileKey, ".bak","") & """ /s /f", 0, True
					WScript.Echo "RENAMED KEY: " & strWindowsUninstallRegPath & "\" & Replace (strProfileKey, ".bak","")
					
					' Change ProfileImagePath
					objShell.Run """" & strReg32 & """ add ""HKLM\" & strWindowsUninstallRegPath & "\" & Replace (strProfileKey, ".bak","") & """ /v ProfileImagePath /t REG_EXPAND_SZ /d """ & Replace (strPath, ".bak", "") & """ /f"
					objShell.Run """" & strReg64 & """ add ""HKLM\" & strWindowsUninstallRegPath & "\" & Replace (strProfileKey, ".bak","") & """ /v ProfileImagePath /t REG_EXPAND_SZ /d """ & Replace (strPath, ".bak", "") & """ /f"
					
					' Delete the old name
					objShell.Run """" & strReg32 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					objShell.Run """" & strReg64 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					
					' rename the user folder
					objShell.Run "%COMSPEC% /C RENAME """ & strPath & """ """ & Replace (objProfileFolder.Name, ".bak", "") & """", 0, True
					WScript.Echo "RENAMED FOLDER: " & Replace (strPath, ".bak", "") & """"
				End If
			End If
		End If
	Next
End If


' DELETE ALL PROFILES
' -------------------
If blnRemoveAll = True And blnSingleUser = False Then
	WScript.Echo "Deleting all users"
	' Set up a list of folders that we don't want to delete
	' If we found the excluded user, add their home to the list.
	' ----------------------------------------------------------
	Dim dctExcludeUserHome
	Set dctExcludeUserHome = CreateObject ("scripting.dictionary")
	objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrProfileKeys
	WScript.Echo "There are " & UBound (arrProfileKeys) - 2 & " profiles on target computer."
	
	' Step through all profiles in the Profile List and remove them.
	For Each strProfileKey In arrProfileKeys	
		' If the Key ID is greater than 8 characters, this means that they are not a default system account.
		If Len (strProfileKey) > 8 Then
			objReg.GetExpandedStringValue HKLM, strWindowsUninstallRegPath & "\" & strProfileKey, "ProfileImagePath", strPath
			strPath = objShell.ExpandEnvironmentStrings (strPath)
			
			' This section we include a way to check the current profile against a dictionary of exclusive profiles.
			' ------------------------------------------------------------------------------------------------------
			' You need to get the username from the Path read from the registry and check if there's a fullstop in the name
			' If there's a fullstop in the name, this should be split up to remove the domain or the computer name part.
			strExcludeUser = ""
			If InStr (objFSO.GetFolder (strPath).Name, ".") Then
				' Get the character "." position using instr
				strExcludeUser = Left (objFSO.GetFolder (strPath).Name, InStr (objFSO.GetFolder (strPath).Name, ".") - 1)
			Else
				' There's no fullstop in the name, just check if it exists
				strExcludeUser = objFSO.GetFolder (strPath).Name
			End If
						
			' We need to check if this is the profile we want to exclude
			' Checking instring because, the folder may contain prefix or suffix if they are corrupted.
			'If (strExcludeUser <> "") And (InStr (LCase (strPath), LCase (strExcludeUser)) > 0) Then
			If (objExcludeUsers.Exists (LCase (strExcludeUser))) Then
				' We are excluding the user's folder. adding the user's folder into the exclude list
				WScript.Echo "Excluding: " & strExcludeUser
				dctExcludeUserHome.Add LCase (strPath), ""
			ElseIf (InStr (LCase (strPath), LCase (objShell.ExpandEnvironmentStrings ("%USERNAME%"))) > 0) Then
				' We also want to exclude the current logged on user
				WScript.Echo "Excluding: " & objShell.ExpandEnvironmentStrings ("%USERNAME%")
				dctExcludeUserHome.Add LCase (strPath), ""
			Else
				' Get user's name and start deleting the profile
				WScript.Echo "Deleting user: " + objFSO.GetFolder (strPath).Name		
				' Delete The User's Registry Entry in Profile List
				objShell.Run """" & strReg32 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
				objShell.Run """" & strReg64 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
				WScript.Echo "Deleted Registry Key: " & strWindowsUninstallRegPath & "\" & strProfileKey
				
				' Delete the User's Home folder
				If objFSO.FolderExists (strPath) = True Then
					intExitCode = objShell.Run ("%COMSPEC% /C RD /S /Q """ & strPath & """", 0, True)
					If intExitCode = 0 Then
						WScript.Echo "Deleted User Folder: " & strPath
					Else
						WScript.Echo "Cannot delete User Folder: " & strPath
					End If
				End If
			End If
		Else
			' We have found a default system account.
			' Read the home folder of that key and add it to the exclude list
			objReg.GetExpandedStringValue HKLM, strWindowsUninstallRegPath & "\" & strProfileKey, "ProfileImagePath", strPath
			dctExcludeUserHome.Add LCase (objShell.ExpandEnvironmentStrings (strPath)), ""
		End If
	Next
	
	' Remove Redundant/Obsolete User Home Folders
	' -------------------------------------------
	' We want to exclude system folders
	dctExcludeUserHome.Add LCase (objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%\Users\All Users")), ""
	dctExcludeUserHome.Add LCase (objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%\Users\Default")), ""
	dctExcludeUserHome.Add LCase (objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%\Users\Default User")), ""
	dctExcludeUserHome.Add LCase (objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%\Users\Public")), ""
	
	Dim objHomeFolders, objSubFolder, strSubFolderPath
	Set objHomeFolders = objFSO.GetFolder (objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%\Users"))
	
	For Each objSubFolder In objHomeFolders.SubFolders
		strSubFolderPath = objSubFolder.Path
		If Not (dctExcludeUserHome.Exists (LCase (strSubFolderPath))) Then
			intExitCode = objShell.Run ("%COMSPEC% /C RD /S /Q """ & strSubFolderPath & """", 0, True)
			If (intExitCode = 0) Then
				WScript.Echo "Successfully deleted orphan home folder: " + strSubFolderPath 
			Else
				WScript.Echo "Cannot delete orphan home folder: " + strSubFolderPath
			End If
		End If
	Next
	Set objHomeFolders = Nothing
End If

' DELETE PROFILES WITH LAST MODIFIED DATE OLDER THAN SPECIFIED DATE
' -----------------------------------------------------------------
If blnLastModified = True Then
	WScript.Echo "Deleting user profiles with last modified date older than: " & strLastModifiedDate
	objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrProfileKeys
	WScript.Echo "There are " & UBound (arrProfileKeys) + 1 & " profiles on target computer."
	For Each strProfileKey In arrProfileKeys
		If Len (strProfileKey) > 8 Then  
			objReg.GetExpandedStringValue HKLM, strWindowsUninstallRegPath & "\" & strProfileKey, "ProfileImagePath", strPath
			strPath = objShell.ExpandEnvironmentStrings (strPath)
			If objFSO.FolderExists (strPath) = True Then
				Set objProfileFolder = objFSO.GetFolder (strPath)
				Set objNTUserFile = objFSO.GetFile (strPath & "\NTUSER.DAT")
				If StrComp (LCase (objProfileFolder.Name), LCase (strExcludeUser)) <> 0 Then
					If strLastModifiedDate > CDate (Grep (objNTUserFile.DateLastModified, "^.*?(\d+/\d+/\d+).*?$", True)) Then
						objShell.Run """" & strReg32 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
						objShell.Run """" & strReg64 & """ delete ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
						WScript.Echo "Deleted: " & strWindowsUninstallRegPath & "\" & strProfileKey
						objShell.Run "%COMSPEC% /C RD /S /Q """ & strPath & """", 0, True
						WScript.Echo "deleted: " & strPath
					End If
				Else
					WScript.Echo "Excluding: " & strExcludeUser
				End If
			Else
					objShell.Run """" & strReg32 & """ DELETE ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					objShell.Run """" & strReg64 & """ DELETE ""HKLM\" & strWindowsUninstallRegPath & "\" & strProfileKey & """ /f", 0, True
					WScript.Echo "Deleted: " & strWindowsUninstallRegPath & "\" & strProfileKey
			End If
		End If
	Next
End If

If blnClearTemp = True Then
	WScript.Echo "Clearing temp folders"
	objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrProfileKeys
	Dim strTemp
	For Each strProfileKey In arrProfileKeys
		If Len (strProfileKey) > 8 Then  
			objReg.GetExpandedStringValue HKLM, strWindowsUninstallRegPath & "\" & strProfileKey, "ProfileImagePath", strPath
			strPath = objShell.ExpandEnvironmentStrings (strPath)
			If objFSO.FolderExists (strPath) = True Then
				Set objProfileFolder = objFSO.GetFolder (strPath)
				If blnSingleUser = True Then
				
					If StrComp (LCase (objProfileFolder.Name), LCase (strTargetUser)) = 0 Then
						strTemp = strPath & "\AppData\Local\Temp"
						If objFSO.FolderExists (strTemp) Then
							objShell.Run "%COMSPEC% /C RD /S /Q """ & strTemp & """", 0, True
							WScript.Echo "Deleted: " & strTemp
						End If
						
						strTemp = strPath & "\AppData\Local\Microsoft\Windows\Temporary Internet Files"
						If objFSO.FolderExists (strTemp) Then
							objShell.Run "%COMSPEC% /C RD /S /Q """ & strTemp & """", 0, True
							WScript.Echo "Deleted: " & strTemp
						End If
					End If
				
				Else
				
					If InStr (LCase (strPath), LCase (strExcludeUser)) > 0 Then
						WScript.Echo "Excluding user: " & strExcludeUser
					ElseIf InStr (LCase (strPath), LCase (objShell.ExpandEnvironmentStrings ("%USERNAME%"))) > 0 Then
						WScript.Echo "Excluding user: " & objShell.ExpandEnvironmentStrings ("%USERNAME%")
					Else
						strTemp = strPath & "\AppData\Local\Temp"
						If objFSO.FolderExists (strTemp) Then
							objShell.Run "%COMSPEC% /C RD /S /Q """ & strTemp & """", 0, True
							WScript.Echo "Deleted: " & strTemp
						End If
						
						strTemp = strPath & "\AppData\Local\Microsoft\Windows\Temporary Internet Files"
						If objFSO.FolderExists (strTemp) Then
							objShell.Run "%COMSPEC% /C RD /S /Q """ & strTemp & """", 0, True
							WScript.Echo "Deleted: " & strTemp
						End If
					End If
				End If
			End If
		End If
	Next
	
	' Delete the windows prefetch folder
	strTemp = objShell.ExpandEnvironmentStrings ("%WINDIR%\Prefetch")
	objShell.Run "%COMSPEC% /C RD /S /Q """ & strTemp & """", 0, True
	WScript.Echo "Deleted: " & strTemp
	
	' Delete the windows Temp folder
	strTemp = objShell.ExpandEnvironmentStrings ("%WINDIR%\Temp")
	objShell.Run "%COMSPEC% /C RD /S /Q """ & strTemp & """", 0, True
	WScript.Echo "Deleted: " & strTemp
End If

If blnDefrag = True Then
	WScript.Echo "Defraging computer..."
	objShell.Run "%COMSPEC% /C DEFRAG c: /U", 0, True
End If

If blnReboot = True Then
	WScript.Echo "Rebooting computer..."
	objShell.Run "SHUTDOWN -r -t 30 -f", 0, False
End If

WScript.Echo "Script completed!"

' -------
' THE END
' -------

' Custom Functions
' ----------------
Function IncludeFile (inFile)
    Dim objFSO, objFile, arrData
    arrData = Array ()
    Set objFSO = CreateObject ("Scripting.FileSystemObject")
    If objFSO.FileExists (inFile)  Then
    	Set objFile = objFSO.OpenTextFile (inFile)
    	arrData = objFile.ReadAll
    	objFile.Close
    	ExecuteGlobal arrData
    	Set objFile = nothing
    End If
End Function
