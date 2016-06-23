' Script Name: Win7DeployCommon.vbs
' Version: 0.60
' Desc: This script contains all the custom variables important for application
'       deployment
' Company: University of Sydney
' Author: Quoc Dat Nguyen
'==============================================================================
' CHANGE LOG
'==============================================================================
' Date			Ver		Description
' -----------|-------|---------------------------------------------------------
' 22/05/2013	0.05 |	Script created
' 17/06/2013 |  0.10 |  Added functions to read ini files
' 24/07/2013 |  0.15 |  Modified Script function disable file open security warning
'                    |  dialog in Windows 7
' 29/06/2014 |  0.20 |  Added firewall exception function
' 08/09/2015 |  0.30 |  Added function (DeleteProductCode) to delete Product Codes from uninstall registry hive
' 08/09/2015 |  0.40 |  Added functions (LoadConfig/FlushConfig) to load/write a configuration file to/from a hash/dictionary.
' 10/09/2015 |  0.50 |  Modified GetInstallLocation and GetUninstallString function so that we do not need to specify the registry bitness.
'                    |  The function will automatically search 32 and 64 bit.
' 24/09/2015 |  0.60 | Modified the Script function to apply lowriskfiletypes to HKLM and HKCU - It's only work with it applying to HKLM.
'==============================================================================

' Get the product install location from the registry
' --------------------------------------------------
Private Function GetInstallLocation (inProductID)
	Const HKLM = &H80000002
	Dim objReg, strUninstallPath, strSubkey, strValue, objFSO, blnFound, arrSubkeyList32, arrSubkeyList64
	Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")	
	arrSubkeyList32 = Array ()
	arrSubkeyList64 = Array ()
	strUninstallPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	arrSubkeyList32 = EnumKey (HKLM, strUninstallPath, 32)
	arrSubkeyList64 = EnumKey (HKLM, strUninstallPath, 64)
	blnFound = False
	
	
	If (Not (IsNull(arrSubkeyList32))) And  (Not (IsNull(arrSubkeyList64))) And (Err.Number = 0) Then
	
		' IF all OK, check the 32bit registry first
		For Each strSubkey In arrSubkeyList32
			If lcase(strSubkey) = lcase(inProductID) Then
				strValue = ReadStringValue (HKLM, strUninstallPath & "\" & strSubkey, "InstallLocation", 32)
				If objFSO.FolderExists (objFSO.GetAbsolutePathName (strValue)) Then
					GetInstallLocation = objFSO.GetAbsolutePathName (strValue)
					blnFound = True
					Exit For
				End If
			End If
		Next
		
		' Now check if its still not found, search under 64bit
		If blnFound = False Then 
			For Each strSubkey In arrSubkeyList64
				If lcase(strSubkey) = lcase(inProductID) Then
					strValue = ReadStringValue (HKLM, strUninstallPath & "\" & strSubkey, "InstallLocation", 64)
					If objFSO.FolderExists (objFSO.GetAbsolutePathName (strValue)) Then
						GetInstallLocation = objFSO.GetAbsolutePathName (strValue)
						blnFound = True
						Exit For
					End If
				End If
			Next
		End If
	End If
End Function

' Get a product Uninstall String
' ------------------------------
Function GetUninstallString (inProductID)
	Const HKLM = &H80000002
	Dim objReg, strUninstallPath, strSubkey, arrSubkeyList32, arrSubkeyList64, strValue, objFSO, blnFound
	Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")	
	arrSubkeyList32 = Array ()
	arrSubkeyList64 = Array ()
	strUninstallPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	arrSubkeyList32 = EnumKey (HKLM, strUninstallPath, 32)
	arrSubkeyList64 = EnumKey (HKLM, strUninstallPath, 64)
	blnFound = False
	
	If (Not (IsNull(arrSubkeyList32))) And (Not (IsNull(arrSubkeyList64))) And (Err.Number = 0) Then
		' Search the 32bit reg hive first
		For Each strSubkey In arrSubkeyList32
			If lcase(strSubkey) = lcase(inProductID) Then
				strValue = ReadStringValue (HKLM, strUninstallPath & "\" & strSubkey, "UninstallString", 32)
				GetUninstallString = Replace (strValue, """", "")
				blnFound = True
				Exit For
			End If
		Next
		
		' then if not found, search 64 bit
		If blnFound = False Then
			For Each strSubkey In arrSubkeyList64
				If lcase(strSubkey) = lcase(inProductID) Then
					strValue = ReadStringValue (HKLM, strUninstallPath & "\" & strSubkey, "UninstallString", 64)
					GetUninstallString = Replace (strValue, """", "")
					blnFound = True
					Exit For
				End If
			Next	
		End If
	End If
End Function

' Go through the GUID registry hive and delete a Product GUID
' -----------------------------------------------------------
Private Function DeleteProductCode (inProductID)	
	Dim strWindowsUninstallRegPath
	Dim arrAllSubKeys
	Dim strProductID
	Dim objReg
	Dim blnProductFound
		
	Const HKLM = &h80000002
	strWindowsUninstallRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	arrAllExistingProductIDs = Array ()
	blnProductFound = False
	
	If (isRunAs64bitMode And isOS64bit) Or  (Not (isRunAs64bitMode) And isOS64bit) Then
	' if either '64bit script host and 64 bit platform' or '32bit script host and 64 bit platform'
		Dim objCtx
		Dim objLocator
		Dim arrBitModes
		Dim intBitMode
		Dim objInParams
		Dim objOutParams
		Dim strSubKey
		Dim strRegCmd
		Dim intExitCode : intExitCode = 0

		arrBitModes = Array (32, 64)
		
		For Each intBitMode In arrBitModes
			Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
			objCtx.Add "__ProviderArchitecture", intBitMode
			objCtx.Add "__RequiredArchitecture", true
			Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
			Set objReg = objLocator.ConnectServer("","root\default","","",,,,objCtx).Get ("StdRegProv")
			
			Set objInParams = objReg.Methods_("EnumKey").InParameters
			objInParams.hDefKey = HKLM
			objInParams.sSubKeyName = strWindowsUninstallRegPath
		    Set objOutParams = objReg.ExecMethod_("EnumKey", objInParams, , objCtx)
		    
		    For Each strSubKey in objOutParams.snames
		    	' If strSubKey is the same as the inProductID then we delete the registry
		    	If StrComp (LCase (strSubKey), LCase (inProductID)) = 0 Then
		    		strRegCmd = GetRegCmd (intBitMode)
		    		intExitCode = objShell.Run ("""" & strRegCmd & """ DELETE ""HKLM\" & strWindowsUninstallRegPath & "\" & strSubKey & """ /f", 0, True)
		    		If intExitCode = 0 Then
		    			blnProductFound = True
		    		End If
		    	End If
			Next
		Next
	Else
		' 32bit script host on 32bit platform
		Set objReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrAllSubKeys
		For Each strSubKey In arrAllSubKeys
			' If strSubKey is the same as the inProductID then we delete the registry
	    	If StrComp (LCase (strSubKey), LCase (inProductID)) = 0 Then
	    		strRegCmd = GetRegCmd (32)
	    		intExitCode = objShell.Run ("""" & strRegCmd & """ DELETE ""HKLM\" & strWindowsUninstallRegPath & "\" & strSubKey & """ /f", 0, True)
	    		If intExitCode = 0 Then
	    			blnProductFound = True
	    		End If
	    	End If
		Next
	End If
	
	DeleteProductCode = blnProductFound
End Function

' Set the Windows Firewall exceptions
' -----------------------------------
Private Function AddFirewallException (inName, inDir, inExecutable, inAction, inProfile, inDescription, inEnabled)
	' Name = name of the rule
	' Dir = in | Out
	' Executeable = path to the file
	' Action = allow | block | bypass
	' Profile = public | private | domain | any
	' Description = description of the rule
	' enabled = yes | no
	
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "netsh advfirewall firewall add rule name=""" & inName & """ dir=" & inDir & " " _
				 & "program=""" & inExecutable & """ " & "action=" & inAction _
				 & " profile=" & inProfile & " description=""" & inDescription & """ enable=" & inEnabled, 0, True
End Function

Private Function DeleteFirewallException (inName, inDir)
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "netsh advfirewall firewall delete rule name=""" & inName & """ dir=" & inDir, 0, True
End Function

' Disable Win7 and Winxp Run security
' -----------------------------------
Private Function Script (inString, inError)
	Dim objShell
	Dim objEnv
	Dim strRegCmd
	Dim intExitCode
	
	intExitCode = 0 
	strRegCmd = GetRegCmd (64)
	
	Set objShell = CreateObject("WScript.Shell")
	Set objEnv = objShell.Environment("SYSTEM")

	If LCase (inString) = "start" Then
		objEnv("SEE_MASK_NOZONECHECKS") = 1
		
		' Per User
		intExitCode = objShell.Run ("""" & strRegCmd & """ ADD ""HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments"" /v ""SaveZoneInformation"" /t REG_DWORD /d 1 /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
		intExitCode = objShell.Run ("""" & strRegCmd & """ ADD ""HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Associations"" /v ""LowRiskFileTypes"" /t REG_SZ /d "".msi;.bat;.cmd;.exe;.reg;.vbs"" /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
		
		' Per Machine
		intExitCode = objShell.Run ("""" & strRegCmd & """ ADD ""HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments"" /v ""SaveZoneInformation"" /t REG_DWORD /d 1 /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
		intExitCode = objShell.Run ("""" & strRegCmd & """ ADD ""HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Associations"" /v ""LowRiskFileTypes"" /t REG_SZ /d "".msi;.bat;.cmd;.exe;.reg;.vbs"" /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
	Else
		objEnv("SEE_MASK_NOZONECHECKS") = 0
		
		' Per User
		intExitCode = objShell.Run ("""" & strRegCmd & """ DELETE ""HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments"" /v ""SaveZoneInformation"" /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
		intExitCode = objShell.Run ("""" & strRegCmd & """ DELETE ""HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Associations"" /v ""LowRiskFileTypes"" /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
		
		' Per Machine
		intExitCode = objShell.Run ("""" & strRegCmd & """ DELETE ""HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Attachments"" /v ""SaveZoneInformation"" /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
		intExitCode = objShell.Run ("""" & strRegCmd & """ DELETE ""HKLM\Software\Microsoft\Windows\CurrentVersion\Policies\Associations"" /v ""LowRiskFileTypes"" /f", 0, True)
		If intExitCode <> 0 Then WScript.Quit intExitCode
		
		' Stop the current script with the error code
		WScript.Quit (inError)
	End If
End Function

' Get a Backup Uninstall String
' ------------------------------
Function GetBackupUninstallString (inProductID, inBit)
	Const HKLM = &H80000002
	Dim objReg, strUninstallPath, strSubkey, arrSubkeyList, strValue, objFSO
	Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")	
	arrSubkeyList = Array ()
	strUninstallPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	arrSubkeyList = EnumKey (HKLM, strUninstallPath, inBit)
	
	If (Not (IsNull(arrSubkeyList))) And  (Err.Number = 0) Then
		For Each strSubkey In arrSubkeyList
			If lcase(strSubkey) = lcase(inProductID) Then
				strValue = ReadStringValue (HKLM, strUninstallPath & "\" & strSubkey, "BackupUninstallString", inBit)
				GetBackupUninstallString = Replace (strValue, """", "")
			End If
		Next
	End If
End Function

' Change a product uninstall string 
' ----------------------------------
Private Function ReplaceUninstallString (inProductID, inUninstallString, inBit)
	Dim strRegCmd
	Dim strSubkey
	Dim objShell
	Dim strUninstallString

	Const HKLM = &H80000002
	Set objShell = WScript.CreateObject("Wscript.Shell")
	strRegCmd = GetRegCmd (inBit)
	strUninstallPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	
	For Each strSubkey In EnumKey (HKLM, strUninstallPath, inBit)
		If LCase (strSubkey) = LCase (inProductID) Then
			' Read the current uninstall path
			' -------------------------------
			strUninstallString = ReadStringValue (HKLM, strUninstallPath & "\" & inProductID, "UninstallString", inBit)
			strUninstallString = Replace (strUninstallString, """", "\""")
			objShell.Run """" & strRegCmd & """ ADD ""HKLM\" & strUninstallPath & "\" & inProductID _
			           & """ /v ""BackupUninstallString"" /t REG_SZ /d """ & strUninstallString & """ /f", 0, True
			strUninstallString = Replace (inUninstallString, """", "\""")
			objShell.Run """" & strRegCmd & """ ADD ""HKLM\" & strUninstallPath & "\" & inProductID _
			           & """ /v ""UninstallString"" /t REG_SZ /d """ & strUninstallString & """ /f", 0, True
			objShell.Run """" & strRegCmd & """ ADD ""HKLM\" & strUninstallPath & "\" & inProductID _
			           & """ /v ""NoModify"" /t REG_DWORD /d 1 /f", 0, True      
		End If
	Next
End Function

' Check if a process is currently running
' ---------------------------------------
Private Function ProcessExists (inProcessExec)
	Dim objWMI
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Dim objProcess, colProcess, strProcessExec
	
	strProcessExec = "'" & inProcessExec & "'"
	Set colProcess = objWMI.ExecQuery ("Select * from Win32_Process Where Name = " & strProcessExec)
	If Not (IsNull (colProcess)) Then
		For Each objProcess in colProcess
			If LCase (objProcess.Caption) = LCase (inProcessExec) Then
				ProcessExists = True
				Exit Function
			End If
		Next
	End If
	ProcessExists = False
End Function

' Function to import a reg file
' -----------------------------
Private Function ImportRegistry (inFile, inBitness)
	Dim objShell
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")
	Set objShell = WScript.CreateObject("Wscript.Shell")
	
	Dim blnImportSucess
	Dim strREGCommand
	blnImportSucess = False
		
	If objFSO.FileExists (inFile) Then
		If inBitness = 32 And isRunAs64bitMode And isOS64bit Then
			' import 32 Registry, 64 bit OS, 64 bit mode, %WINDIR%\SYSWOW64\REG.EXE
			strREGCommand = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSWOW64\REGEDIT.EXE"
		ElseIf inBitness = 32 And Not (isRunAs64bitMode) And isOS64bit Then
			' import 32 Registry, 64 bit OS, 32 bit mode, %WINDIR%\SYSWOW64\REG.EXE
			strREGCommand = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSWOW64\REGEDIT.EXE"
		ElseIf inBitness = 64 And Not (isRunAs64bitMode) And isOS64bit Then
			' import 64 Registry, 64 bit OS, 32bit mode, %WINDIR%\SYSNATIVE\REG.EXE
			strREGCommand = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\REGEDIT.EXE"
		Else
			' The rest
			strREGCommand = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\REGEDIT.EXE"
		End If
		
		intExitCode = objShell.Run (strREGCommand & " /S """ & inFile & """", 0, true)
		If intExitCode = 0 Then blnImportSucess = True
	End If
	
	ImportRegistry = blnImportSucess
End Function

' Check if the product is installed on the system, windows xp or windows 7 wether its 64bit or 32bit
' --------------------------------------------------------------------------------------------------
Private Function ProductIDExists (inProductID)
	' There are three known conditions we need to cater for
	' Script executed by 64bit script host on 64bit platform
	' Script executed by 32bit script host on 64bit platform
	'       - With the two conditions above, we need to access 64bit and 32bit registry portals.
	'       - the one of the two ways to access 64bit registry from a 32bit runmode is through stdRegProv library.
	'		- The other way is to use SYSNATIVE and REG.EXE QUERY command.
	' script executed by 32bit script host on a 32bit platform
	
	Dim strWindowsUninstallRegPath
	Dim arrAllExistingProductIDs
	Dim strProductID
	Dim objReg
	Dim blnProductFound
		
	Const HKLM = &h80000002
	strWindowsUninstallRegPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
	arrAllExistingProductIDs = Array ()
	blnProductFound = False
	
	If (isRunAs64bitMode And isOS64bit) Or  (Not (isRunAs64bitMode) And isOS64bit) Then
	' if either '64bit script host and 64 bit platform' or '32bit script host and 64 bit platform'
		Dim objCtx
		Dim objLocator
		Dim arrBitModes
		Dim strBitMode
		Dim objInParams
		Dim objOutParams
		Dim strSubKey

		arrBitModes = Array (32, 64)
		
		For Each strBitMode In arrBitModes
			Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
			objCtx.Add "__ProviderArchitecture", strBitMode
			objCtx.Add "__RequiredArchitecture", true
			Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
			Set objReg = objLocator.ConnectServer("","root\default","","",,,,objCtx).Get ("StdRegProv")
			
			Set objInParams = objReg.Methods_("EnumKey").InParameters
			objInParams.hDefKey = HKLM
			objInParams.sSubKeyName = strWindowsUninstallRegPath
		    Set objOutParams = objReg.ExecMethod_("EnumKey", objInParams, , objCtx)
		    
		    For Each strSubKey in objOutParams.snames
				arrAllExistingProductIDs = AddArrayRecord (arrAllExistingProductIDs, strSubKey)	
			Next
		Next
	Else
		' 32bit script host on 32bit platform
		Set objReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		objReg.EnumKey HKLM, strWindowsUninstallRegPath, arrAllExistingProductIDs		
	End If
	
	For Each strProductID In arrAllExistingProductIDs
		If LCase (inProductID) = LCase (strProductID) Then
			blnProductFound = True
		End If
	Next
	ProductIDExists = blnProductFound
End Function

' Check if the OS is 64bit
' ------------------------
Private Function isOS64bit
	Dim strOSbitness
	strOSbitness = GetObject ("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2:Win32_Processor='cpu0'").AddressWidth
	If strOSbitness = 64 Then
		isOS64bit = True
	Else
		isOS64bit = False
	End If
End Function

' Make directory and subdirectories on the filesystem when you input c:\dir1\subdir1\subdir2\subdir3
' --------------------------------------------------------------------------------------------------
Private Function MakeDir (inDir)
	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (Not (inDir) = "") And (Not (objFSO.FolderExists (inDir))) Then
		Dim arrTrunks, strPath, n, Drive
		If Left (inDir, 2) = "\\" Then
			arrTrunks = Split (Right (inDir, Len (inDir) - 2), "\")
			strPath =  "\\" & arrTrunks (n)
		ElseIf (InStr (left (inDir, 2), ":") <> 0) Then
			arrTrunks = Split (inDir, "\")
			For Each Drive In objFSO.Drives
				If Not ((Drive.DriveType = 0) Or (Drive.DriveType = 4)) Then
					If (UCase (Drive.DriveLetter) = UCase (Left (arrTrunks (0), 1))) Then
						strPath = arrTrunks (0)
					End If
				End If
			Next
		Else
			Exit Function
		End If
		If Not (IsEmpty (strPath)) Then
			For n = 1 To UBound (arrTrunks)
				strPath = strPath & "\" & arrTrunks (n)
				If Not (objFSO.FolderExists (strPath)) Then
					objFSO.CreateFolder strPath
				End If
			Next
		End If
	End If
End Function

' Detect if the application is running under 32bit or 64bit Environment
' ---------------------------------------------------------------------
Private Function isRunAs64bitMode
	Dim strScriptTranslatorDir
	If isOS64bit Then
		strScriptTranslatorDir = WScript.CreateObject("Scripting.Filesystemobject").GetParentFolderName (WScript.FullName)
		If InStr (LCase (strScriptTranslatorDir), "syswow64") Then
			isRunAs64bitMode = False
		Else
			isRunAs64bitMode = True
		End If
	Else
		isRunAs64bitMode = False
	End If
End Function

' Get the OS Version in xx.xx
' ---------------------------
Private Function GetOSVersion
	Dim objWMIService, colOperatingSystems, objOperatingSystem, strOSVersion
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

	For Each objOperatingSystem in colOperatingSystems
    	strOSVersion = Grep (objOperatingSystem.Version, "^([0-9]+\.[0-9]+).*?$", True)
	Next
	If strOSVersion <> "" Then
		GetOSVersion = CDbl (strOSVersion)
	Else
		GetOSVersion = 0
	End If
End Function

' Return the pattern matched in a string
' --------------------------------------
Private Function Grep (inString, inPattern, inCase)
	Dim objRegExp
	Set objRegExp = New RegExp
	objRegExp.Pattern = inPattern
	objRegExp.IgnoreCase = inCase
	Grep = objRegExp.Replace (inString, "$1")
	Set objRegExp = Nothing
End Function

' Return true or false on matched string pattern
' ----------------------------------------------
Function MatchString (inString, MatchPattern, IgnoreCase)
	Dim objRegExp
	Set objRegExp = New RegExp
	objRegExp.Global = True
	objRegExp.IgnoreCase = IgnoreCase
	objRegExp.Pattern = MatchPattern
	MatchString = objRegExp.Test (inString)
End Function

' Add new element to an existing Array
' ------------------------------------
Private Function AddArrayRecord (arrTempArray(), strValue)
	ReDim Preserve  arrTempArray (UBound(arrTempArray) + 1)
	arrTempArray (UBound(arrTempArray)) = strValue
	AddArrayRecord = arrTempArray
End Function

' Kill target process running on the system
' -----------------------------------------
Private Function KillProcess (inProcessExec)
	Dim objWMI
	Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Dim objProcess, colProcess, strProcessExec
	
	strProcessExec = "'" & inProcessExec & "'"
	Set colProcess = objWMI.ExecQuery ("Select * from Win32_Process Where Name = " & strProcessExec)
	If Not (IsNull (colProcess)) Then
		For Each objProcess in colProcess
			objProcess.Terminate()
		Next
	End If
End Function

' Create Microsoft ActiveSetup Entry for the product
' --------------------------------------------------
Private Function CreateActiveSetup (inProductID, inProductName, inStubPath, inVersion, inBit)
	Const HKLM = "&H80000002"
	Dim arrProductCodes
	Dim strActiveSetupKey
	Dim objShell
	Dim strRegCmd
	Dim strVersion
	Dim intVersion
	
	intVersion = CInt (inVersion)
	strRegCmd = GetRegCmd (inBit)
	Set objShell = WScript.CreateObject("Wscript.Shell")
	
	strActiveSetupKey = "SOFTWARE\Microsoft\Active Setup\Installed Components"
	
	arrProductCodes = EnumKey (HKLM, strActiveSetupKey, inBit)
	If IsArray (arrProductCodes) Then
		' Check if there's already a key created, if there is, delete it and create a new key
		' -----------------------------------------------------------------------------------
		Dim strProductID
		For Each strProductID In arrProductCodes
			If lcase(strProductID) = lcase(inProductID) Then
				' Get the version number
				' ----------------------
				strVersion = ReadStringValue (HKLM, strActiveSetupKey & "\" & strProductID, "Version", inBit)
				If IsNumeric (strVersion) Then
					intVersion = CInt (strVersion) + 1
				End If
				objShell.Run """" & strRegCmd & """ DELETE ""HKLM\" & strActiveSetupKey & "\" & strProductID & """ /f", 0, True
			End If
		Next
	End If

	' Create the Active Setup Key
	' ---------------------------
	Dim strKeyPath
	strKeyPath = "HKLM\" & strActiveSetupKey & "\" & inProductID
	objShell.Run """" & strRegCmd & """ ADD """ & strKeyPath & """ /f", 0, True
	
	' Write the Active Setup Values
	' -----------------------------
	objShell.Run """" & strRegCmd & """ ADD """ & strKeyPath & """ /t REG_SZ /ve /d """ & inProductName & """ /f", 0, True
	objShell.Run """" & strRegCmd & """ ADD """ & strKeyPath & """ /t REG_SZ /v ""ComponentID"" /d """ & inProductName & """ /f", 0, True
	objShell.Run """" & strRegCmd & """ ADD """ & strKeyPath & """ /t REG_SZ /v ""StubPath"" /d """ & Replace (inStubPath, """", "\""") & """ /f", 0, True
	objShell.Run """" & strRegCmd & """ ADD """ & strKeyPath & """ /t REG_SZ /v ""Version"" /d """ & intVersion & """ /f", 0, True
End Function

' Delete Microsoft ActiveSetup Entry for the product
' --------------------------------------------------
Private Function DeleteActiveSetup (inProductID, inBit)
	Const HKLM = &H80000002
	Dim arrProductCodes
	Dim strActiveSetupKey
	Dim objShell
	Dim strRegCmd
	strRegCmd = GetRegCmd (inBit)
	Set objShell = WScript.CreateObject("Wscript.Shell")
	
	strActiveSetupKey = "SOFTWARE\Microsoft\Active Setup\Installed Components"
	
	arrProductCodes = EnumKey (HKLM, strActiveSetupKey, inBit)
	If IsArray (arrProductCodes) Then
		' Check if there's already a key created, if there is, delete it and create a new key
		' -----------------------------------------------------------------------------------
		Dim strProductID
		For Each strProductID In arrProductCodes
			If lcase(strProductID) = lcase(inProductID) Then
				objShell.Run """" & strRegCmd & """ DELETE ""HKLM\" & strActiveSetupKey & "\" & strProductID & """ /f", 0, True
			End If
		Next
	End If
End Function

' List SubKeys in Registry for Windows 7 and XP
' ---------------------------------------------
Private Function EnumKey (inRoot, inKeyPath, inBit)
	Dim objReg
	If (isRunAs64bitMode And isOS64bit) Or (Not (isRunAs64bitMode) And isOS64bit) Then
		' if either '64bit script host and 64 bit platform' or '32bit script host and 64 bit platform'
		Dim objCtx
		Dim objLocator
		Dim objInParams
		Dim objOutParams
		Dim arrSubKeys
		
		arrSubKeys = Array ()
		Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
		objCtx.Add "__ProviderArchitecture", inBit
		objCtx.Add "__RequiredArchitecture", true
		Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
		Set objReg = objLocator.ConnectServer("","root\default","","",,,,objCtx).Get ("StdRegProv")

		Set objInParams = objReg.Methods_("EnumKey").InParameters
		objInParams.hDefKey = inRoot
		objInParams.sSubKeyName = inKeyPath
		
	    Set objOutParams = objReg.ExecMethod_("EnumKey", objInParams, , objCtx)
	    Dim strSubKey
	   	For Each strSubKey in objOutParams.snames
			arrSubKeys = AddArrayRecord (arrSubKeys, strSubKey)	
		Next
	Else
		' 32bit script host on 32bit platform
		Set objReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		objReg.EnumKey inRoot, inKeyPath, arrSubKeys
	End If
	EnumKey = arrSubKeys
End Function

' Read String values from the Registry for Windows 7 and XP
' ---------------------------------------------------------
Private Function ReadStringValue (inRoot, inKeyPath, inValueName, inBit)
	Dim objReg
	If (isRunAs64bitMode And isOS64bit) Or (Not (isRunAs64bitMode) And isOS64bit) Then
		' if either '64bit script host and 64 bit platform' or '32bit script host and 64 bit platform'
		Dim objCtx
		Dim objLocator
		Dim objInParams
		Dim objOutParams
		Dim strValue
		
		Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
		objCtx.Add "__ProviderArchitecture", inBit
		objCtx.Add "__RequiredArchitecture", true
		Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
		Set objReg = objLocator.ConnectServer("","root\default","","",,,,objCtx).Get ("StdRegProv")

		Set objInParams = objReg.Methods_("GetStringValue").InParameters
		objInParams.hDefKey = inRoot
		objInParams.sSubKeyName = inKeyPath
		objInParams.Svaluename = inValueName
		
	    Set objOutParams = objReg.ExecMethod_("GetStringValue", objInParams, , objCtx)
	    ReadStringValue = objOutParams.SValue
	Else
		' 32bit script host on 32bit platform
		Set objReg = GetObject("WinMgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
		objReg.GetStringValue inRoot, inKeyPath, inValueName, strValue
		ReadStringValue = strValue
	End If
End Function

' Choose select the 32 bit and the 64bit of the reg.exe command
' -------------------------------------------------------------
Private Function GetRegCmd (inBit)
	' Returns the REG command and paths according to the Platform bits
	' This Function helps us easily manipulate registry
	' ----------------------------------------------------------------
	Dim objShell
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")
	Set objShell = WScript.CreateObject("Wscript.Shell")
	
	Dim strREGCommand
		
	If inBit = 32 And isRunAs64bitMode And isOS64bit Then
		' import 32 Registry, 64 bit OS, 64 bit mode, %WINDIR%\SYSWOW64\REG.EXE
		strREGCommand = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSWOW64\REG.EXE"
	ElseIf inBit = 64 And Not (isRunAs64bitMode) And isOS64bit Then
		' import 64 Registry, 64 bit OS, 32bit mode, %WINDIR%\SYSNATIVE\REG.EXE
		strREGCommand = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSNATIVE\REG.EXE"
	Else
		' The rest
		strREGCommand = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSTEM32\REG.EXE"
	End If

	GetRegCmd = strREGCommand
End Function

' Choose the correct System32 Folder
' ----------------------------------
Private Function GetSystem32 (inBit)
	Dim obdsjShell
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")
	Set objShell = WScript.CreateObject("Wscript.Shell")
	
	Dim strSystem32
		
	If inBit = 32 And isRunAs64bitMode And isOS64bit Then
		strSystem32 = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSWOW64"
	ElseIf inBit = 64 And Not (isRunAs64bitMode) And isOS64bit Then
		strSystem32 = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSNATIVE"
	Else
		strSystem32 = objShell.ExpandEnvironmentStrings ("%WINDIR%") & "\SYSTEM32"
	End If

	GetSystem32 = strSystem32
End Function

' Choose correct Program Files Folder
' -----------------------------------
Private Function getProgramFilesDir (inBit)
	Dim objShell
	Dim objFSO
	Set objFSO = WScript.CreateObject("Scripting.Filesystemobject")
	Set objShell = WScript.CreateObject("Wscript.Shell")
	
	Dim strProgramFiles
		
	If inBit = 32 And isRunAs64bitMode And isOS64bit Then
		strProgramFiles = objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%") & "\PROGRAM FILES (x86)"
	ElseIf inBit = 64 And Not (isRunAs64bitMode) And isOS64bit Then
		' import 64 Registry, 64 bit OS, 32bit mode, %WINDIR%\SYSNATIVE\REG.EXE
		strProgramFiles = objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%") & "\PROGRAM FILES"
	Else
		' The rest
		strProgramFiles = objShell.ExpandEnvironmentStrings ("%SYSTEMDRIVE%") & "\PROGRAM FILES"
	End If

	getProgramFilesDir = strProgramFiles
End Function

' Check if the computer still require a reboot
'---------------------------------------------
Function isPendingReboot
	Const HKEY_LOCAL_MACHINE = &H80000002
	Const REG_MULTI_SZ = 7

	Dim objReg, strKeyPath, strValueName, arrValues
	Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	strKeyPath = "SYSTEM\CurrentControlSet\Control\Session Manager"
	strValueName = "PendingFileRenameOperations"
	objReg.GetMultiStringValue HKEY_LOCAL_MACHINE, strKeyPath, strValueName, arrValues
	If Not IsNull(arrValues) Then
		isPendingReboot = True
	Else
		isPendingReboot = False
	End If
End Function

Function LoadConfig (inConfigFile, inSplitter, inComment)
	Dim objFSO, objTextFile, objConfigDict
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objConfigDict = CreateObject ("Scripting.Dictionary")

	Const FOR_READING = 1
	Const FOR_WRITING = 2
	Const FOR_APPENDING = 8
	Const AS_ASCII = 0
	Const AS_UNICODE = -1
	Const AS_DEFAULT = -2

	If objFSO.FileExists (inConfigFile) Then
		
		' Read the file and catching exceptions such as 'access denied'
		On Error Resume Next
			Set objTextFile = objFSO.OpenTextFile (inConfigFile, FOR_READING, True, AS_ASCII)

			If Err.Number <> 0 Then
				Err.Clear
				Exit Function
			End If
		On Error Goto 0
		
		Dim strReadLine : strReadLine = ""
		Dim arrKeyValue : arrKeyValue = Array ()
		Dim strKey : strKey = ""
		Dim strValue : strValue = ""
		
		' Read the config file
		Do Until (objTextFile.AtEndOfStream)
			' Get rid of all the spaces and characters at the front and at the end so we can catch comment character
		    strReadLine = RealTrim (objTextFile.ReadLine)
		    
		    ' Making sure that what we read isn't just an empty string
		    If (Len (strReadLine) > 0) Then
		    
		    	' Making sure that we do not read comments.
		    	If (Left (strReadLine, 1) <> inComment) Then
		    		
		    		' clearing the data from the previous iteration
		    		arrKeyValue = Array ()
		    		
			    	arrKeyValue = Split (strReadLine, inSplitter)
					
					' Only add or change value when we have a key value pair from the split
					If UBound (arrKeyValue) = 1 Then
					
						' Make sure that all Key names are uppercase
						strKey = UCase (RealTrim (arrKeyValue(0)))
						strValue = RealTrim (arrKeyValue(1))
						
				    	If (objConfigDict.Exists (strKey)) Then
				    		' change value
			    			objConfigDict.Item strKey = strValue
			    		Else
			    			' Add key value
			    			objConfigDict.Add strKey, strValue
				    	End If
				    	
				    End If
				    
				End If 
				
			End If
		Loop
		
		' close file and free the memory of the array
		strReadLine = ""
		Erase arrKeyValue
		objTextFile.Close
	End If
	
	' pass back the reference if the dictionary object
	Set LoadConfig = objConfigDict
	
	' Free all the memory of the objects created
	Set objTextFile = Nothing
	Set objFSO = Nothing
End Function

Function FlushConfig (inDictionary, inJoiner, inConfigFile)
	Dim objFSO, objTextFile 
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	Const FOR_READING = 1
	Const FOR_WRITING = 2
	Const FOR_APPENDING = 8
	Const AS_ASCII = 0
	Const AS_UNICODE = -1
	Const AS_DEFAULT = -2
	
	' Open text file for writing and catch any exception such as 'access denied'
	On Error Resume Next
		Set objTextFile = objFSO.OpenTextFile (inConfigFile, FOR_WRITING, True, AS_ASCII)
		If Err.Number <> 0 Then
			Err.Clear
			FlushConfig = False
			Exit Function
		End If
	On Error Goto 0
	
	Dim strKey
	Dim blnErrorOccurred : blnErrorOccurred = False
	For Each strKey In inDictionary.Keys
	
		' Catch write error exception and report back s false
		On Error Resume Next
			objTextFile.WriteLine strKey & inJoiner & inDictionary.Item (strKey)
			If Err.Number <> 0 Then
				Err.Clear
				blnErrorOccurred = true
			End If
		On Error Goto 0
		If blnErrorOccurred = True Then Exit Function
	Next
	
	' close file
	objTextFile.Close

	Set objTextFile = Nothing
	Set objFSO = Nothing
	FlushConfig = True
End Function

' Check if Text character is an ASCII whitespace
' ----------------------------------------------
Function isWhiteSpace(charIn)
	Dim intValue
	intValue = Asc(charIn)
	isWhiteSpace = intValue = 9 Or intValue = 10 Or intValue = 13 Or intValue = 32 Or intValue = 0 Or intValue = 11 Or intValue = 12
End Function

' Remove leading and ending whitespaces of a string
' -------------------------------------------------
Function RealTrim(strIn)
	If Len(strIn) = 0 Then
		RealTrim = strIn
		Exit Function
	End If
	Dim intPos, strOut
	intPos = 1
	While isWhiteSpace(Mid(strIn, intPos, 1))
		intPos = intPos + 1
		If intPos > Len(strIn) Then
			RealTrim = ""
			Exit Function
		End If
	Wend
	strOut = Mid(strIn, intPos)
	intPos = Len(strOut)
	While isWhiteSpace(Mid(strOut, intPos, 1))
		intPos = intPos - 1
	Wend
	RealTrim = Left(strOut, intPos)
End Function