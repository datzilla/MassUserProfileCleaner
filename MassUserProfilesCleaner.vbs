' =======================================================
' Writen by Dat Nguyen
' This script executes another VBScript remotely.
' =======================================================
 Option Explicit

' ------------------------
' SCRIPT CONFIGURATIONS
' ------------------------
Dim strPCList, arrPCList, strdomain, strSourcePath, strComputer, strLocalDestinationPath
Dim strSuffix, strInstallCMD, intExitCode, strAppFolder, strDestinationUNCPath, strRunCMD

' IMPORTANT: INSERT YOUR USERNAME HERE
' ------------------------------------
strPCList = "\PC_LIST.txt"
strDomain = ".econ.usyd.edu.au"
strSuffix = "\c$"
strInstallCMD = "\run.cmd"
strAppFolder = "\UserProfilesCleaner"

Dim objFile, objFSO, objShell, objReg
Set objFSO = CreateObject("scripting.filesystemobject")
Set objShell = WScript.CreateObject("Wscript.Shell")

Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8
Const TristateTrue  = -1
Const TristateFalse = 0
Const TristateUseDefault = -2
Const HKEY_LOCAL_MACHINE = &H80000002

arrPCList = Array ()
strSourcePath = objFSO.GetParentFolderName (WScript.ScriptFullName)

' Open PC list file and load all lines into an array.
' ---------------------------------------------------
Set objFile = objFSO.OpenTextFile (strSourcePath & strPCList, ForReading, False, TristateFalse)
	While Not objFile.AtEndOfStream
		arrPCList = AddArrayRecord (arrPCList, objFile.ReadLine)
	Wend
objFile.Close
Set objFile = Nothing

For Each strComputer In arrPCList
	If Not (Left (strComputer, 1) = "'") Then
		If strComputer <> "" Then
			If (Ping (strComputer & strdomain)= True) Then
				On Error Resume Next
					' Check if we can connect to the admin share
					' ------------------------------------------
					strDestinationUNCPath = "\\" & strComputer & strSuffix & strAppFolder
					strLocalDestinationPath = "C:" & strAppFolder & strInstallCMD
					strRunCMD = """" & strSourcePath & "\psexec.exe"" \\" & strComputer & " -i -s -n 30 -w c:\WINDOWS\TEMP " & strLocalDestinationPath

					If objFSO.FolderExists ("\\" & strComputer & "\C$") Then
						objFSO.CopyFolder strSourcePath & strAppFolder, strDestinationUNCPath, True
						If objFSO.FileExists (strDestinationUNCPath & strInstallCMD) Then							
							intExitCode = objShell.Run (strRunCMD, 0, False)
							If intExitCode = 0 Then
								WScript.Echo strComputer & ": Profile cleaning started!"						
							Else
								WScript.Echo strComputer & ": Profiles cleaning failed!"
							End If
							
							If objFSO.FolderExists (strDestinationUNCPath) Then
								'intExitCode = objShell.Run ("%COMSPEC% /C RD /S /Q """ & strDestinationUNCPath & """", 0, True)
							End If 
						Else
							WScript.Echo strComputer & ": No script found..."
						End If
					Else
						WScript.Echo strComputer & ": Can't access the admin share"
					End If
				On Error Goto 0
			Else
				WScript.Echo strComputer & ": Offline"
			End If
		End If
	End If
Next

' Destroy all object and array
' ----------------------------
arrPCList = Array ()

WScript.Quit

' Check if a target return ping reply using WMI
' ---------------------------------------------
Function Ping(strHost)
    dim objPing, objRetStatus
    set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '" & strHost & "'")
    for each objRetStatus in objPing
        if IsNull(objRetStatus.StatusCode) or objRetStatus.StatusCode <> 0 then
    Ping = False
        else
            Ping = True
        end if
    next
End Function

' Add new record to an existing Array
' -----------------------------------
Function AddArrayRecord (arrTempArray(), strValue)
	ReDim Preserve  arrTempArray (UBound(arrTempArray) + 1)
	arrTempArray (UBound(arrTempArray)) = RealTrim (strValue)
	AddArrayRecord = arrTempArray
End Function

' Check if Text character is an ASCII whitespace
' ----------------------------------------------
Function isWhiteSpace(charIn)
	Dim intValue
	intValue = Asc(charIn)
	isWhiteSpace = intValue = 9 Or intValue = 10 Or intValue = 13 Or intValue = 32 Or intValue = 12 Or intValue = 0 Or intValue = 44
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
