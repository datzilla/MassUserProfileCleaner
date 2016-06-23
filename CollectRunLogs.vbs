' =======================================================
' Writen by Dat Nguyen
' This script collects the logs from the remote computers
' =======================================================
 Option Explicit

' ------------------------
' SCRIPT CONFIGURATIONS
' ------------------------
Dim strPCList, arrPCList, strdomain, strSourcePath, strComputer, strLocalDestinationPath
Dim strDestinationSuffix, strLogFileExt, intExitCode, strLogFolder, strDestinationUNCPath, strRunCMD

Dim objFile, objFSO, objShell, objReg
Set objFSO = CreateObject("scripting.filesystemobject")
Set objShell = WScript.CreateObject("Wscript.Shell")

' IMPORTANT: INSERT YOUR USERNAME HERE
' ------------------------------------
strPCList = "\PC_LIST.txt"
strDomain = ".econ.usyd.edu.au"
strDestinationSuffix = "\c$"
strLogFileExt = ".runlog"
strLogFolder = objFSO.GetParentFolderName (Wscript.ScriptFullName) & "\Logs"

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
					strDestinationUNCPath = "\\" & strComputer & strDestinationSuffix

					If objFSO.FolderExists (strDestinationUNCPath) Then
						
						If objFSO.FileExists (strDestinationUNCPath & "\" & strComputer & strLogFileExt) Then							
							' Copy the log files to local location
							objFSO.CopyFile strDestinationUNCPath & "\" & strComputer & strLogFileExt, strLogFolder & "\" & strComputer & strLogFileExt, True
							WScript.Echo strComputer & ": log copied"
						Else
							WScript.Echo strComputer & ": No log found..."
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