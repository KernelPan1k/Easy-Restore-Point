#RequireAdmin
#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#AutoIt3Wrapper_Outfile=easyrestorepoint.exe
#AutoIt3Wrapper_Res_Description=Easy Restore Point By Kernel-Panik
#AutoIt3Wrapper_Res_Fileversion=1.0.0.3
#AutoIt3Wrapper_Res_ProductName=Easy Restore Point
#AutoIt3Wrapper_Res_ProductVersion=0.3
#AutoIt3Wrapper_Res_Comment=Easily create a restore point
#AutoIt3Wrapper_Res_CompanyName=kernel-panik
#AutoIt3Wrapper_Res_LegalCopyright=kernel-panik
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/rm /sf=1 /sv=1
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include-once
#include <GUIConstantsEx.au3>
#include <WindowsConstants.au3>
#include <MsgBoxConstants.au3>
#include <Date.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>

Global $sToolVersion = '0.3'
Global $sToolReportFile = "easyrestorepoint.txt"
Global $bPowerShellAvailable = Null
Global $bDevVersion = False
Global $__g_oSR_WMI = Null
Global $__g_oSR = Null

If $bDevVersion = True Then
	AutoItSetOption("MustDeclareVars", 1)
EndIf

If FileExists(@DesktopDir & "\" & $sToolReportFile) Then
	FileDelete(@DesktopDir & "\" & $sToolReportFile)
EndIf

Func LogMessage($message)
	Dim $sToolReportFile

	FileWrite(@DesktopDir & "\" & $sToolReportFile, $message & @CRLF)
EndFunc   ;==>LogMessage

If Not IsAdmin() Then
	MsgBox(16, "Fail", "Admin required")
	QuitTool()
EndIf

Func PowershellIsAvailable()
	Dim $bPowerShellAvailable

	If IsBool($bPowerShellAvailable) Then Return $bPowerShellAvailable

	Local $iPid = Run("powershell.exe", @TempDir, @SW_HIDE)

	If @error <> 0 Or Not $iPid Then
		$bPowerShellAvailable = False

		Return $bPowerShellAvailable
	EndIf

	ProcessClose($iPid)

	$bPowerShellAvailable = True

	Return $bPowerShellAvailable
EndFunc   ;==>PowershellIsAvailable

Func QuitTool($bOpen = False)
	Dim $sToolReportFile
	Dim $bDevVersion

	If $bOpen Then
		Run("notepad.exe " & @DesktopDir & "\" & $sToolReportFile)
	EndIf
	Exit
EndFunc   ;==>QuitTool

Func SR_EnumRestorePointsPowershell()
	Local $aRestorePoints[1][3], $iCounter = 0
	$aRestorePoints[0][0] = $iCounter
	Local $sOutput
	Local $aRP[0]
	Local $bStart = False

	If PowershellIsAvailable() = False Then
		Return $aRestorePoints
	EndIf

	Local $iPid = Run('Powershell.exe -Command "$date = @{Expression={$_.ConvertToDateTime($_.CreationTime)}}; Get-ComputerRestorePoint | Select-Object -Property SequenceNumber, Description, $date"', @SystemDir, @SW_HIDE, $STDOUT_CHILD)

	While 1
		$sOutput &= StdoutRead($iPid)
		If @error Then ExitLoop
	WEnd

	Local $aTmp = StringSplit($sOutput, @CRLF)

	For $i = 1 To $aTmp[0]
		Local $sRow = StringStripWS($aTmp[$i], $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)

		If $sRow = "" Then ContinueLoop

		If StringInStr($sRow, "-----") Then
			$bStart = True
		ElseIf $bStart = True Then
			_ArrayAdd($aRP, $sRow)
		EndIf
	Next

	For $i = 0 To UBound($aRP) - 1
		Local $sRow = $aRP[$i]
		Local $aRow = StringSplit($sRow, " ")
		Local $iNbr = $aRow[0]

		If $iNbr >= 4 Then
			Local $sTime = StringStripWS(_ArrayPop($aRow), $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
			Local $sDate = StringStripWS(_ArrayPop($aRow), $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)
			Local $iSequence = Number(StringStripWS($aRow[1], $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES))

			$sDate = StringReplace($sDate, '.', '/')
			$sDate = StringReplace($sDate, '-', '/')

			_ArrayDelete($aRow, 0)
			_ArrayDelete($aRow, 0)
			Local $sDescription = _ArrayToString($aRow, " ")
			$sDescription = StringStripWS($sDescription, $STR_STRIPLEADING + $STR_STRIPTRAILING + $STR_STRIPSPACES)

			$iCounter += 1
			ReDim $aRestorePoints[$iCounter + 1][3]
			$aRestorePoints[$iCounter][0] = $iSequence
			$aRestorePoints[$iCounter][1] = $sDescription
			$aRestorePoints[$iCounter][2] = $sDate & " " & $sTime
		EndIf
	Next

	$aRestorePoints[0][0] = $iCounter

	Return $aRestorePoints
EndFunc   ;==>SR_EnumRestorePointsPowershell

Func SR_WMIDateStringToDate($dtmDate)
	Return _
			(StringMid($dtmDate, 5, 2) & "/" & _
			StringMid($dtmDate, 7, 2) & "/" & _
			StringLeft($dtmDate, 4) & " " & _
			StringMid($dtmDate, 9, 2) & ":" & _
			StringMid($dtmDate, 11, 2) & ":" & _
			StringMid($dtmDate, 13, 2))
EndFunc   ;==>SR_WMIDateStringToDate

Func SR_EnumRestorePoints()
	Dim $__g_oSR_WMI

	Local $aRestorePoints[1][3], $iCounter = 0
	$aRestorePoints[0][0] = $iCounter

	If Not IsObj($__g_oSR_WMI) Then
		$__g_oSR_WMI = ObjGet("winmgmts:root/default")
	EndIf

	If Not IsObj($__g_oSR_WMI) Then
		Return SR_EnumRestorePointsPowershell()
	EndIf

	Local $RPSet = $__g_oSR_WMI.InstancesOf("SystemRestore")

	If Not IsObj($RPSet) Then
		Return SR_EnumRestorePointsPowershell()
	EndIf

	For $rP In $RPSet
		$iCounter += 1
		ReDim $aRestorePoints[$iCounter + 1][3]
		$aRestorePoints[$iCounter][0] = $rP.SequenceNumber
		$aRestorePoints[$iCounter][1] = $rP.Description
		$aRestorePoints[$iCounter][2] = SR_WMIDateStringToDate($rP.CreationTime)
	Next

	$aRestorePoints[0][0] = $iCounter

	Return $aRestorePoints
EndFunc   ;==>SR_EnumRestorePoints

Func SR_RemoveRestorePoint($rpSeqNumber)
	Local $aRet = DllCall('SrClient.dll', 'DWORD', 'SRRemoveRestorePoint', 'DWORD', $rpSeqNumber)

	If @error Then _
			Return SetError(1, 0, 0)

	If $aRet[0] = 0 Then _
			Return 1

	Return SetError(1, 0, 0)
EndFunc   ;==>SR_RemoveRestorePoint

Func convertDate($sDtmDate)
	Local $sY = StringLeft($sDtmDate, 4)
	Local $sM = StringMid($sDtmDate, 6, 2)
	Local $sD = StringMid($sDtmDate, 9, 2)
	Local $sT = StringRight($sDtmDate, 8)

	Return $sM & "/" & $sD & "/" & $sY & " " & $sT
EndFunc   ;==>convertDate

Func ClearDailyRestorePoint()
	Local Const $aRP = SR_EnumRestorePoints()

	If $aRP[0][0] = 0 Then
		Return Null
	EndIf

	Local Const $dTimeBefore = convertDate(_DateAdd('n', -1470, _NowCalc()))

	For $i = 1 To $aRP[0][0]
		Local $iDateCreated = $aRP[$i][2]

		If $iDateCreated > $dTimeBefore Then
			SR_RemoveRestorePoint($aRP[$i][0])

			Sleep(200)
		EndIf
	Next
EndFunc   ;==>ClearDailyRestorePoint

Func ShowCurrentRestorePoint()
	LogMessage(@CRLF & "- Display System Restore Point -" & @CRLF)

	Local Const $aRP = SR_EnumRestorePoints()

	If $aRP[0][0] = 0 Then
		LogMessage("  [X] No System Restore point found")
		Return
	EndIf

	For $i = 1 To $aRP[0][0]
		LogMessage("    ~ [I] RP named " & $aRP[$i][1] & " created at " & $aRP[$i][2] & " found")
	Next
EndFunc   ;==>ShowCurrentRestorePoint

Func CheckIsRestorePointExist()
	Local Const $aRP = SR_EnumRestorePoints()
	Local Const $iNbr = $aRP[0][0]

	If $iNbr = 0 Then
		Return False
	EndIf

	Return $aRP[$iNbr][1] = 'EasyRestorePoint'
EndFunc   ;==>CheckIsRestorePointExist

Func CreateSystemRestorePointWmi()
	#RequireAdmin
	RunWait(@ComSpec & ' /c ' & 'wmic.exe /Namespace:\\root\default Path SystemRestore Call CreateRestorePoint "EasyRestorePoint", 100, 7', "", @SW_HIDE)

	Sleep(2000)
EndFunc   ;==>CreateSystemRestorePointWmi

Func CreateSystemRestorePointPowershell()
	#RequireAdmin

	If PowershellIsAvailable() = True Then
		RunWait('Powershell.exe -Command Checkpoint-Computer -Description "EasyRestorePoint"', @ScriptDir, @SW_HIDE)
	EndIf

	Sleep(2000)
EndFunc   ;==>CreateSystemRestorePointPowershell

Func CreateSystemRestorePoint()
	CreateSystemRestorePointWmi()

	Local $bExist = CheckIsRestorePointExist()

	ProgressSet(50)

	If $bExist = False Then
		ClearDailyRestorePoint()
		CreateSystemRestorePointPowershell()
	EndIf
EndFunc   ;==>CreateSystemRestorePoint

Func SR_Enable($DriveL)
	Dim $__g_oSR

	If Not IsObj($__g_oSR) Then
		$__g_oSR = ObjGet("winmgmts:{impersonationLevel=impersonate}!root/default:SystemRestore")
	EndIf

	If Not IsObj($__g_oSR) Then
		Return 0
	EndIf

	If $__g_oSR.Enable($DriveL) = 0 Then
		Return 1
	EndIf

	Return 0
EndFunc   ;==>SR_Enable

Func EnableRestoration()
	Local $iSR_Enabled = SR_Enable(@HomeDrive & '\')

	If $iSR_Enabled = 0 Then
		If PowershellIsAvailable() = True Then
			RunWait("Powershell.exe -Command  Enable-ComputeRrestore -drive '" & @HomeDrive & "\' | Set-Content -Encoding utf8 ", @ScriptDir, @SW_HIDE)
		EndIf
	EndIf
EndFunc   ;==>EnableRestoration

Func CreateRestorePoint()
	ProgressOn("Easy Restore Point by kernel-panik", "Create a restore point", "0%")

	EnableRestoration()
	ProgressSet(25, "25%")
	CreateSystemRestorePoint()
	ProgressSet(75, "75%")
	ShowCurrentRestorePoint()
	ProgressSet(100, "100%")
EndFunc   ;==>CreateRestorePoint

Func GetHumanVersion()
	Switch @OSVersion
		Case "WIN_VISTA"
			Return "Windows Vista"
		Case "WIN_7"
			Return "Windows 7"
		Case "WIN_8"
			Return "Windows 8"
		Case "WIN_81"
			Return "Windows 8.1"
		Case "WIN_10"
			Return "Windows 10"
		Case Else
			Return "Unsupported OS"
	EndSwitch
EndFunc   ;==>GetHumanVersion

LogMessage("# Run at " & _Now())
LogMessage("# Easy Restore Point (kernel-panik) version " & $sToolVersion)
LogMessage("# Website https://kernel-panik.me/")
LogMessage("# Run by " & @UserName & " from " & @WorkingDir)
LogMessage("# Computer Name: " & @ComputerName)
LogMessage("# OS: " & GetHumanVersion() & " " & @OSArch & " (" & @OSBuild & ") " & @OSServicePack)

CreateRestorePoint()

QuitTool(True)

