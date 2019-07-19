#RequireAdmin

#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile=easyrestorepoint.exe
#AutoIt3Wrapper_Res_Description=Easy Restore Point By Kernel-Panik
#AutoIt3Wrapper_Res_Fileversion=1
#AutoIt3Wrapper_Res_ProductName=Easy Restore Point
#AutoIt3Wrapper_Res_ProductVersion=0.1
#AutoIt3Wrapper_Res_CompanyName=kernel-panik
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Res_LegalCopyright=kernel-panik
#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/rm /sf=1 /sv=1
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#include-once
#include "SystemRestore.au3""

Global $sToolVersion = '0.1'

Func QuitTool($bAutoDelete = False)
	If $bAutoDelete Then
		Run(@ComSpec & ' /c timeout 3 && del /F /Q "' & @ScriptFullPath & '"', @TempDir, @SW_HIDE)
		FileDelete(@ScriptFullPath)
	EndIf

	Exit
EndFunc   ;==>QuitKprm

Func _IsInternetConnected()
	Local $aReturn = DllCall('connect.dll', 'long', 'IsInternetConnected')

	If @error Then
		Return False
	EndIf

	Return $aReturn[0] = 0
EndFunc   ;==>_IsInternetConnected

Func CheckVersion()
	Dim $sToolVersion

	If _IsInternetConnected() = False Then Return

	Local $aCurrentVersion = StringSplit($sToolVersion, ".")
	Local $sVersion = _HTTP_Get("https://toolslib.net/api/softwares/951/version")
	Local $aVersion = StringSplit($sVersion, ".")
	Local $bNeedsUpdate = False

	For $i = 0 To UBound($aCurrentVersion) - 1
		If $aCurrentVersion[$i] < $aVersion[$i] Then
			$bNeedsUpdate = True
			ExitLoop
		EndIf
	Next

	If $bNeedsUpdate Then
		Local $sDownloadedFilePath = DownloadLatest()
		_SelfUpdate($sDownloadedFilePath, True, 5, False, False)

		If @error Then
			MsgBox($MB_SYSTEMMODAL, "Easy Restore Point - Updater", "The script must be a compiled exe to work correctly or the update file must exist.")
			QuitTool(True)
		EndIf
	EndIf
EndFunc   ;==>CheckVersionOfKpRm

Func DownloadLatest()
	ProgressOn("Easy Restore Point - Updater", "Downloading..", "0%")
	; Download from the following URL.
	Local $sURL = "https://download.toolslib.net/download/direct/951/latest"
	; Save the downloaded file to the temporary folder.
	Local $sFilePath = _WinAPI_GetTempFileName(@TempDir)
	; Remote file size
	Local $iRemoteFileSize = InetGetSize($sURL)
	; Download the file in the background with the selected option of 'force a reload from the remote site.'
	Local $hDownload = InetGet($sURL, $sFilePath, $INET_FORCERELOAD, $INET_DOWNLOADBACKGROUND)

	; Wait for the download to complete by monitoring when the 2nd index value of InetGetInfo returns True.
	Do
		Sleep(250)

		; Get bytes received
		Local $iBytesReceived = InetGetInfo($hDownload, $INET_LOCALCACHE)
		; Calculate percentage
		Local $iPercentage = Int($iBytesReceived / $iRemoteFileSize * 100)
		; Set progress bar
		ProgressSet($iPercentage, $iPercentage & "%")
	Until InetGetInfo($hDownload, $INET_DOWNLOADCOMPLETE)

	; Retrieve the number of total bytes received and the filesize.
	Local $iBytesSize = InetGetInfo($hDownload, $INET_DOWNLOADREAD)
	Local $iFileSize = FileGetSize($sFilePath)

	; Close the handle returned by InetGet.
	InetClose($hDownload)

	; We are done.
	ProgressOff()

	; Display details about the total number of bytes read and the filesize.
;~     MsgBox($MB_SYSTEMMODAL, "", "The total download size: " & $iBytesSize & @CRLF & _
;~             "The total filesize: " & $iFileSize)

	Return $sFilePath

EndFunc   ;==>DownloadLatest

Func convertDate($sDtmDate)
	Local $sY = StringLeft($sDtmDate, 4)
	Local $sM = StringMid($sDtmDate, 6, 2)
	Local $sD = StringMid($sDtmDate, 9, 2)
	Local $sT = StringRight($sDtmDate, 8)

	Return $sM & "/" & $sD & "/" & $sY & " " & $sT
EndFunc   ;==>convertDate

Func ClearDayRestorePoint()
	Local Const $aRP = _SR_EnumRestorePoints()

	If $aRP[0][0] = 0 Then
		Return Null
	EndIf

	Local Const $dTimeBefore = convertDate(_DateAdd('n', -1470, _NowCalc()))

	For $i = 1 To $aRP[0][0]
		Local $iDateCreated = $aRP[$i][2]

		If $iDateCreated > $dTimeBefore Then
			_SR_RemoveRestorePoint($aRP[$i][0])
		EndIf
	Next
EndFunc   ;==>ClearDayRestorePoint

Func ShowCurrentRestorePoint()
	Local Const $aRP = _SR_EnumRestorePoints()

	If $aRP[0][0] = 0 Then
		LogMessage("No System Restore point found")
		Return
	EndIf

	For $i = 1 To $aRP[0][0]
		LogMessage("RP named " & $aRP[$i][1] & " created at " & $aRP[$i][2] & " found")
	Next
EndFunc   ;==>ShowCurrentRestorePoint

Func CreateSystemRestorePoint()
	#RequireAdmin
	RunWait(@ComSpec & ' /c ' & 'wmic.exe /Namespace:\\root\default Path SystemRestore Call CreateRestorePoint "Easy Restore Point", 100, 7', "", @SW_HIDE)

	Return @error
EndFunc   ;==>CreateSystemRestorePoint


Func CreateRestorePoint()
	Local $iSR_Enabled = _SR_Enable()

	If $iSR_Enabled = 0 Then
		Sleep(5000)
		$iSR_Enabled = _SR_Enable()
	EndIf

	ClearDayRestorePoint()
	Sleep(1500)
	CreateSystemRestorePoint()
	Sleep(1500)
	ShowCurrentRestorePoint()
EndFunc   ;==>CreateRestorePoint
