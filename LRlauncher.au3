#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Okke Garling

	Script Function:
	Launcher for LoadRunner automation

#ce ----------------------------------------------------------------------------
#AutoIt3Wrapper_Icon=targets-io.ico
#AutoIt3Wrapper_Change2CUI=y
#include <Date.au3>
#include <Array.au3>

;TODO: command line uitwerken
If $CmdLine[0] > 0 Then
	$sScenarioPath = $CmdLine[1]
Else
	ConsoleWriteError("Please provide at least one command line option argument: the path to the LoadRunner scenario file to be used." & @CRLF)
	Exit 1
EndIf
;~ $sScenarioPath = "C:\scripts\Jscen.lrs" ; debugging only!
$sScenarioPath = "C:\scripts\kort.lrs" ; debugging only!

Const $sIni = StringTrimRight(@ScriptName, 3) & "ini"
Const $sLRpath = IniRead($sIni, "LoadRunner", "LRpath", "C:\Program Files (x86)\HP\LoadRunner\bin\wlrun.exe")
Const $nTimeout = IniRead($sIni, "LoadRunner", "TimeoutDefault", "60")
Const $sGraphiteHost = IniRead($sIni, "Graphite", "GraphiteHost", "172.21.42.150")
Const $nGraphitePort = IniRead($sIni, "Graphite", "GraphitePort", "3000")
Const $sProductName = IniRead($sIni, "targets io", "ProductName", "LOADRUNNER")
Const $sDashboardName = IniRead($sIni, "targets io", "DashboardName", "LOAD")
Const $sProductRelease = IniRead($sIni, "targets io", "ProductRelease", "1.0")
Const $nRampupPeriod = IniRead($sIni, "targets io", "RampupPeriod", "10")
Const $nTimeZoneOffset = IniRead($sIni, "LR2Graphite", "TimeZoneOffset", "-1")

If Not LrsScriptPaths($sScenarioPath) Then
	ConsoleWriteError("Something went wrong while correcting script paths in scenario file." & $sScenarioPath & @CRLF)
	Exit 1
EndIf

; check if old LRA dir exists and if so, rename it to .old
If FileExists(@WorkingDir & "\LRR\LRA") Then
	If Not DirMove(@WorkingDir & "\LRR\LRA", @WorkingDir & "\LRR\LRA.old", 1) Then
		ConsoleWriteError("Old analysis directory detected and unable rename to " & @WorkingDir & "\LRR\LRA.old" & @CRLF & "Locked?" & @CRLF)
		MsgBox(16, "Error", "Old analysis directory detected and unable rename to " & @WorkingDir & "\LRR\LRA.old" & @CRLF & "Locked?", 5)
		Exit 1
	EndIf
EndIf

; send start event to targets io
$nRnd = Random(1,99999999,1)
; TODO: testrunid overnemen uit Jenkins of anders timestamp van maken _DateTimeSplit
SendJSONRunningTest("start", $sProductName, $sDashboardName, $nRnd, "", $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)

; check and run LoadRunner controller
If ProcessExists("wlrun.exe") Then
	ConsoleWriteError("LoadRunner controller process already running. Now closing." & @CRLF)
	If Not ProcessClose("wlrun.exe") Then
		ConsoleWriteError("LoadRunner controller process already running. Now closing." & @CRLF)
		Exit 1
	EndIf
EndIf
$iPid = Run($sLRpath & " -Run -InvokeAnalysis -TestPath " & $sScenarioPath & " -ResultName " & @WorkingDir & "\LRR")
If $iPid = 0 Or @error Then
	ConsoleWriteError("Something went wrong starting the scenario file with LoadRunner" & @CRLF)
	Exit 1
EndIf
; wait until timeout
; TODO: keep alives sturen!
If Not ProcessWaitClose($iPid, $nTimeout * 60) Then
	ConsoleWriteError("LoadRunner controller took too long to complete scenario. Timeout set at " & $nTimeout & " minutes. " & @CRLF & "Please check if LoadRunner is stalling or set higher timeout value. LoadRunner process, if still running, will be closed now..." & @CRLF)
	If Not ProcessClose($iPid) Then
		ConsoleWriteError("Unable to close LoadRunner controller process. Please do so manually." & @CRLF)
	EndIf
	Exit 1
EndIf

; send end event to targets io
; TODO: keepalive op termijn verwijderen
SendJSONRunningTest("keepalive", $sProductName, $sDashboardName, $nRnd, "", $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)
SendJSONRunningTest("end", $sProductName, $sDashboardName, $nRnd, "", $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)

; wait until completion of LoadRunner analysis
Sleep(2000) ; give analysis tool time to start
If Not ProcessWaitClose("AnalysisUI.exe", 300) Then
	ConsoleWriteError("LoadRunner analysis took too long to complete processing results. Timeout set at 5 minutes. " & @CRLF & "LoadRunner analysis process, if still running, will be closed now..." & @CRLF)
	If Not ProcessClose("AnalysisUI.exe") Then
		ConsoleWriteError("Unable to close LoadRunner analysis process. Please close manually." & @CRLF)
	EndIf
	Exit 1
EndIf

; check if results and analysis are properly executed and LoadRunner processes are closed
If Not FileExists(@WorkingDir & "\LRR\LRA\LRA.mdb") Then
	ConsoleWriteError("Error: LoadRunner Access database is not present. Remaining LoadRunner processes will be closed." & @CRLF)
	Exit 1
EndIf
If ProcessExists($iPid) Or ProcessExists("wlrun.exe") Or WinExists("HP LoadRunner Controller") Then
	ConsoleWriteError("LoadRunner controller process detected, now closing." & @CRLF)
	If Not ProcessClose($iPid) Then
		ProcessClose("wlrun.exe")
		WinClose("HP LoadRunner Controller")
	EndIf
EndIf
If ProcessExists("AnalysisUI.exe") Then
	ConsoleWriteError("LoadRunner analysis process detected, now closing." & @CRLF)
	If Not ProcessClose("AnalysisUI.exe") Then
		ConsoleWrite("LoadRunner analysis process could not be closed. Now exiting.")
		Exit 1  ; continuation not possible because of potential file lock for LR2Graphite
	EndIf
EndIf

; launching LR2Graphite
$ret = RunWait(@WorkingDir & '\LR2Graphite.exe "' & @WorkingDir & '\LRR\LRA\LRA.mdb" 172.21.42.150 2003 -1')
If $ret <> 0 Or @error Then
	ConsoleWriteError("Something went wrong during LR2Graphite execution. Now exiting." & @CRLF)
	Exit 1
EndIf

; return success
Exit 0

Func SendJSONRunningTest($sEvent, $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)
; Creating the object
$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
;~ $oHTTP.SetTimeouts(30000,60000,30000,30000)
$oHTTP.Open("POST", "http://" & $sGraphiteHost & ":" & $nGraphitePort & "/running-test/" & $sEvent, False)
$oHTTP.SetRequestHeader("Content-Type", "application/json")
$oHTTP.SetRequestHeader("Cache-Control", "no-cache")

$oHTTP.Send('{"testRunId": "' & $sTestrunId & '", ' & _
			'"dashboardName": "' & $sDashboardName & '", ' & _
			'"productName": "' & $sProductName & '", ' & _
			'"buildResultsUrl": "' & $sBuildResultsUrl & '", ' & _
			'"productRelease": "' & $sProductRelease & '", ' & _
			'"rampUpPeriod": "' & $nRampupPeriod & '"}')

; Download the body response if any, and get the server status response code.
$oReceived = $oHTTP.ResponseText
$oStatusCode = $oHTTP.Status

If $oStatusCode <> 200 then
	ConsoleWriteError("Response status code not 200 OK, but " & $oStatusCode & @CRLF)
	Return False
EndIf
EndFunc ; SendJSONRunningTest

Func LrsScriptPaths($sFile)
	$hFile = FileOpen($sFile, 0)
	If $hFile = -1 Then
		ConsoleWriteError("Unable to open scenario file." & @CRLF)
		Return False
	EndIf
	$hFileTmp = FileOpen($sFile & ".tmp", 2)
	If $hFileTmp = -1 Then
		ConsoleWriteError("Unable to open temporary scenario file." & @CRLF)
		Return False
	EndIf

	While 1
		$sLine = FileReadLine($hFile)
		If @error = -1 Then ExitLoop ; when EOF is reached
		;If StringRight($sLine, 3) = "usr" Then
		If StringRight($sLine, 3) = "usr" Then
			$aPath = StringSplit($sLine, "\")
			;_ArrayDisplay($aPath)
			;MsgBox(0,0,@WorkingDir & "\" & $aPath[$aPath[0] - 1] & "\" & $aPath[$aPath[0]])
			FileWriteLine($hFileTmp, @WorkingDir & "\" & $aPath[$aPath[0] - 1] & "\" & $aPath[$aPath[0]])
		Else
			FileWriteLine($hFileTmp, $sLine)
		EndIf
	WEnd
	FileClose($hFile)
	FileClose($hFileTmp)
	If Not FileMove($sFile, $sFile & ".old", 1) Then
		ConsoleWriteError("Unable to rename original scenario file to extension .old" & @CRLF)
		Return False
	EndIf
	If Not FileMove($sFile & ".tmp", $sFile, 1) Then
		ConsoleWriteError("Unable to rename temporary scenario file to new scenario file" & @CRLF)
		Return False
	EndIf
EndFunc ; LrsScriptPaths