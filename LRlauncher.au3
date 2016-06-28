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

Const $sIni = StringTrimRight(@ScriptName, 3) & "ini"
If $CmdLine[0] > 0 Then
	$sScenarioPath = $CmdLine[1]
	If $CmdLine[0] = 5 Then ; Jenkins mode!
		$sProductName = $CmdLine[2]
		$sDashboardName = $CmdLine[3]
		$sTestrunId = $CmdLine[4]
		$sBuildResultsUrl = $CmdLine[5]
	Else ; standalone mode
		$sProductName = IniRead($sIni, "targets io", "ProductName", "LOADRUNNER")
		$sDashboardName = IniRead($sIni, "targets io", "DashboardName", "LOAD")
		$sTestrunId = "LoadRunner-" & StringReplace(_DateTimeFormat(_NowCalc(), 2), "/", "-") & "-" & Random(1,99999,1)
		$sBuildResultsUrl = ""
	EndIf
Else
	ConsoleWriteError("Please provide at least one command line option argument: the path to the LoadRunner scenario file to be used." & @CRLF & @CRLF)
	ConsoleWriteError("LRlauncher.exe <path to scenario file>" & @CRLF & @CRLF & "or Jenkins mode:" & @CRLF)
	ConsoleWriteError("LRlauncher.exe <path to scenario file> <ProductName> <DashboardName> <TestrunId> <BuildResultsUrl>" & @CRLF & @CRLF)
	ConsoleWriteError("Please note: to be used script directories must be present in working directory from where LRlauncher is executed." & @CRLF)
	Exit 1
EndIf

$sLRpath = IniRead($sIni, "LoadRunner", "LRpath", "C:\Program Files (x86)\HP\LoadRunner\bin\wlrun.exe")
$nTimeout = IniRead($sIni, "LoadRunner", "TimeoutDefault", "90")
$sGraphiteHost = IniRead($sIni, "Graphite", "GraphiteHost", "172.21.42.150")
$nGraphitePort = IniRead($sIni, "Graphite", "GraphitePort", "3000")
$sProductRelease = IniRead($sIni, "targets io", "ProductRelease", "1.0")
$nRampupPeriod = IniRead($sIni, "targets io", "RampupPeriod", "10")
$nTimeZoneOffset = IniRead($sIni, "LR2Graphite", "TimeZoneOffset", "-1")

If Not LrsScriptPaths($sScenarioPath) Then
	ConsoleWriteError("Something went wrong while patching script paths in scenario file " & $sScenarioPath & @CRLF & "Now exiting." & @CRLF)
	Exit 1
Else
	ConsoleWrite("Scenario file " & $sScenarioPath & " patched successfully." & @CRLF)
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
ConsoleWrite("Sending start event to targets io: ")
; TODO: return value afvangen
SendJSONRunningTest("start", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)
ConsoleWrite("successfull" & @CRLF)

; check and run LoadRunner controller
If ProcessExists("wlrun.exe") Then
	ConsoleWriteError("LoadRunner controller process already running. Now closing." & @CRLF)
	If Not ProcessClose("wlrun.exe") Then
		ConsoleWriteError("Unable to close LoadRunner controller process. Now exiting." & @CRLF)
		Exit 1
	EndIf
EndIf
If WinExists("HP LoadRunner Controller") Then WinClose("HP LoadRunner Controller")
ConsoleWrite("Run LoadRunner scenario started." & @CRLF)
$iPid = Run($sLRpath & ' -Run -InvokeAnalysis -TestPath "' & $sScenarioPath & '" -ResultName "' & @WorkingDir & '\LRR"')
If $iPid = 0 Or @error Then
	ConsoleWriteError("Something went wrong starting the scenario file with LoadRunner" & @CRLF)
	Exit 1
EndIf

; wait until timeout
$sTestStart = _Now()
While _DateDiff("n", _Now, $sTestStart) < 1;$nTimeout
	Sleep(15000) ; keepalive interval
	SendJSONRunningTest("keepalive", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)
	ConsoleWrite("Sending keepalive event to targets io." & @CRLF)
	If Not ProcessExists($iPid) Then ExitLoop
WEnd
If ProcessExists($iPid) Then
	ConsoleWriteError("LoadRunner controller took too long to complete scenario. Timeout set at " & $nTimeout & " minutes. " & @CRLF & "Please check if LoadRunner is stalling or set higher timeout value. LoadRunner process, if still running, will be closed now..." & @CRLF)
	If Not ProcessClose($iPid) Then
		ConsoleWriteError("Unable to close LoadRunner controller process. Please do so manually." & @CRLF)
	EndIf
	Exit 1
EndIf
;If Not ProcessWaitClose($iPid, $nTimeout * 60) Then

; send end event to targets io
; TODO: keepalive op termijn verwijderen
;~ SendJSONRunningTest("keepalive", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)
;~ ConsoleWrite("Sending keepalive event to targets io: ")
ConsoleWrite("Sending end event to targets io: " & @CRLF)
If Not SendJSONRunningTest("end", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod) Then
	ConsoleWriteError("unsuccessful. Test will have status incompleted in targets-io." & @CRLF)
EndIf
; TODO: rv s afvangen en writen

ConsoleWrite("Analyzing results." & @CRLF)
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
Sleep(2000) ; give analysis tool time to start
If Not FileExists(@WorkingDir & "\LRR\LRA\LRA.mdb") Then
	ConsoleWriteError("Error: LoadRunner Access database is not present. Remaining LoadRunner processes will be closed." & @CRLF)
	Exit 1
Else
	ConsoleWrite("Analysis Access database is present." & @CRLF)
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
If Not FileExists(@WorkingDir & "\LR2Graphite.exe") Then
	ConsoleWriteError("Unable to proceed: file LR2Graphite.exe not found in working directory " & @WorkingDir & @CRLF)
	Exit 1
EndIf
ConsoleWrite("Launching LR2Graphite." & @CRLF)
$ret = RunWait(@WorkingDir & '\LR2Graphite.exe "' & @WorkingDir & '\LRR\LRA\LRA.mdb" 172.21.42.150 2003 -1')
If $ret <> 0 Or @error Then
	ConsoleWriteError("Something went wrong during LR2Graphite execution. Now exiting." & @CRLF)
	Exit 1
EndIf

If WinExists("HP LoadRunner Controller") Then WinClose("HP LoadRunner Controller")

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
	ConsoleWriteError("Response status code not 200 OK, but " & $oStatusCode & @CRLF & "Response body: " & @CRLF & $oReceived)
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
		If StringRight($sLine, 3) = "usr" Then
			$aPath = StringSplit($sLine, "\")
			FileWriteLine($hFileTmp, "Path=" & @WorkingDir & "\" & $aPath[$aPath[0] - 1] & "\" & $aPath[$aPath[0]])
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
	Return True
EndFunc ; LrsScriptPaths