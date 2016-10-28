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
#include <String.au3>

Global $iEventError = 0
Local $oMyError = ObjEvent("AutoIt.Error", "ErrFunc") ; install a custom error handler

If $CmdLine[0] = 5 Then ; Jenkins mode!
	$sScenarioPath = $CmdLine[1]
	$sDashboardName = $CmdLine[2]
	$sTestrunId = $CmdLine[3]
	$sBuildResultsUrl = $CmdLine[4]
	$sRunMode = $CmdLine[5]
ElseIf $CmdLine[0] = 2 Then ; standalone mode
	$sScenarioPath = $CmdLine[1]
	$sDashboardName = $CmdLine[2]
	$sTestrunId = "LOADRUNNER-" & StringReplace(_DateTimeFormat(_NowCalc(), 2), "/", "-") & "-" & Random(1, 99999, 1)
	$sBuildResultsUrl = ""
	$sRunMode = "standalone"
Else ; invalid amount of command line  parameters entered
	ConsoleWriteError("Please provide two or five command line parameter(s):" & @CRLF & @CRLF)
	ConsoleWriteError("LRlauncher.exe <path to scenario file> <DashboardName>" & @CRLF & @CRLF & "or Jenkins mode:" & @CRLF)
	ConsoleWriteError("LRlauncher.exe <path to scenario file> <DashboardName> <TestrunId> <BuildResultsUrl> <standalone|parallel>" & @CRLF & @CRLF)
	ConsoleWriteError("Please note: to be used LoadRunner scenario, the LoadRunner script directories and LRlauncher helper files should all be present in the same, current working directory or residing in a subdirectory." & @CRLF & @CRLF)
	Exit 1
EndIf

Global $sScenarioPathPrefix = DeterminePathPrefix($sScenarioPath)
$sIni = StringTrimRight(@ScriptName, 3) & "ini"
$sLRpath = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "LoadRunner", "LRpath", "C:\Program Files (x86)\HP\LoadRunner\bin\wlrun.exe")
$nTimeout = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "LoadRunner", "TimeoutDefault", "90")
$sHost = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "targets-io", "Host", "targets-io.klm.com")
$nPort = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "targets-io", "Port", "10003")
$sGraphiteHost = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "Graphite", "GraphiteHost", "172.21.42.178")
$nGraphitePort = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "Graphite", "GraphitePort", "2113")
$sProductName = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "targets-io", "ProductName", "LOADRUNNER")
$sProductRelease = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "targets-io", "ProductRelease", "1.0")
$nRampupPeriod = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "targets-io", "RampupPeriod", "10")
$nTimeZoneOffset = IniRead(@WorkingDir & $sScenarioPathPrefix & $sIni, "LR2Graphite", "TimeZoneOffset", "-1")

If Not LrsScriptPaths($sScenarioPath) Then
	ConsoleWriteError("Something went wrong while patching script paths in scenario file " & $sScenarioPath & @CRLF & "Now exiting." & @CRLF)
	Exit 1
Else
	ConsoleWrite("Scenario file " & $sScenarioPath & " patched successfully: script paths have been adapted to current working folder." & @CRLF)
EndIf

; check if old LRA dir exists and if so, rename it to .old
If FileExists(@WorkingDir & $sScenarioPathPrefix & "LRR\LRA") Then
	If Not DirMove(@WorkingDir & $sScenarioPathPrefix & "LRR\LRA", @WorkingDir & $sScenarioPathPrefix & "LRR\LRA.old", 1) Then
		ConsoleWriteError("Old analysis directory detected and unable rename to " & @WorkingDir & $sScenarioPathPrefix & "LRR\LRA.old" & @CRLF & "Locked?" & @CRLF)
		Exit 1
	EndIf
EndIf

; send start event to targets-io when run mode is not parallel
If $sRunMode <> "parallel" Then
	ConsoleWrite("Sending start event to targets-io: ")
	If Not SendJSONRunningTest("start", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sHost, $nPort, $sProductRelease, $nRampupPeriod) Then
		ConsoleWriteError("Sending end event unsuccessful: test will have status incompleted in targets-io." & @CRLF)
	Else
		ConsoleWrite("successful" & @CRLF)
	EndIf
EndIf

; check and run LoadRunner controller
If ProcessExists("wlrun.exe") Then
	ConsoleWriteError("LoadRunner controller process already running. Now closing." & @CRLF)
	If Not ProcessClose("wlrun.exe") Then
		ConsoleWriteError("Unable to close LoadRunner controller process. Now exiting." & @CRLF)
		Exit 1
	EndIf
EndIf
If WinExists("HP LoadRunner Controller") Then WinClose("HP LoadRunner Controller")
ConsoleWrite("Running of LoadRunner scenario started." & @CRLF)
$iPid = Run($sLRpath & ' -Run -InvokeAnalysis -TestPath "' & $sScenarioPath & '" -ResultName "' & @WorkingDir & $sScenarioPathPrefix & 'LRR"')
If $iPid = 0 Or @error Then
	ConsoleWriteError("Something went wrong starting the scenario file with LoadRunner. Now exiting." & @CRLF)
	Exit 1
EndIf

; wait until timeout
$sTestStart = _NowCalc()
If $sRunMode <> "parallel" Then ConsoleWrite("Sending keepalive events to targets-io during test: ")
While _DateDiff("s", $sTestStart, _NowCalc()) < $nTimeout * 60
	Sleep(15000) ; keepalive interval 15sec by default
	If $sRunMode <> "parallel" Then
		SendJSONRunningTest("keepalive", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sHost, $nPort, $sProductRelease, $nRampupPeriod)
		ConsoleWrite(".")
	EndIf
	If Not ProcessExists($iPid) Then ExitLoop
WEnd
ConsoleWrite(@CRLF)
If ProcessExists($iPid) Then
	ConsoleWriteError("LoadRunner controller took too long to complete scenario. Timeout set at " & $nTimeout & " minutes. " & @CRLF & "Please check if LoadRunner is stalling or set higher timeout value. LoadRunner process, if still running, will be closed now..." & @CRLF)
	If Not ProcessClose($iPid) Then
		ConsoleWriteError("Unable to close LoadRunner controller process. Please do so manually." & @CRLF)
	EndIf
	Exit 1
EndIf

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
If Not FileExists(@WorkingDir & $sScenarioPathPrefix & "LRR\LRA\LRA.mdb") Then
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
		Exit 1 ; continuation not possible because of potential file lock for LR2Graphite
	EndIf
EndIf

; launching LR2Graphite
If Not FileExists(@WorkingDir & $sScenarioPathPrefix & "LR2Graphite.exe") Then
	ConsoleWriteError("Unable to proceed: file LR2Graphite.exe not found in working directory " & @WorkingDir & @CRLF)
	Exit 1
EndIf
ConsoleWrite("Launching LR2Graphite." & @CRLF)
$ret = RunWait(@WorkingDir & $sScenarioPathPrefix & 'LR2Graphite.exe "' & @WorkingDir & $sScenarioPathPrefix & 'LRR\LRA\LRA.mdb" ' & $sGraphiteHost & ' ' & $nGraphitePort & ' ' & $nTimeZoneOffset)
If $ret <> 0 Or @error Then
	ConsoleWriteError("Something went wrong during LR2Graphite execution. Now exiting." & @CRLF)
	Exit 2 ; errorlevel 2 = LR2Graphite
Else
	ConsoleWrite("LoadRunner metrics successfully imported into Graphite." & @CRLF)
EndIf

If WinExists("HP LoadRunner Controller") Then WinClose("HP LoadRunner Controller")

; send end event to targets-io (at this point, otherwise if sooner targets-io is not able to calculate benchmark results)
If $sRunMode <> "parallel" Then
	; an extra keepalive to solve timing issue for rare occasions when previous keepalive failed
	SendJSONRunningTest("keepalive", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sHost, $nPort, $sProductRelease, $nRampupPeriod)
	ConsoleWrite("Sending end event to targets-io." & @CRLF)
	If Not SendJSONRunningTest("end", $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sHost, $nPort, $sProductRelease, $nRampupPeriod) Then
		ConsoleWriteError("Sending end event unsuccessful: test will have status incompleted in targets-io." & @CRLF)
	EndIf
EndIf

; assertions
If $sRunMode <> "parallel" Then
	If Not AssertionRequest($sProductName, $sDashboardName, $sTestrunId) Then
		ConsoleWriteError("Failed on assertions." & @CRLF)
		Exit 3 ; errorlevel 3 = assertions
	Else
		Exit 0 ; return success
	EndIf
EndIf

Func SendJSONRunningTest($sEvent, $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sHost, $nPort, $sProductRelease, $nRampupPeriod)
	; Creating the object
	$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
	$oHTTP.SetTimeouts(30 * 1000, 240 * 1000, 30 * 1000, 30 * 1000)
	$oHTTP.Open("POST", "http://" & $sHost & ":" & $nPort & "/running-test/" & $sEvent, False)
	$oHTTP.SetRequestHeader("Content-Type", "application/json")
	$oHTTP.SetRequestHeader("Cache-Control", "no-cache")

	$oHTTP.Send('{"testRunId": "' & $sTestrunId & '", ' & _
			'"dashboardName": "' & $sDashboardName & '", ' & _
			'"productName": "' & $sProductName & '", ' & _
			'"buildResultsUrl": "' & $sBuildResultsUrl & '", ' & _
			'"productRelease": "' & $sProductRelease & '", ' & _
			'"rampUpPeriod": "' & $nRampupPeriod & '"}')

	If Not $iEventError Then
		; download the body response if any, and get the server status response code.
		$oReceived = $oHTTP.ResponseText
		$oStatusCode = $oHTTP.Status
		If $oStatusCode <> 200 Then
			ConsoleWriteError("Targets-io event response status code not 200 OK, but " & $oStatusCode & @CRLF & "Response body: " & @CRLF & $oReceived & @CRLF)
			Return False
		EndIf
	Else
		$iEventError = 0 ; reset after displaying a COM Error occurred
	EndIf

	Return True
EndFunc   ;==>SendJSONRunningTest

Func DeterminePathPrefix($sPath)
	$aPath = StringSplit($sPath, "\")
	Return "\" & StringTrimRight($sPath, StringLen($aPath[$aPath[0]]))
EndFunc	;==>DeterminePathPrefix

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
			If Not FileExists(@WorkingDir & $sScenarioPathPrefix & $aPath[$aPath[0] - 1] & "\" & $aPath[$aPath[0]]) Then
				ConsoleWriteError("Error: script found in scenario file " & $sFile & " does not exist in location: " & @WorkingDir & $sScenarioPathPrefix & $aPath[$aPath[0] - 1] & "\" & $aPath[$aPath[0]] & @CRLF)
				FileClose($hFile)
				FileClose($hFileTmp)
				Return False
			EndIf
			FileWriteLine($hFileTmp, "Path=" & @WorkingDir & $sScenarioPathPrefix & $aPath[$aPath[0] - 1] & "\" & $aPath[$aPath[0]])
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
EndFunc   ;==>LrsScriptPaths

Func AssertionRequest($sProductName, $sDashboardName, $sTestrunId)
	; Creating the object
	$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
	$oHTTP.SetTimeouts(30 * 1000, 240 * 1000, 30 * 1000, 30 * 1000)
	$oHTTP.Open("GET", "http://" & $sHost & ":" & $nPort & "/testrun/" & StringUpper($sProductName) & "/" & StringUpper($sDashboardName) & "/" & StringUpper($sTestrunId), False)
	$oHTTP.SetRequestHeader("Content-Type", "application/json")
	$oHTTP.SetRequestHeader("Cache-Control", "no-cache")
	$oHTTP.Send()
	$sReceived = $oHTTP.ResponseText
	$nStatusCode = $oHTTP.Status

	If $nStatusCode <> 200 Then
		ConsoleWriteError("Targets-io response status code not 200 OK, but " & $nStatusCode & @CRLF & "Response body: " & @CRLF & $sReceived)
		SetError(1, 0, "Assertion request failed.")
	EndIf

	$aBenchmarkResultPreviousOK = _StringBetween($sReceived, '"benchmarkResultPreviousOK":', ',')
	$aBenchmarkResultFixedOK = _StringBetween($sReceived, '"benchmarkResultFixedOK":', ',')
	$aMeetsRequirement = _StringBetween($sReceived, '"meetsRequirement":', ',')

	If $aBenchmarkResultPreviousOK[0] = "false" Or $aBenchmarkResultFixedOK[0] = "false" Or $aMeetsRequirement[0] = "false" Then
		If $aMeetsRequirement[0] = "false" Then $sReturn = "Requirements not met: " & "http://" & $sGraphiteHost & ":" & $nGraphitePort & "/#!/requirements/" & StringUpper($sProductName) & "/" & StringUpper($sDashboardName) & "/" & StringUpper($sTestrunId) & "/failed/" & @CRLF
		If $aBenchmarkResultPreviousOK[0] = "false" Then $sReturn += "Benchmark with previous test result failed: " & "http://" & $sGraphiteHost & ":" & $nGraphitePort & "/#!/benchmark-previous-build/" & StringUpper($sProductName) & "/" & StringUpper($sDashboardName) & "/" & StringUpper($sTestrunId) & "/failed/" & @CRLF
		If $aBenchmarkResultFixedOK[0] = "false" Then $sReturn += "Benchmark with fixed baseline failed: " & "http://" & $sGraphiteHost & ":" & $nGraphitePort & "/#!/benchmark-fixed-baseline/" & StringUpper($sProductName) & "/" & StringUpper($sDashboardName) & "/" & StringUpper($sTestrunId) & "/failed/" & @CRLF
		ConsoleWrite($sReturn)
		Return False
	Else
		Return True
	EndIf
EndFunc   ;==>AssertionRequest

; This is a custom error handler
Func ErrFunc()
	ConsoleWriteError("Error trying to send a Targets-io event: " & $oMyError.description)
    $iEventError = 1 ; Use to check when a COM Error occurs
EndFunc   ;==>ErrFunc
