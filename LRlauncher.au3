#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Okke Garling

	Script Function:
	Launcher for LoadRunner automation

#ce ----------------------------------------------------------------------------
#AutoIt3Wrapper_Icon=targets-io.ico
#include <Date.au3>
#include <Array.au3>

;TODO: command line uitwerken
If $CmdLine[0] > 0 Then
	$sScenarioPath = $CmdLine[1]
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


$hLrs = FileOpen($sScenarioPath)
If $hLrs = -1 Then
	ConsoleWriteError("Unable to open LoadRunner scenario file: " & $sScenarioPath & @CRLF)
	MsgBox(16, "Error", "Unable to open LoadRunner scenario file: " & $sScenarioPath, 5)
	Exit False
EndIf

; TODO: nieuwe scen file aanmaken en daarin script paths aanpassen
While 1
	$sResDirKeyValue = FileReadLine($hLrs)
	If StringLeft($sResDirKeyValue, 11) = "Result_file" Then
		$sResDir = StringTrimLeft($sResDirKeyValue, 12)
		ExitLoop
	EndIf
WEnd
FileClose($hLrs)

; check if old LRA dir exists and if so, rename it to .old
If FileExists(@WorkingDir & "\LRR\LRA") Then
	If Not DirMove(@WorkingDir & "\LRR\LRA", @WorkingDir & "\LRR\LRA.old", 1) Then
		ConsoleWriteError("Old analysis directory detected and unable rename to " & @WorkingDir & "\LRR\LRA.old" & @CRLF & "Locked?" & @CRLF)
		MsgBox(16, "Error", "Old analysis directory detected and unable rename to " & @WorkingDir & "\LRR\LRA.old" & @CRLF & "Locked?", 5)
		Exit False
	EndIf
EndIf

; send start event to targets io
$nRnd = Random(1,99999999,1)
;testrunid overnemen uit Jenkins of anders timestamp van maken _DateTimeSplit
SendJSONRunningTest("start", $sProductName, $sDashboardName, $nRnd, "", $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)

; run LoadRunner controller
$iPid = Run($sLRpath & " -Run -InvokeAnalysis -TestPath " & $sScenarioPath & " -ResultName " & @WorkingDir & "\LRR")
If $iPid = 0 Or @error Then
	ConsoleWriteError("Something went wrong starting the scenario file with LoadRunner" & @CRLF)
	Exit False
EndIf
; wait until timeout
; TODO: keep alives sturen!
If Not ProcessWaitClose($iPid, $nTimeout * 60) Then
	ConsoleWriteError("LoadRunner controller took too long to complete scenario. Timeout set at " & $nTimeout & " minutes. " & @CRLF & "Please check if LoadRunner is stalling or set higher timeout value. LoadRunner process, if still running, will be closed now..." & @CRLF)
;~ 	MsgBox(16, "Error", "LoadRunner controller took too long to complete scenario. Timeout set at " & $nTimeout & " minutes. " & @CRLF & "Please check if LoadRunner is stalling or set higher timeout value. LoadRunner process, if still running, will be closed now...", 20)
	If Not ProcessClose($iPid) Then
		ConsoleWriteError("Unable to close LoadRunner controller process. Please do so manually." & @CRLF)
		MsgBox(16, "Error", "Unable to close LoadRunner controller process.", 5)
	EndIf
	Exit False
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
		ConsoleWriteError("Unable to close LoadRunner analysis process. Please do so manually." & @CRLF)
		MsgBox(16, "Error", "Unable to close LoadRunner analysis process.", 5)
	EndIf
	Exit False
EndIf

; check if results and analysis are properly executed and LoadRunner processes are closed
If Not FileExists(@WorkingDir & "\LRR\LRA\LRA.mdb") Then
	ConsoleWriteError("Error: LoadRunner Access database is not present. Remaining LoadRunner processes will be closed." & @CRLF)
	Exit 1
EndIf
If ProcessExists($iPid) Or ProcessExists("wlrun.exe") Or WinExists("HP LoadRunner Controller") Then
	ConsoleWriteError("LoadRunner controller process detected, now closing." & @CRLF)
	; TODO: return values afvangen
	ProcessClose($iPid)
	ProcessClose("wlrun.exe")
	WinClose("HP LoadRunner Controller")
EndIf
If ProcessExists("AnalysisUI.exe") Then
	ConsoleWriteError("LoadRunner analysis process detected, now closing." & @CRLF)
	; TODO: return values afvangen
	ProcessClose("AnalysisUI.exe")
EndIf

; launching LR2Graphite
$ret = RunWait(@WorkingDir & '\LR2Graphite.exe "' & @WorkingDir & '\LRR\LRA\LRA.mdb" 172.21.42.150 2003 -1')
If $ret <> 0 Or @error Then
	ConsoleWriteError("Something went wrong during LR2Graphite execution." & @CRLF)
	Exit 1
EndIf

; return success
Exit 0

Func SendJSONRunningTest($sEvent, $sProductName, $sDashboardName, $sTestrunId, $sBuildResultsUrl, $sGraphiteHost, $nGraphitePort, $sProductRelease, $nRampupPeriod)

; Creating the object
$oHTTP = ObjCreate("winhttp.winhttprequest.5.1")
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
 MsgBox(4096, "Response code", $oStatusCode)
EndIf

;~ ; Saves the body response regardless of the Response code
;~  $file = FileOpen("Received.html", 2) ; The value of 2 overwrites the file if it already exists
;~  FileWrite($file, $oReceived)
;~  FileClose($file)
 EndFunc ; SendJSONRunningTest
