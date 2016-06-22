#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Okke Garling

	Script Function:
	Launcher for LoadRunner automation

#ce ----------------------------------------------------------------------------
#AutoIt3Wrapper_Icon=targets-io.ico

If $CmdLine[0] > 0 Then
	$sScenarioPath = $CmdLine[1]
EndIf
$sScenarioPath = "C:\scripts\kort.lrs" ; debugging only!

Const $sIni = StringTrimRight(@ScriptName, 3) & "ini"
Const $sLRpath = IniRead($sIni, "LoadRunner", "LRpath", "C:\Program Files (x86)\HP\LoadRunner\bin\wlrun.exe")
Const $nTimeout = IniRead($sIni, "LoadRunner", "TimeoutDefault", "60")

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

; insert JSON
curl -X POST -H "Content-Type: application/json" -H "Cache-Control: no-cache"  -d "{   \"testRunId\":  \"$JOB_NAME-$BUILD_NUMBER\",
    \"dashboardName\":  \"LOAD\",
    \"productName\":  \"LOADRUNNER\",
    \"buildResultsUrl\": \"$BUILD_URL\",
    \"productRelease\": \"1.0\",
    \"rampUpPeriod\": \"10\"
}" "http://172.21.42.150:3000/running-test/start"


curl -X POST -H "Content-Type: application/json" -H "Cache-Control: no-cache"  -d "{   \"testRunId\":  \"$JOB_NAME-$BUILD_NUMBER\",
    \"dashboardName\":  \"LOAD\",
    \"productName\":  \"LOADRUNNER\",
    \"buildResultsUrl\": \"$BUILD_URL\",
    \"productRelease\": \"1.0\",
    \"rampUpPeriod\": \"10\"
}" "http://172.21.42.150:3000/running-test/keepalive"


curl -X POST -H "Content-Type: application/json" -H "Cache-Control: no-cache"  -d "{   \"testRunId\":  \"$JOB_NAME-$BUILD_NUMBER\",
    \"dashboardName\":  \"LOAD\",
    \"productName\":  \"LOADRUNNER\",
    \"buildResultsUrl\": \"$BUILD_URL\",
    \"productRelease\": \"1.0\",
    \"rampUpPeriod\": \"10\"
}" "http://172.21.42.150:3000/running-test/end"

;testrunid overnemen uit Jenkins of anders timestamp van maken
Func SendRunningTest($sEvent,





If FileExists(@WorkingDir & "\LRR\LRA") Then
	If Not DirMove(@WorkingDir & "\LRR\LRA", @WorkingDir & "\LRR\LRA.old", 1) Then
		ConsoleWriteError("Old analysis directory detected and unable rename to " & @WorkingDir & "\LRR\LRA.old" & @CRLF & "Locked?" & @CRLF)
		MsgBox(16, "Error", "Old analysis directory detected and unable rename to " & @WorkingDir & "\LRR\LRA.old" & @CRLF & "Locked?", 5)
		Exit False
	EndIf
EndIf

Exit

$iPid = Run($sLRpath & " -Run -InvokeAnalysis -TestPath " & $sScenarioPath & " -ResultName " & @WorkingDir & "\LRR")
If $iPid = 0 Or @error Then
	ConsoleWriteError("Something went wrong starting the scenario file with LoadRunner" & @CRLF)
	Exit False
EndIf
If Not ProcessWaitClose($iPid, $nTimeout * 60) Then
	ConsoleWriteError("LoadRunner controller took too long to complete scenario. Timeout set at " & $nTimeout & " minutes. " & @CRLF & "Please check if LoadRunner is stalling or set higher timeout value. LoadRunner process, if still running, will be closed now..." & @CRLF)
;~ 	MsgBox(16, "Error", "LoadRunner controller took too long to complete scenario. Timeout set at " & $nTimeout & " minutes. " & @CRLF & "Please check if LoadRunner is stalling or set higher timeout value. LoadRunner process, if still running, will be closed now...", 20)
	If Not ProcessClose($iPid) Then
		ConsoleWriteError("Unable to close LoadRunner controller process. Please do so manually." & @CRLF)
		MsgBox(16, "Error", "Unable to close LoadRunner controller process.", 5)
	EndIf
	Exit False
EndIf
;If @error Then
