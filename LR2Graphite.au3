#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Okke Garling

	Script Function:
	LoadRunner response time metrics export to Graphite

#ce ----------------------------------------------------------------------------
#AutoIt3Wrapper_Icon=targets-io.ico
#include "ADO.au3"
#include <Array.au3>

If $CmdLine[0] = 4 Then
	$sMDB_FileFullPath = $CmdLine[1]
	If StringRight($sMDB_FileFullPath, 3) <> "mdb" Then
		SearchPath($sMDB_FileFullPath) ; meant for integration with Jenkins plugin where currently it is not possible to know beforehand what the full path is due to randomly chosen 6 hexadecimals as subdirectory
		If ProcessExists("wlrun.exe") Then
			ConsoleWriteError("LoadRunner controller, wlrun.exe, will be closed forcefully" & @CRLF)
			TrayTip("LoadRunner", "controller process, wlrun.exe, will be closed forcefully", 3)
			If Not ProcessClose("wlrun.exe") Then
				ConsoleWriteError("Unable to close wlrun.exe" & @CRLF)
				Msgbox(16, "LoadRunner", "Unable to close wlrun.exe", 5)
				Exit 1
			EndIf
		EndIf
	EndIf
	$sGraphiteHost = $CmdLine[2]
	$nGraphitePort = $CmdLine[3]
	$nTimeZoneOffset = $CmdLine[4]
Else
	ConsoleWrite("Please specify following mandatory command line options: LR2Graphite <path to LR mdb> <Graphite host> <Graphite port> <timezone offset (hours)>" & @CRLF)
	$sMDB_FileFullPath = FileOpenDialog("Location of LoadRunner mdb database file?", "", "LoadRunner analysis (*.mdb)")
	If $sMDB_FileFullPath = "" Or @error Then
		MsgBox(16, "Error", "No valid location for LoadRunner mdb databbase file specified.")
		Exit 1
	EndIf
	$sGraphiteHost = InputBox("Graphite", "Graphite hostname or IP address?", "")
	If $sGraphiteHost = "" Or @error Then
		MsgBox(16, "Error", "No valid Graphite hostname or IP address specified.")
		Exit 1
	EndIf
	$nGraphitePort = InputBox("Graphite", "Graphite port number?", "2003")
	If $nGraphitePort = "" Or @error Then
		MsgBox(16, "Error", "No valid Graphite port number specified.")
		Exit 1
	EndIf
	$nTimeZoneOffset = InputBox("Timezone", "Timezone offset? (hours)", "0")
	If $nTimeZoneOffset = "" Or @error Then
		MsgBox(16, "Error", "No valid timezone offset specified.")
		Exit 1
	EndIf
EndIf

; TODO: ini file
$sConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $sMDB_FileFullPath
$sGraphiteRootNamespace = "LoadRunner"
Global Const $nGraphiteResolution = 10 ; aggregation resolution (seconds); determines timespan of "bucket"
Global Const $nPercentile = 99   ; percentage expressed in a number between 0 and 100
Global $nPayloadBytes = 0
Global $iSocket = Null
Global $sTransactionState = ""
Global $nStartTime, $nEndTime

TCPStartup()

; if Graphite host can be entered as hostname, so it is converted to IP address (even it is already a valid IP address)
$sGraphiteHost = TCPNameToIP($sGraphiteHost)
If $sGraphiteHost = "" Or @error Then
	TCPShutdown()
	ConsoleWriteError("Unable to resolve entered Graphite hostname into IP address. Now exiting." & @CRLF)
	MsgBox(16, "Graphite hostname", "Unable to resolve entered Graphite hostname into IP address. Now exiting.", 5)
	Sleep(5)
	Exit 1
EndIf

$oConnection = _ADO_Connection_Create()
_ADO_Connection_OpenConString($oConnection, $sConnectionString)
If @error Then
	SetError(@error, @extended, $ADO_RET_FAILURE)
	TCPShutdown()
EndIf

;determine available scripts
$aScriptTable = _ADO_Execute($oConnection, "SELECT * FROM Script", True)
If @error Then
	MsgBox(16, "Query script table", "Query failed." & @CRLF & "Valid LoadRunner database specified? (and not OUTPUT.MDB?)", 10)
	_ADO_Connection_Close($oConnection)
	$oConnection = Null
	TCPShutdown()
	Exit 1
EndIf
$nScripts = UBound($aScriptTable[2])

; determine start time of test
$aStartTime = _ADO_Execute($oConnection, "SELECT * FROM Result", True)
$nStartTime = ($aStartTime[2])[0][4] + ($nTimeZoneOffset * 3600) ; timezone offset correction

If $nScripts > 1 Then
	$sText = "scripts"
Else
	$sText = "script"
EndIf
ConsoleWrite("LR2Graphite found " & $nScripts & " " & $sText & " in database " & $sMDB_FileFullPath & ":" & @CRLF)
For $i = 0 to $nScripts - 1  ; loop through available scripts
	ConsoleWrite(($aScriptTable[2])[$i][1] & @CRLF)
Next

Local $aMeasurementsPerScript[$nScripts] ; dimension array to store metrics per script
For $i = 0 to $nScripts - 1  ; loop through available scripts
	$sScriptName = ($aScriptTable[2])[$i][1]
	$sQueryPassed = "SELECT Event_map.[Event Name], Event_meter.[End Time], Event_meter.Value FROM Event_map INNER JOIN Event_meter ON Event_map.[Event ID] = Event_meter.[Event ID] WHERE (((Event_meter.[Script ID])=" & $i & ") AND ((Event_meter.Status1)=1)) ORDER BY Event_map.[Event Name], Event_meter.[End Time];" ; query for passed transactions
	$sQueryFailed = "SELECT Event_map.[Event Name], Event_meter.[End Time], Event_meter.Value FROM Event_map INNER JOIN Event_meter ON Event_map.[Event ID] = Event_meter.[Event ID] WHERE (((Event_meter.[Script ID])=" & $i & ") AND ((Event_meter.Status1)=0)) ORDER BY Event_map.[Event Name], Event_meter.[End Time];" ; query for failed transactions
	$aAdoRetPassed = _ADO_Execute($oConnection, $sQueryPassed, True)
	$aAdoRetFailed = _ADO_Execute($oConnection, $sQueryFailed, True)
	If Not IsArray($aAdoRetPassed) And Not IsArray($aAdoRetFailed) Then ; if both passed and failed result sets aren't arrays then something is wrong with mdb
		ConsoleWriteError("Querying transactions for script " & $sScriptName & " resulted in something unexpected. Corrupt result set for this script?" & @CRLF)
		MsgBox(16, "Query measurements per script", "Querying transactions for script " & $sScriptName & " resulted in something unexpected." & @CRLF & "Corrupt result set for this script?", 10)
		ContinueLoop ; better luck for next script in line
	EndIf

	; from this point:	either a failed and/or passed transaction has occurred
	; processing passed transactions
	If Not IsArray($aAdoRetPassed) Then ; if no passed transactions occurred then failed transactions did: writing zeroes for metric tpps (passed)
		ConsoleWrite("Test failed: no passed transactions found for script " & $sScriptName & @CRLF & "Now writing zero values for metric tpps:" & @CRLF)
		; sending "0" values to Graphite:
		$nEndTime = $nStartTime + _ArrayMax($aAdoRetFailed[2], 1, 0, UBound($aAdoRetFailed[2]) - 1, 1)
		ZeroWritingGraphite("passed", $sScriptName, _ArrayUnique($aAdoRetFailed[2]))
	Else
		$sTransactionState = "passed" ; toggle "passed" mode
		$aMeasurementsPerScript[$i] = $aAdoRetPassed[2]
		$rc = ProcessScript($aMeasurementsPerScript[$i], $sScriptName)
		if @error Or Not $rc Then MsgBox(16, "Error", "Processing passed transactions for script " & $sScriptName & " failed.", 10)
	EndIf

	; processing failed transactions:
	If Not IsArray($aAdoRetFailed) Then ; if no failed transactions occurred then passed transactions did: writing zeroes for metric tfps (failed)
		ConsoleWrite(@CRLF & "Good: no failed transactions found for script " & $sScriptName & @CRLF & "Now writing zero values for metric tfps:" & @CRLF)
		; sending "0" values to Graphite:
		$nEndTime = $nStartTime + _ArrayMax($aAdoRetPassed[2], 1, 0, UBound($aAdoRetPassed[2]) - 1, 1)
		ZeroWritingGraphite("failed", $sScriptName, _ArrayUnique($aAdoRetPassed[2]))
	Else
		$sTransactionState = "failed" ; toggle "failed" mode
		$aMeasurementsPerScript[$i] = $aAdoRetFailed[2]
		$rc = ProcessScript($aMeasurementsPerScript[$i], $sScriptName)
		if @error Or Not $rc Then MsgBox(16, "Error", "Processing failed transactions for script " & $sScriptName & " failed.", 10)
	EndIf
Next

; Clean Up
_ADO_Connection_Close($oConnection)
$oConnection = Null

TrayTip("Ready!", "LoadRunner metrics exported to Graphite", 3)
Sleep(3000)

TCPShutdown()

Exit 0

Func ProcessScript ($aMeasurements, ByRef $sScript)
	TrayTip("", "Now processing " & $sTransactionState & " transactions for script " & $sScript, 1)
	ConsoleWrite(@CRLF & "Now processing " & $sTransactionState & " transactions for script " & $sScript & ":" & @CRLF & @CRLF)
	$aUniqueTransactions = _ArrayUnique($aMeasurements)
 	;_ArrayDisplay($aMeasurements)
	For $i = 1 to $aUniqueTransactions[0]
		$aTransactionSplitIndices = _ArrayFindAll($aMeasurements, $aUniqueTransactions[$i])
		Local $aTransactionsSplit[(UBound($aTransactionSplitIndices))][2] ; dimension array to store all measurements per transaction per script
		For $index = 0 to UBound($aTransactionSplitIndices) - 1
			$aTransactionsSplit[$index][0] = $aMeasurements[$aTransactionSplitIndices[$index]][1]
			$aTransactionsSplit[$index][1] = $aMeasurements[$aTransactionSplitIndices[$index]][2]
		Next
		$ret = ProcessTransaction($aTransactionsSplit, $aUniqueTransactions[$i], $sScript)
		if @error Or Not $ret Then
			MsgBox(16, "Error", "Processing transactions for transaction " & $aUniqueTransactions[$i] & " belonging to script " & $sScriptName & " failed.", 10)
			Return False
		EndIf
	Next

	Return True
EndFunc   ; ProcessScript

Func ProcessTransaction ($aTransactions, ByRef $sTransactionName, ByRef $sScript)
	ConsoleWrite("Processing transaction " & $sTransactionName & " from script " & $sScript & ": ")
	$nLastTransactionIndex = UBound($aTransactions) - 1
	;_ArrayDisplay($aTransactions)
	$nBuckets = Floor($aTransactions[$nLastTransactionIndex][0]) / $nGraphiteResolution + 1; amount of buckets needed
	Local $aBucketsHolder[$nBuckets]

	For $i = 0 to $nLastTransactionIndex
		$nBucketNr = Floor(($aTransactions[$i][0]) / $nGraphiteResolution)
		PlaceInBucket($aBucketsHolder[$nBucketNr], $aTransactions[$i][1])
	Next

;	_ArrayDisplay($aBucketsHolder[$nBuckets - 1])
	$ret = ProcessBuckets($aBucketsHolder, $sTransactionName, $sScript)
	if @error Or Not $ret Then
		MsgBox(16, "Error", "Processing buckets for transaction " & $sTransactionName & " belonging to script " & $sScript & " failed.", 10)
		Return False
	EndIf

	Return True
EndFunc		; ProcessTransaction

Func PlaceInBucket (ByRef $aBucket, ByRef $nValue)
		If IsArray($aBucket) Then
			_ArrayAdd($aBucket, $nValue)
		Else
			Local $aTemp[1]
			$aTemp[0] = $nValue
			$aBucket = $aTemp
		EndIf
EndFunc		; PlaceInBucket

Func ProcessBuckets ($aBuckets, ByRef $sTransactionName, ByRef $sScript)
	ConsoleWrite(UBound($aBuckets) & " buckets to process" & @CRLF)
	For $i = 0 to UBound($aBuckets) - 1
		If IsArray($aBuckets[$i]) Then ExportToGraphite($sGraphiteRootNamespace & "." & StringReplace($sScript, " ", "_") & "." & StringReplace($sTransactionName, " ", "_"), _ArrayMin($aBuckets[$i]), Average($aBuckets[$i]), _ArrayMax($aBuckets[$i]), Percentile($aBuckets[$i], $nPercentile), UBound($aBuckets[$i]) / $nGraphiteResolution, $nStartTime + (($i + 1) * $nGraphiteResolution))
	Next
	Return True
EndFunc

Func ExportToGraphite($sMetricPath, $nMin, $nAvg, $nMax, $nPerc, $nTps, $nEpoch)
	If $iSocket = Null Then
		$iSocket = TCPConnect($sGraphiteHost, $nGraphitePort)
		If $iSocket <= 0 Or @error Then
			ConsoleWriteError("Error occurred while connecting to Graphite host. Please check hostname/IP address and port number." & @CRLF)
			MsgBox(16, "Error", "Error connecting to Graphite host " & $sGraphiteHost & " on port " & $nGraphitePort, 10)
			Exit 1
		EndIf
	EndIf

	If $sTransactionState = "passed" Then
		$ret = TCPSend($iSocket, $sMetricPath & ".min " & $nMin & " " & $nEpoch & @LF)
		If @error Then ConsoleWriteError("TCPSend errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & ".min" & @CRLF)
		If Not @error Then $nPayloadBytes += $ret
		$ret = TCPSend($iSocket, $sMetricPath & ".avg " & $nAvg & " " & $nEpoch & @LF)
		If @error Then ConsoleWriteError("TCPSend errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & ".avg" & @CRLF)
		If Not @error Then $nPayloadBytes += $ret
		$ret = TCPSend($iSocket, $sMetricPath & ".max " & $nMax & " " & $nEpoch & @LF)
		If @error Then ConsoleWriteError("TCPSend errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & ".min" & @CRLF)
		If Not @error Then $nPayloadBytes += $ret
		$ret = TCPSend($iSocket, $sMetricPath & ".perc " & $nPerc & " " & $nEpoch & @LF)
		If @error Then ConsoleWriteError("TCPSend errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & ".perc" & @CRLF)
		If Not @error Then $nPayloadBytes += $ret
		$ret = TCPSend($iSocket, $sMetricPath & ".tpps " & $nTps & " " & $nEpoch & @LF)
		If @error Then ConsoleWriteError("TCPSend errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & ".tps" & @CRLF)
		If Not @error Then $nPayloadBytes += $ret
	ElseIf $sTransactionState = "failed" Then
		$ret = TCPSend($iSocket, $sMetricPath & ".tfps " & $nTps & " " & $nEpoch & @LF)
		If @error Then ConsoleWriteError("TCPSend errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & ".tps" & @CRLF)
		If Not @error Then $nPayloadBytes += $ret
	EndIf

	If $nPayloadBytes > 1048576 Then  ; after 1MB use new TCP connection
		$ret = TCPCloseSocket($iSocket)
		$iSocket = Null
		$nPayloadBytes = 0
		If @error Then
			ConsoleWriteError("TCPClose errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & @CRLF)
			$ret = TCPShutdown()
			If @error Then ConsoleWriteError("TCPShutdown errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & @CRLF)
			Exit 1
		EndIf
	EndIf
EndFunc ; ExportToGraphite

Func ZeroWritingGraphite($sTransState, $sScript, $aTransactionList)
	If $iSocket = Null Then
		$iSocket = TCPConnect($sGraphiteHost, $nGraphitePort)
		If $iSocket <= 0 Or @error Then
			ConsoleWriteError("Error occurred while connecting to Graphite host. Please check hostname/IP adrress and port number." & @CRLF)
			MsgBox(16, "Error", "Error connecting to Graphite host " & $sGraphiteHost & " on port " & $nGraphitePort, 10)
			Exit 1
		EndIf
	EndIf

	If $sTransState = "passed" Then $sMetricTps = "tpps"
	If $sTransState = "failed" Then $sMetricTps = "tfps"

	For $i = 1 to $aTransactionList[0]
		ConsoleWrite('Writing "0" (zero) values for not ' & $sTransState & ' transaction ' & $aTransactionList[$i] & ' in script ' & $sScript & @CRLF)
		For $nEpoch = $nStartTime to $nEndTime Step 10
			;ConsoleWrite($nEpoch & ": " & $i & @CRLF)
			$ret = TCPSend($iSocket, $sGraphiteRootNamespace & "." & StringReplace($sScript, " ", "_") & "." & StringReplace($aTransactionList[$i], " ", "_") & "." & $sMetricTps & " 0" & " " & $nEpoch & @LF)
			If @error Then ConsoleWriteError("TCPSend errorcode: " & @error & " while writing zeroes for transaction " & $aTransactionList[$i] & " of script " & $sScript & @CRLF)
			If Not @error Then $nPayloadBytes += $ret
			If $nPayloadBytes > 1048576 Then  ; after 1MB use new TCP connection
				$ret = TCPCloseSocket($iSocket)
				$iSocket = Null
				$nPayloadBytes = 0
				If @error Then
					ConsoleWriteError("TCPClose errorcode: " & @error & " socket: " & $iSocket & @CRLF)
					$ret = TCPShutdown()
					If @error Then ConsoleWriteError("TCPShutdown errorcode: " & @error & " socket: " & $iSocket & @CRLF)
					Exit 1
				EndIf
				$iSocket = TCPConnect($sGraphiteHost, $nGraphitePort)
				If $iSocket <= 0 Or @error Then
					ConsoleWriteError("Error occurred while connecting to Graphite host. Please check hostname/IP adrress and port number." & @CRLF)
					MsgBox(16, "Error", "Error connecting to Graphite host " & $sGraphiteHost & " on port " & $nGraphitePort, 10)
					Exit 1
				EndIf
			EndIf
		Next
	Next
EndFunc ; ZeroWritingGraphite

Func Average($aValues)
	$nSum = 0
	For $i = 0 to UBound($aValues) - 1
		$nSum += $aValues[$i]
	Next
	Return $nSum / UBound($aValues)
EndFunc ; Average

Func Percentile ($aNumbers, $nPercentile)  ;  Linear Interpolation Between Closest Ranks method (https://en.wikipedia.org/wiki/Percentile)
	_ArraySort($aNumbers)        ; sort array in ascending order
	If $nPercentile = 100 Then Return $aNumbers[UBound($aNumbers - 1)] ; in the unlikely event one requests the 100th percentile the last item/value in the array (after sort the maximum value) is returned
	$nTotalNumbers = UBound($aNumbers)
	$nIndex = (($nPercentile / 100) * ($nTotalNumbers - 1)) + 1

	$nElementIndexMin = Floor($nIndex)
	$nElementIndexMax = Ceiling($nIndex)
	$nFraction = $nIndex - $nElementIndexMin

	Return $aNumbers[$nElementIndexMin - 1] + ($nFraction * ($aNumbers[$nElementIndexMax - 1] - $aNumbers[$nElementIndexMin - 1]))
EndFunc ; Percentile

Func SearchPath (ByRef $sPath)
	Local $hSearch = FileFindFirstFile($sPath & "\*.")
	    If $hSearch = -1 Then
        MsgBox(16, "Error", "No subdirectory found.", 10)
        Return False
    EndIf

    While 1
        $sDir = FileFindNextFile($hSearch)
        ; If there is no more file matching the search.
        If @error Then ExitLoop
		If StringLen($sDir) = 6 Then $sPath = $sPath & "\" & $sDir & "\LRA\LRA.mdb"
		ExitLoop
	WEnd

	FileClose($hSearch)
	Return True
EndFunc ; SearchPath