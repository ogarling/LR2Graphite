#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Okke Garling

	Script Function:
	LoadRunner response time metrics export to Graphite

#ce ----------------------------------------------------------------------------

#AutoIt3Wrapper_icon=targets-io.ico
#include "ADO.au3"
#include <Array.au3>
;~ #include <MsgBoxConstants.au3>
;~ #include <AutoItConstants.au3>
;~ #include <FileConstants.au3>


If $CmdLine[0] = 3 Then
	$sMDB_FileFullPath = $CmdLine[1]
	$sGraphiteHost = $CmdLine[2]
	$nGraphitePort = $CmdLine[3]
Else
	ConsoleWrite("Please specify following mandatory command line options: LR2Graphite <path to LR mdb> <Graphite host> <Graphite port>" & @CRLF)
;~ 	$sMDB_FileFullPath = "lt.mdb"
	$sMDB_FileFullPath = FileOpenDialog("Location of LoadRunner mdb database file?", "", "LoadRunner analysis (*.mdb)")
	If $sMDB_FileFullPath = "" Or @error Then
		MsgBox(16, "Error", "No valid location for LoadRunner mdb databbase file specified.")
		Exit
	EndIf
	$sGraphiteHost = InputBox("Graphite", "Graphite hostname or IP address?", "172.21.42.150")
	If $sGraphiteHost = "" Or @error Then
		MsgBox(16, "Error", "No valid Graphite hostname or IP address specified.")
		Exit
	EndIf
	$nGraphitePort = InputBox("Graphite", "Graphite port number?", "2003")
	If $nGraphitePort = "" Or @error Then
		MsgBox(16, "Error", "No valid Graphite port number specified.")
		Exit
	EndIf
EndIf

$sConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $sMDB_FileFullPath
;~ $sGraphiteHost = "172.21.42.150"
;~ $nGraphitePort = 2003
$sGraphiteRootNamespace = "LoadRunner"
Global Const $nGraphiteResolution = 10 ; aggregation resolution (seconds); determines timespan of "bucket"
Global Const $nPercentile = 99   ; percentage expressed in a number between 0 and 100
Global $nPayloadBytes = 0
Global $iSocket = Null

TCPStartup()

$oConnection = _ADO_Connection_Create()
_ADO_Connection_OpenConString($oConnection, $sConnectionString)
If @error Then SetError(@error, @extended, $ADO_RET_FAILURE)

;determine available scripts
$aScriptTable = _ADO_Execute($oConnection, "SELECT * FROM Script", True)
	If @error Then MsgBox(16, "Query script table", "Query failed. Valid LoadRunner database specified? (and not OUTPUT.MDB?)")
$nScripts = UBound($aScriptTable[2])
	If @error Then MsgBox(16, "Error", "Amount of scripts could not be determined.")
; determine start time of test
$aStartTime = _ADO_Execute($oConnection, "SELECT * FROM Result", True)
$nStartTime = ($aStartTime[2])[0][4]
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
	$sQuery = "SELECT Event_map.[Event Name], Event_meter.[End Time], Event_meter.Value FROM Event_map INNER JOIN Event_meter ON Event_map.[Event ID] = Event_meter.[Event ID] WHERE (((Event_meter.[Script ID])=" & $i & ")) ORDER BY Event_map.[Event Name], Event_meter.[End Time];"
	$rc = _ADO_Execute($oConnection, $sQuery, True)
	If Not IsArray($rc) Or @error Then
		MsgBox(16, "Query measurements per script", "Query for script " & $sScriptName & " failed.")
		ContinueLoop
	Else
		$aMeasurementsPerScript[$i] = (_ADO_Execute($oConnection, $sQuery, True))[2]
		If @error Then MsgBox(16, "Query measurements per script", "Query for script " & $sScriptName & " failed.")
		$rc = ProcessScript($aMeasurementsPerScript[$i], $sScriptName)
		if @error Or Not $rc Then MsgBox(16, "Error", "Processing for script " & $sScriptName & " failed.")
	EndIf
Next

; Clean Up
_ADO_Connection_Close($oConnection)
$oConnection = Null

TrayTip("Ready!", "LoadRunner metrics exported to Graphite", 4)
Sleep(4000)

TCPShutdown()

Func ProcessScript ($aMeasurements, ByRef $sScript)
	TrayTip("", "Now processing script " & $sScript, 1)
	ConsoleWrite(@CRLF & "Now processing script " & $sScript & ":" & @CRLF & @CRLF)
	$aUniqueTransactions = _ArrayUnique($aMeasurements)
;~ 	_ArrayDisplay($aUniqueTransactions)
	For $i = 1 to $aUniqueTransactions[0]
		$aTransactionSplitIndices = _ArrayFindAll($aMeasurements, $aUniqueTransactions[$i])
		Local $aTransactionsSplit[(UBound($aTransactionSplitIndices))][2] ; dimension array to store all measurements per transaction per script
		For $index = 0 to UBound($aTransactionSplitIndices) - 1
			$aTransactionsSplit[$index][0] = $aMeasurements[$aTransactionSplitIndices[$index]][1]
			$aTransactionsSplit[$index][1] = $aMeasurements[$aTransactionSplitIndices[$index]][2]
		Next
		$ret = ProcessTransaction($aTransactionsSplit, $aUniqueTransactions[$i], $sScript)
		if @error Or Not $ret Then
			MsgBox(16, "Error", "Processing transactions for transaction " & $aUniqueTransactions[$i] & " belonging to script " & $sScriptName & " failed.")
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
		MsgBox(16, "Error", "Processing buckets for transaction " & $sTransactionName & " belonging to script " & $sScript & " failed.")
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
		If IsArray($aBuckets[$i]) Then ExportToGraphite($sGraphiteRootNamespace & "." & StringReplace($sScript, " ", "_") & "." & StringReplace($sTransactionName, " ", "_"), _ArrayMin($aBuckets[$i]), Average($aBuckets[$i]), _ArrayMax($aBuckets[$i]), Percentile($aBuckets[$i], $nPercentile), $nStartTime + (($i + 1) * $nGraphiteResolution))
	Next
	Return True
EndFunc

Func ExportToGraphite ($sMetricPath, $nMin, $nAvg, $nMax, $nPerc, $nEpoch)
	If $iSocket = Null Then $iSocket = TCPConnect($sGraphiteHost, $nGraphitePort)
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
;~ 		ConsoleWrite("Payload bytes: " & $nPayloadBytes & @CRLF)

	If $nPayloadBytes > 256000 Then  ; after 250KB use new TCP connection
		$ret = TCPCloseSocket($iSocket)
		$iSocket = Null
		$nPayloadBytes = 0
		If @error Then
			ConsoleWriteError("TCPClose errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & @CRLF)
			$ret = TCPShutdown()
			If @error Then ConsoleWriteError("TCPShutdown errorcode: " & @error & " socket: " & $iSocket & " metric path: " & $sMetricPath & @CRLF)
			Exit
		EndIf
	EndIf
EndFunc ; ExportToGraphite

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
