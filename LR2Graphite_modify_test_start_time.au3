#cs ----------------------------------------------------------------------------

	AutoIt Version: 3.3.14.2
	Author:         Okke Garling

	Script Function:
	Change start time of LoadRunner test

#ce ----------------------------------------------------------------------------

#AutoIt3Wrapper_icon=targets-io.ico
#include "ADO.au3"
#include <Array.au3>

$sMDB_FileFullPath = FileOpenDialog("Location of LoadRunner mdb database file?", "", "LoadRunner analysis (*.mdb)")
If $sMDB_FileFullPath = "" Or @error Then
	MsgBox(16, "Error", "No valid location for LoadRunner mdb databbase file specified.")
	Exit
EndIf

$nHoursFromNow = InputBox("Graphite", "How many hours from now in the present?", "4")
If $nHoursFromNow = "" Or @error Then
	MsgBox(16, "Error", "No valid input specified.")
	Exit
EndIf

If FileCopy($sMDB_FileFullPath, $sMDB_FileFullPath & ".old", 1) Then MsgBox(64, "Information", "Copy of database stored to " & $sMDB_FileFullPath & ".old")

$sConnectionString = "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & $sMDB_FileFullPath
$oConnection = _ADO_Connection_Create()
_ADO_Connection_OpenConString($oConnection, $sConnectionString)
If @error Then SetError(@error, @extended, $ADO_RET_FAILURE)

; determine start time of test
$nEpochNow = _DateDiff('s', "1970/01/01 00:00:00", _NowCalc())
ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $nEpochNow = ' & $nEpochNow & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console
$nEpochStartTest = $nEpochNow - ($nHoursFromNow * 3600)
$aStartTime = _ADO_Execute($oConnection, "UPDATE Result SET [Start Time]=" & $nEpochStartTest & ";", True)

; Clean Up
_ADO_Connection_Close($oConnection)
$oConnection = Null

TrayTip("Ready!", "LoadRunner analysis database test start time modified", 4)
Sleep(4000)
