
#include-once
#Region ADO.au3 - Option, Includes, Setup
#Tidy_Parameters=/sort_funcs /reel
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w 7
#include-once
#include "ADO_CONSTANTS.au3"
#include <Array.au3>
#include <Date.au3>
#include <StringConstants.au3>
#include <AutoItConstants.au3>

#EndRegion ADO.au3 - Option, Includes, Setup

#Region ADO.au3 - UDF Header
; #INDEX# ========================================================================
; Title .........: ADO.au3
; AutoIt Version : 3.3.10.2++
; Language ......: English
; Description ...: A collection of Function for use with an ADO database like MS SQL, MS Access ...
; Author ........: Chris Lambert, mLipok
; Modified ......: eltorro, Elias Assad Neto, CarlH
; Version .......: 2.1.13 BETA - Work in progress 2016/03/18
; ================================================================================

#CS
	2015/08/18
	.	new collection of Functions for EVENT handling - grouped in #Region ADO.au3 - Functions - Event's Handling

	2015/08/24
	.	using ADO_CONSTANTS.au3

	2015/09/02
	.	removed $oConnection = -1, currently all function use ByRef $oConnection

	2015/09/15
	.	Renamed: $_eSQL_RESULT_ >> $ADOSQL_RESULT_ - mLipok
	.	Renamed: $_eSQL_ERROR_ >> $ADOSQL_ERROR_ - mLipok

	2015/10/04 >> 2015/11/06
	.	Renamed: Enums: $ADOSQL_RESULT_ >> $ADO_RET_- mLipok
	.	Renamed: Enums: $ADOSQL_ERROR_ >> $ADO_ERR_- mLipok
	.	Renamed: Enums: $ADO_RET_ERROR >> $ADO_RET_FAILURE- mLipok
	.	Renamed: Enums: $ADO_RET_OK >> $ADO_RET_SUCCESS- mLipok
	.	Renamed: Enums: $ADO_ERR_PARAMETERS >> $ADO_ERR_INVALIDPARAMETERTYPE - mLipok
	.	Renamed: Enums: $ADO_ERR_OK >> $ADO_ERR_SUCCESS - mLipok
	.	Renamed: Function: _SQLVerison >> _ADO_Version - mLipok
	.	Renamed: Function: _SQL_Close >> _ADO_Connection_Close - mLipok
	.	Renamed: Function: _SQL_Startup >> _ADO_Connection_Create - mLipok
	.	Renamed: Function: _SQL_Execute >> _ADO_Execute - mLipok
	.	Renamed: Function: __SQL_EVENT >> __ADO_EVENT - mLipok
	.	New: Function: __ADO_IsValidObjectType - mLipok
	.	New: Enums: $ADO_EXT_INTERNALFUNCTION - mLipok
	.	New: Function: _ADO_Recordset_ToArray - mLipok
	.	Refactored: _SQL_GetTable2D - mLipok
	.	Remove: $ADO_ERR_OTHER >> $ADO_ERR_GENERAL - mLipok
	.	Changed: Function: _SQL_FetchNames : Parameter $oRecordset is now ByRef - mLipok
	.	Added: Function: Parameter: _ADO_Recordset_ToArray >> $bFieldNamesInFirstRow = True - mLipok
	.			this was a speed issue as the entire table was moved step by step
	.	Added: Enums: $ADO_RS_ARRAY_* for use with Return form _ADO_Recordset_ToArray when $bFieldNamesInFirstRow was used - mLipok
	.	Added: Function: _ADO_Recordset_Display - mLipok
	.	Added: Function: __ADO_RecordsetArray_Display - mLipok
	.	Renamed: Variable: $oADODB_Connection >> $oConnection - mLipok
	.	Added: Function: _ADO_Execute: Validation for $oConnection - mLipok
	.	New: Function: __ADO_Command_IsValid - mLipok
	.	New: Function: __ADO_Connection_IsValid - mLipok
	.	New: Function: __ADO_Recordset_IsValid - mLipok
	.	New: Enums: $ADO_ERR_NOCURRENTRECORD - mLipok
	.	Renamed: $ADO_* >> $ADO_* - mLipok
	.	Renamed: _SQL_CommandTimeout >> _ADO_Connection_CommandTimeout - mLipok
	.
	.
	2015/11/06 >>
	.	Removed: Function: _SQL_GetErrMsg() - mLipok
	.	Removed: Variable: $g__sSQL_ErrorDescription - mLipok
	.	Renamed: Function: _SQL_PROVIDER_VERSION >> _ADO_MSSQL_GetProviderVersion - mLipok
	.	Renamed: Function: _SQL_DRIVER_VERSION >> _ADO_MSSQL_GetDriverVersion - mLipok
	.	Removed: Parameter: Function: _SQL_FetchData() $aRow - mLipok
	.	Removed: Parameter: Function: _SQL_FetchNames() $aRow - mLipok
	.	Refactored: _SQL_FetchNames - mLipok
	.	Refactored: _SQL_FetchData - mLipok
	.	Changed: _ADO_Recordset_Display - parameters order - $iAlternateColors <> $bFieldNamesInFirstRow mLipok
	.	Changed: _ADO_Recordset_Display - $bFieldNamesInFirstRow  now default is = False  - mLipok
	.	Added: Function: _ADO_RecordsetArray_IsValid - mLipok
	.	Refactored: Function: __ADO_RecordsetArray_Display - added _ADO_RecordsetArray_IsValid  - mLipok
	.	New: Function: _ADO_RecordsetArray_GetContent - mLipok
	.	New: Function: _ADO_RecordsetArray_GetFieldNames - mLipok
	.
	.
	2016/02/24
	.	Removed: Function: $__sSQL_Last_ConnectionString - mLipok
	.	Removed: Function: _SQL_QuerySingleRowAsString - mLipok
	.	Removed: Function: _SQL_QuerySingleRow - mLipok
	.	Removed: Function: _SQL_GetTable - mLipok
	.	Removed: Function: _SQL_GetTableAsString - mLipok
	.	Removed: Function: _ADO_SQLConnection_DBName - mLipok
	.	Removed: Function: _SQL_RegisterErrorHandler - mLipok
	.	Removed: Function: _SQL_UnRegisterErrorHandler - mLipok
	.	Removed: Function: _SQL_GetTable2D --> look in _ADO_Execute --> third parameter $bReturnAsArray - mLipok
	.	Added: 	Parameter in function: $bReturnAsArray - mLipok
	.
	.	Changed: Function: _ADO_Recordset_ToArray - Parameter - $bFieldNamesInFirstRow is not optional any more - mLipok
	.			(This is first step to change Behavior)
	.	Renamed: Function: _ADO_RecordsetArray_IsValid >> __ADO_RecordsetArray_IsValid - is now INTERNAL - mLipok
	.	Renamed: Function: _SQL_AccessConnect >> _ADO_Connection_OpenAccess - mLipok
	.	Renamed: Function: _SQL_ExcelConnect >> _ADO_Connection_OpenExcel - mLipok
	.	Renamed: Function: _ADO_Connection_OpenJet >> _ADO_Connection_OpenJet - mLipok
	.	Renamed: Function: _ADO_SQLConnectionOpen >> _ADO_Connection_OpenMSSQL - mLipok
	.	Refactored:	_ADO_Connection_OpenMSSQL : $sAPPNAME - mLipok
	.	Change:	_ADO_Connection_OpenMSSQL : parameter : reordering - mLipok
	.	Added:	_ADO_Connection_OpenMSSQL : parameter : $sWSID - mLipok
	.	Added:	_ADO_Connection_OpenMSSQL : parameter : $bUseProviderInsteadDriver - mLipok
	.	Change:	__ADO_MSSQL_CONNECTION_STRING_SQLAuth : parameter : reordering - mLipok
	.	Added:	__ADO_MSSQL_CONNECTION_STRING_SQLAuth : parameter : $sAPPNAME - mLipok
	.	Added:	Function: _ADO_Connection_PropertiesToArray - mLipok
	.			Thanks to @water for wiki tutorial: https://www.autoitscript.com/wiki/ADO_Tools
	.
	2016/02/24 FIRST PUBLIC RELEASE
	.
	2016/02/24 '2.1.7 BETA'
	.	Changed: Function: _ADO_Recordset_ToArray: Parameter is now Optional: $bFieldNamesInFirstRow = False - mLipok
	.	Removed: Function: _ADO_ExecuteQueryToArray --> look in _ADO_Execute --> third parameter $bReturnAsArray - mLipok
	.	Changed: Enums and constants moved to: ADO_CONSTANTS.au3 - mLipok
	.			Thanks to @BrewManNH
	.	Changed: ADO_CONSTANTS.au3: New region: #Region ADO_CONSTANTS.au3 - ADO.au3 UDF Constants  - mLipok
	.	Changed: ADO_CONSTANTS.au3: New region: #Region ADO_CONSTANTS.au3 - MSDN Enumerated Constants  - mLipok
	.	New: Function: _ADO_UDFVersion() - mLipok
	.	Removed: Global Variable $__sSQL_UDFVersion -->> look for: _ADO_UDFVersion() - mLipok
	.	Added: New example: ADO_EXAMPLE__PostgreSQL.au3 - mLipok
	.
	.
	2016/02/26 '2.1.8 BETA'
	.	Added: Function: _ADO_ConnectionString_Access - mLipok
	.	Added: Function: _ADO_ConnectionString_Excel - mLipok
	.	Removed: Function: _ADO_Connection_OpenAccess - mLipok
	.			Look for: _ADO_Connection_OpenConString and _ADO_ConnectionString_Access
	.	Changed: Example: ADO_EXAMPLE__PostgreSQL.au3 >> ADO_EXAMPLE.au3 - mLipok
	.	ADO_EXAMPLE.au3: New Comments in script - mLipok
	.	ADO_EXAMPLE.au3: New Function: _Example_MSAccess() - mLipok
	.	ADO_EXAMPLE.au3: New Function: _Example_MSExcel() - mLipok
	.	ADO_EXAMPLE.au3: New Function: _Example_MSSQL() - mLipok
	.	ADO_EXAMPLE.au3: Renamed Function: _Example_PostgreSQL() - mLipok
	.
	.
	2016/03/01 '2.1.9 BETA'
	.	ADO_CONSTANTS.au3: CleanUp: $ADO_adErr - mLipok
	.	ADO_CONSTANTS.au3: CleanUp/Fixed: _ADO_ERROR_Description() - mLipok
	.	Moved: Function: _ADO_ERROR_Description - From: ADO_CONSTANTS.au3 To: ADO.au3 - mLipok
	.	Added: Function: _ADO_GetProvidersList - mLipok
	.			Thanks to @water for wiki tutorial: https://www.autoitscript.com/wiki/ADO_Tools
	.	Removed: Function: _SQL_FetchData - mLipok
	.	Removed: Function: _SQL_FetchNames - mLipok
	.	Changed: Function: _ADO_EVENTS_SetUp - Default is Disabled - mLipok
	.	Refactored: Function: __ADO_RecordsetArray_Display - mLipok
	.	ADO_EXAMPLE.au3: New Function: _Example_MySQL() - mLipok
	.
	.
	2016/03/01 '2.1.10 BETA'
	.	New: Function: _ADO_ConnectionString_MySQL() - mLipok
	.	Added: in few function added COM Error Handler - mLipok
	.	ADO_EXAMPLE.au3: Function: _Example_MySQL() - some change - mLipok
	.
	.
	2016/03/08 '2.1.11 BETA'
	.	New: Function: _ADO_OpenSchema_Catalogs - mLipok
	.	New: Function: _ADO_OpenSchema_Tables - mLipok
	.	New: Function: _ADO_OpenSchema_Columns - mLipok
	.	New: Function: _ADO_OpenSchema_Indexes - mLipok
	.	New: Function: _ADO_OpenSchema_Views - mLipok
	.	New: Function: _ADO_Schema_GetAllCatalogs - mLipok
	.	New: Function: _ADO_Schema_GetAllTables - mLipok
	.	New: Function: _ADO_Schema_GetAllViews - mLipok
	.	Removed: Function: _SQL_GetTableName - mLipok
	.	Removed: Function: _ADO_Connection_OpenExcel - mLipok
	.			Look for: _ADO_Connection_OpenConString and _ADO_ConnectionString_Excel
	.	Changed: ADO_EXAMPLE.au3 - _Example_MySQL() - mLipok
	.	Changed: ADO_EXAMPLE.au3 - _Example_PostgreSQL() - mLipok
	.	Renamed: Function: _ADO_Command >> _ADO_Command_Create - mLipok
	.	Changed: Function: _ADO_Command_Create: Parameters removed - $sQuery - mLipok
	.	New: Function: _ADO_Command_CreateParameter - mlipok
	.	New: Function: _ADO_Command_Execute - mlipok
	.	Added: ADO_EXAMPLE.au3 - _Example_MSSQL_COMMAND_StoredProcedure() - mLipok
	.
	.
	2016/03/09 '2.1.12 BETA'
	.	New: Enums: $ADO_ERR_ISCLOSEDOBJECT - mLipok
	.	New: Function: __ADO_Connection_IsOpen - mLipok
	.			__ADO_Connection_IsOpen is a wrapper for __ADO_Connection_IsValid  which also check for $oConnection.state and set $ADO_ERR_ISCLOSEDOBJECT
	.			__ADO_Connection_IsOpen is now used in few functions which uses $oConnection
	.	Changed: Function: __ADO_Recordset_IsNotEmpty - checking $oRecordset.state and return $ADO_ERR_ISCLOSEDOBJECT - mLipok
	.	Changed: Function: _ADO_Command_Execute - mlipok
	.			Now return recordset
	.	Changed: ADO_EXAMPLE.au3 - _Example_MSSQL_COMMAND_StoredProcedure() - mLipok
	.	Removed: Function: _ADO_Connection_OpenJet - mLipok
	.			Look for: _ADO_Connection_OpenConString or _ADO_ConnectionString_Excel
	.
	.
	2016/03/18 '2.1.13 BETA'
	.	Changed: _ADO_COMErrorHandler - now showing also _ADO_UDFVersion()  - mLipok
	.	New: Enums: $ADO_ERR_ISNOTREADYOBJECT - mLipok
	.	Renamed: Function: __ADO_Connection_IsOpen >> __ADO_Connection_IsReady - mLipok
	.	Changed: Function: __ADO_Connection_IsReady : new feature checking connection state and seting  $ADO_ERR_ISNOTREADYOBJECT - mLipok
	.	New: Function: __ADO_Recordset_IsReady - mLipok
	.			__ADO_Recordset_IsReady is a wrapper for __ADO_Recordset_IsValid
	.				which also check for $oRecordset.state and set $ADO_ERR_ISCLOSEDOBJECT also $ADO_ERR_ISNOTREADYOBJECT
	.			__ADO_Recordset_IsReady is now used in few functions which uses $oRecordset
	.	Changed: Function: __ADO_Recordset_IsNotEmpty : now using __ADO_Recordset_IsReady instead __ADO_Recordset_IsValid - mLipok
	.			as __ADO_Recordset_IsReady is wrapper for __ADO_Recordset_IsValid
	.			so now __ADO_Recordset_IsNotEmpty checking old and new feature
	.
	.	!!!!!!!!!!!!!!!!!!!!!!!!
	.	Renamed: _ADO_ERROR_Description >> _ADO_MSDNErrorValueEnum_Description
	.	New: Function: _ADO_GetErrorDescription - mLipok
	.	New: Function: _ADO_ConsoleError - mLipok
	.
	.
	.
	@LAST
	.
	. 2015-09-01 _ADO_EVENTS_SetUp() not working properly
	.
	. TODO: https://msdn.microsoft.com/en-us/library/ms675114(v=vs.85).aspx
	. TODO:  Descripition to check:  On Success  - Returns $ADO_RET_SUCCESS

#CE

#EndRegion ADO.au3 - UDF Header

#Region ADO.au3 - Variable Declaration

Global $__g_fnFetchProgress = Null
; #VARIABLES# ====================================================================
#EndRegion ADO.au3 - Variable Declaration

#Region ADO.au3 - Functions

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_RecordsetArray_Display
; Description ...:
; Syntax ........: __ADO_RecordsetArray_Display(Byref $aRocordset[, $sTitle = ''[, $iAlternateColors = Default]])
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
;                  $sTitle              - [optional] a string value. Default is ''.
;                  $iAlternateColors    - [optional] an integer value. Default is Default.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_RecordsetArray_Display(ByRef $aRocordset, $sTitle = '', $iAlternateColors = Default)
	If __ADO_RecordsetArray_IsValid($aRocordset) Then
		Local $sArrayHeader = _ArrayToString($aRocordset[$ADO_RS_ARRAY_FIELDNAMES], '|')
		Local $aSelect = _ADO_RecordsetArray_GetContent($aRocordset)
		_ArrayDisplay($aSelect, $sTitle, "", 0, '|', $sArrayHeader, Default, $iAlternateColors)
		If @error Then Return SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

	ElseIf UBound($aRocordset) Then
		_ArrayDisplay($aRocordset, $sTitle, "", 0, Default, Default, Default, $iAlternateColors)
		If @error Then Return SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

	EndIf

	Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM1, $ADO_RET_FAILURE)

EndFunc   ;==>__ADO_RecordsetArray_Display

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_RecordsetArray_IsValid
; Description ...:
; Syntax ........: __ADO_RecordsetArray_IsValid(Byref $aRocordset)
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
; Return values .: True/False
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_RecordsetArray_IsValid(ByRef $aRocordset)
	If _
			UBound($aRocordset, $UBOUND_DIMENSIONS) = 1 _
			And UBound($aRocordset, $UBOUND_ROWS) = $ADO_RS_ARRAY_ENUMCOUNTR _
			And $aRocordset[$ADO_RS_ARRAY_GUID] = $ADO_RS_GUID _
			 Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, True)
	EndIf
	Return SetError($ADO_ERR_INVALIDARRAY, $ADO_EXT_DEFAULT, False)
EndFunc   ;==>__ADO_RecordsetArray_IsValid

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_Display
; Description ...: Display Recordset content with _ArrayDisplay()
; Syntax ........: _ADO_Recordset_Display(Byref $vRocordset[, $sTitle = ''[, $iAlternateColors = Default[,
;                  $bFieldNamesInFirstRow = False]]])
; Parameters ....: $vRocordset          - [in/out] a variant value.
;                  $sTitle              - [optional] a string value. Default is ''.
;                  $iAlternateColors    - [optional] an integer value. Default is Default.
;                  $bFieldNamesInFirstRow- [optional] a boolean value. Default is False.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_Display(ByRef $vRocordset, $sTitle = '', $iAlternateColors = Default, $bFieldNamesInFirstRow = False)
	Local $vResult = $ADO_RET_FAILURE
	If UBound($vRocordset) Then
		$vResult = __ADO_RecordsetArray_Display($vRocordset, $sTitle)
		Return SetError(@error, @extended, $vResult)
	ElseIf __ADO_Recordset_IsNotEmpty($vRocordset) = $ADO_RET_SUCCESS Then
		Local $aRecordset_GetRowsResult = _ADO_Recordset_ToArray($vRocordset, $bFieldNamesInFirstRow)
		$vResult = __ADO_RecordsetArray_Display($aRecordset_GetRowsResult, $sTitle, $iAlternateColors)
		Return SetError(@error, @extended, $vResult)
	Else
		Return SetError(@error, @extended, $ADO_RET_FAILURE) ; @error and @extended returned from __ADO_Recordset_IsNotEmpty
	EndIf

EndFunc   ;==>_ADO_Recordset_Display

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_Find
; Description ...: Searches a Recordset for the row that satisfies the specified criteria.
; Syntax ........: _ADO_Recordset_Find(Byref $oRecordset, $Criteria[, $SkipRows = 0[, $SearchDirection = $ADO_adSearchForward[, $Start = $ADO_adBookmarkCurrent]]])
; Parameters ....: $oRecordset          - [in/out] An unknown value.
;                  $Criteria            - An unknown value.
;                  $SkipRows            - [optional] An unknown value. Default is 0.
;                  $SearchDirection     - [optional] An unknown value. Default is $ADO_adSearchForward.
;                  $Start               - [optional] An unknown value. Default is $ADO_adBookmarkCurrent.
; Return values .: None - see remarks
; Author ........: mLipok
; Modified ......:
; Remarks .......: If the criteria is met, the current row position is set on the found record; otherwise, the position is set to the end (or start) of the Recordset.
; Related .......:
; Link ..........: http://msdn.microsoft.com/en-us/library/windows/desktop/ms676117(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_Find(ByRef $oRecordset, $Criteria, $SkipRows = 0, $SearchDirection = $ADO_adSearchForward, $Start = $ADO_adBookmarkCurrent)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler
	__ADO_Recordset_IsNotEmpty($oRecordset)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oRecordset.Find($Criteria, $SkipRows, $SearchDirection, $Start)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

EndFunc   ;==>_ADO_Recordset_Find

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_ToArray
; Description ...: Transform $oRecordset to an 2Dimensional Array
; Syntax ........: _ADO_Recordset_ToArray(Byref $oRecordset[, $bFieldNamesInFirstRow = False])
; Parameters ....: $oRecordset          - [in/out] an object.
;                  $bFieldNamesInFirstRow- [optional] a boolean value. Default is False.
; Return values .: On Success - $aResult
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_ToArray(ByRef $oRecordset, $bFieldNamesInFirstRow = False)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsNotEmpty($oRecordset)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	; save current Recordset rows postion to $oRecordset_Bookmark
	Local $oRecordset_Bookmark = Null
	If $oRecordset.Supports($ADO_adBookmark) Then $oRecordset_Bookmark = $oRecordset.Bookmark

	Local $aRecordset_GetRowsResult = $oRecordset.GetRows()
	If @error Then ; Trap COM error, report and return
		Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)
	ElseIf UBound($aRecordset_GetRowsResult) Then
		Local $aResult[0]

		; Restore Recordset row position from stored $oRecordset_Bookmark
		If $oRecordset_Bookmark = Null Then
			$oRecordset.moveFirst()
		Else
			$oRecordset.Bookmark = $oRecordset_Bookmark
		EndIf

		Local $iColumns_count = UBound($aRecordset_GetRowsResult, $UBOUND_COLUMNS)
		Local $iRows_count = UBound($aRecordset_GetRowsResult)

		If $bFieldNamesInFirstRow Then
			; Adjust the array to fit the column names and move all data down 1 row
			ReDim $aRecordset_GetRowsResult[$iRows_count + 1][$iColumns_count]

			; Move all records down
			For $iRow_idx = $iRows_count To 1 Step -1
				For $y = 0 To $iColumns_count - 1
					$aRecordset_GetRowsResult[$iRow_idx][$y] = $aRecordset_GetRowsResult[$iRow_idx - 1][$y]
				Next
			Next

			; Add the coloumn names
			For $iCol_idx = 0 To $iColumns_count - 1 ;get the column names and put into 0 array element
				$aRecordset_GetRowsResult[0][$iCol_idx] = $oRecordset.Fields($iCol_idx).Name
			Next
			$aResult = $aRecordset_GetRowsResult
			Return SetError($ADO_ERR_SUCCESS, $iRows_count + 1, $aResult)
		Else
			ReDim $aResult[$ADO_RS_ARRAY_ENUMCOUNTR]
			Local $aFiledNames_Temp[$iColumns_count]

			For $iCol_idx = 0 To $iColumns_count - 1 ;get the column names and put into 0 array element
				$aFiledNames_Temp[$iCol_idx] = $oRecordset.Fields($iCol_idx).Name
			Next
			$aResult[$ADO_RS_ARRAY_GUID] = $ADO_RS_GUID
			$aResult[$ADO_RS_ARRAY_FIELDNAMES] = $aFiledNames_Temp
			$aResult[$ADO_RS_ARRAY_RSCONTENT] = $aRecordset_GetRowsResult
			Return SetError($ADO_ERR_SUCCESS, $iRows_count, $aResult)
		EndIf
	EndIf

	Return SetError($ADO_ERR_RECORDSETEMPTY, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
EndFunc   ;==>_ADO_Recordset_ToArray

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Recordset_ToString
; Description ...:
; Syntax ........: _ADO_Recordset_ToString(Byref $oRecordset[, $sDelim = "|"[, $bReturnColumnNames = True]])
; Parameters ....: $oRecordset          - [in/out] an object.
;                  $sDelim              - [optional] a string value. Default is "|".
;                  $bReturnColumnNames  - [optional] a boolean value. Default is True.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Recordset_ToString(ByRef $oRecordset, $sDelim = "|", $bReturnColumnNames = True)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsNotEmpty($oRecordset)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	#forceref $bReturnColumnNames ; no yet implemented

	; save current Recordset rows postion to $oRecordset_Bookmark
	Local $oRecordset_Bookmark = Null
	If $oRecordset.Supports($ADO_adBookmark) Then $oRecordset_Bookmark = $oRecordset.Bookmark

	; GetString Method (ADO)
	; https://msdn.microsoft.com/en-us/library/ms676975(v=vs.85).aspx
	Local $sString = $oRecordset.GetString($ADO_adClipString, $oRecordset.RecordCount, $sDelim, @CR, 'Null')
	If @error Then ; Trap COM error, report and return
		Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)
	ElseIf IsString($sString) Then
		; Restore Recordset row position from stored $oRecordset_Bookmark
		If $oRecordset_Bookmark = Null Then
			$oRecordset.moveFirst()
		Else
			$oRecordset.Bookmark = $oRecordset_Bookmark
		EndIf

		Return SetError($ADO_ERR_SUCCESS, $oRecordset.RecordCount, $sString)
	Else
		Return SetError($ADO_ERR_RECORDSETEMPTY, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf
EndFunc   ;==>_ADO_Recordset_ToString

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_RecordsetArray_GetContent
; Description ...:
; Syntax ........: _ADO_RecordsetArray_GetContent(Byref $aRocordset)
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_RecordsetArray_GetContent(ByRef $aRocordset)
	__ADO_RecordsetArray_IsValid($aRocordset)
	If @error Then Return SetError(@error, @extended, Null)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aRocordset[$ADO_RS_ARRAY_RSCONTENT])
EndFunc   ;==>_ADO_RecordsetArray_GetContent

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_RecordsetArray_GetFieldNames
; Description ...:
; Syntax ........: _ADO_RecordsetArray_GetFieldNames(Byref $aRocordset)
; Parameters ....: $aRocordset          - [in/out] an array of unknowns.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_RecordsetArray_GetFieldNames(ByRef $aRocordset)
	__ADO_RecordsetArray_IsValid($aRocordset)
	If @error Then Return SetError(@error, @extended, Null)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aRocordset[$ADO_RS_ARRAY_FIELDNAMES])
EndFunc   ;==>_ADO_RecordsetArray_GetFieldNames
#EndRegion ADO.au3 - Functions

#Region ADO.au3 - Functions - Connection & Management

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Command_IsValid
; Description ...:
; Syntax ........: __ADO_Command_IsValid(Byref $oCommand)
; Parameters ....: $oCommand            - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Command_IsValid(ByRef $oCommand)
	Local $iValidationResult = __ADO_IsValidObjectType($oCommand, 'ADODB.Command')
	Return SetError(@error, @extended, $iValidationResult)
EndFunc   ;==>__ADO_Command_IsValid

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Connection_IsReady
; Description ...:
; Syntax ........: __ADO_Connection_IsReady(Byref $oConnection)
; Parameters ....: $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Connection_IsReady(ByRef $oConnection)
	Local $iValidationResult = __ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $oConnection.state = $ADO_adStateClosed Then
		Return SetError($ADO_ERR_ISCLOSEDOBJECT, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	ElseIf $oConnection.state <> $ADO_adStateOpen Then
		Return SetError($ADO_ERR_ISNOTREADYOBJECT, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $iValidationResult)
EndFunc   ;==>__ADO_Connection_IsReady

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Connection_IsValid
; Description ...:
; Syntax ........: __ADO_Connection_IsValid(Byref $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Connection_IsValid(ByRef $oConnection)
	Local $iValidationResult = __ADO_IsValidObjectType($oConnection, 'ADODB.Connection')
	Return SetError(@error, @extended, $iValidationResult)
EndFunc   ;==>__ADO_Connection_IsValid

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_IsValidObjectType
; Description ...:
; Syntax ........: __ADO_IsValidObjectType(Byref $oObjectToCheck, $sRequiredProgID)
; Parameters ....: $oObjectToCheck      - [in/out] an object.
;                  $sRequiredProgID     - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: Descripition @TODO
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_IsValidObjectType(ByRef $oObjectToCheck, $sRequiredProgID)
	If Not IsString($sRequiredProgID) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_INTERNALFUNCTION, $ADO_RET_FAILURE)
	ElseIf $sRequiredProgID = '' Then
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_INTERNALFUNCTION, $ADO_RET_FAILURE)
	ElseIf Not IsObj($oObjectToCheck) Then
		Return SetError($ADO_ERR_ISNOTOBJECT, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	ElseIf StringInStr(ObjName($oObjectToCheck, $OBJ_PROGID), $sRequiredProgID) = 0 Then
		Return SetError($ADO_ERR_INVALIDOBJECTTYPE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
EndFunc   ;==>__ADO_IsValidObjectType

; #FUNCTION# ====================================================================================================================
; Name ..........: __ADO_MSSQL_CONNECTION_STRING_SQLAuth
; Description ...:
; Syntax ........: __ADO_MSSQL_CONNECTION_STRING_SQLAuth($sServer, $sDataBase, $sUserName, $sPassword[, $bUseProviderInsteadDriver = True])
; Parameters ....: $sServer             - A string value.
;                  $sDataBase           - A string value.
;                  $sUserName           - A string value.
;                  $sPassword           - A string value.
;                  $bUseProviderInsteadDriver- [optional] A binary value. Default is True.
; Return values .: $sConnectionString
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........: https://msdn.microsoft.com/pl-pl/library/ms130822(v=sql.110).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_MSSQL_CONNECTION_STRING_SQLAuth($sServer, $sDataBase, $sUserName, $sPassword, $sAppName = Default, $bUseProviderInsteadDriver = True)
	Local Static $sConnectionString = ''

	Local Static $sLastParameters = Default
	Local $sNewParameters = $sServer & $sDataBase & $sUserName & $sPassword & $sAppName & $bUseProviderInsteadDriver

	If $sLastParameters <> $sNewParameters Then
		If $bUseProviderInsteadDriver Then
			$sConnectionString = "PROVIDER=" & _ADO_MSSQL_GetProviderVersion() & ";SERVER=" & $sServer & ";DATABASE=" & $sDataBase & ";UID=" & $sUserName & ";PWD=" & $sPassword & ";"
			If $sAppName <> Default And $sAppName <> '' Then $sConnectionString &= 'Application Name=' & $sAppName & ';'
		Else
			$sConnectionString = "DRIVER={" & _ADO_MSSQL_GetDriverVersion() & "};SERVER=" & $sServer & ";DATABASE=" & $sDataBase & ";UID=" & $sUserName & ";PWD=" & $sPassword & ";"
			If $sAppName <> Default And $sAppName <> '' Then $sConnectionString &= 'APPNAME=' & $sAppName & ';'
		EndIf
		$sLastParameters = $sNewParameters
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $sConnectionString)

EndFunc   ;==>__ADO_MSSQL_CONNECTION_STRING_SQLAuth

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Recordset_IsNotEmpty
; Description ...:
; Syntax ........: __ADO_Recordset_IsNotEmpty(Byref $oRecordset)
; Parameters ....: $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Recordset_IsNotEmpty(ByRef $oRecordset)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Recordset_IsReady($oRecordset)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $oRecordset.state = $ADO_adStateClosed Then
		Return SetError($ADO_ERR_ISCLOSEDOBJECT, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	ElseIf $oRecordset.bof = -1 And $oRecordset.eof = True Then ; no current record
		Return SetError($ADO_ERR_NOCURRENTRECORD, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	ElseIf $oRecordset.RecordCount = 0 Then
		Return SetError($ADO_ERR_RECORDSETEMPTY, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf
EndFunc   ;==>__ADO_Recordset_IsNotEmpty

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Recordset_IsReady
; Description ...:
; Syntax ........: __ADO_Recordset_IsReady(Byref $oRecordset)
; Parameters ....: $oRecordset         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Recordset_IsReady(ByRef $oRecordset)
	__ADO_Recordset_IsValid($oRecordset)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $oRecordset.state = $ADO_adStateClosed Then
		Return SetError($ADO_ERR_ISCLOSEDOBJECT, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	ElseIf $oRecordset.state <> $ADO_adStateOpen Then
		Return SetError($ADO_ERR_ISNOTREADYOBJECT, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
EndFunc   ;==>__ADO_Recordset_IsReady

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_Recordset_IsValid
; Description ...:
; Syntax ........: __ADO_Recordset_IsValid(Byref $oRecordset)
; Parameters ....: $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_Recordset_IsValid(ByRef $oRecordset)
	Local $iValidationResult = __ADO_IsValidObjectType($oRecordset, 'ADODB.Recordset')
	Return SetError(@error, @extended, $iValidationResult)
EndFunc   ;==>__ADO_Recordset_IsValid

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Command_Create
; Description ...:
; Syntax ........: _ADO_Command_Create(Byref $oConnection[, $sQuery = ''[, $iCommandType = $ADO_adCmdText]])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sQuery              - [optional] a string value. Default is ''.
;                  $iCommandType        - [optional] an integer value. Default is $ADO_adCmdText.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Command_Create(ByRef $oConnection, $iCommandType = $ADO_adCmdText)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf Not IsInt($iCommandType) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	EndIf

	Local $oCommand = ObjCreate("ADODB.Command")
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oCommand.ActiveConnection = $oConnection
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oCommand.CommandType = $iCommandType
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oCommand)
EndFunc   ;==>_ADO_Command_Create

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Command_CreateParameter
; Description ...: Creates a new Parameter object with the specified properties.
; Syntax ........: _ADO_Command_CreateParameter(Byref $oCommand, $sName, $iSize, $vValue[, $iType = $ADO_adChar[, $iDirection = $ADO_adParamInputOutput ]])
; Parameters ....: $oCommand            - [in/out] an object.
;                  $sName               - a string value.
;                  $iSize               - an integer value.
;                  $vValue              - a variant value.
;                  $iType               - [optional] an integer value. Default is $ADO_adChar.
;                  $iDirection          - [optional] an integer value. Default is $ADO_adParamInputOutput .
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/ms677209(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_Command_CreateParameter(ByRef $oCommand, $sName, $iSize, $vValue, $iType = $ADO_adChar, $iDirection = $ADO_adParamInputOutput)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Command_IsValid($oCommand)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf Not IsString($sName) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	ElseIf $sName = '' Then
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	ElseIf Not IsInt($iSize) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM3, $ADO_RET_FAILURE)
	ElseIf Not $iSize > 0 Then
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_PARAM3, $ADO_RET_FAILURE)
	EndIf

	Local $oParameter = Null
	If $vValue = Default Then
		$oParameter = $oCommand.CreateParameter($sName, $iType, $iDirection, $iSize)
	Else
		$oParameter = $oCommand.CreateParameter($sName, $iType, $iDirection, $iSize, $vValue)
	EndIf
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	$oCommand.Parameters.Append($oParameter)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
EndFunc   ;==>_ADO_Command_CreateParameter

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Command_Execute
; Description ...: Executes the query, SQL statement, or stored procedure specified in the CommandText or CommandStream property of the Command object.
; Syntax ........: _ADO_Command_Execute(Byref $oCommand, $sQuery)
; Parameters ....: $oCommand            - [in/out] an object.
;                  $sQuery              - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/ms681559(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_Command_Execute(ByRef $oCommand, $sQuery)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Command_IsValid($oCommand)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf Not IsString($sQuery) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	ElseIf $sQuery = '' Then
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	EndIf

	$oCommand.CommandText = $sQuery
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $iRecordsAffected = -1
	Local $oRecordset = $oCommand.Execute($iRecordsAffected)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $iRecordsAffected, $oRecordset)
EndFunc   ;==>_ADO_Command_Execute

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_Close
; Description ...: Closes an open ADODB.Connection
; Syntax.........:  _ADO_Connection_Close (ByRef $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: On Success - Returns $ADO_RET_SUCCESS
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Connection_Close(ByRef $oConnection)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oConnection.Close
	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

EndFunc   ;==>_ADO_Connection_Close

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_CommandTimeout
; Description ...: Sets and retrieves SQL CommandTimeout
; Syntax.........:  _ADO_Connection_CommandTimeout(ByRef $oConnection,$iTimeout)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $iTimeout   			- The timeout period to set if left blank the current value will be retrieved
; Return values .: On Success - Returns SQL Command timeout period
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........; https://msdn.microsoft.com/en-us/library/ms678265(v=vs.85).aspx
; Example .......; no
; ==============================================================================
Func _ADO_Connection_CommandTimeout(ByRef $oConnection, $iTimeOut = Default)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $iTimeOut = Default Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.CommandTimeout)
	ElseIf Not IsInt($iTimeOut) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	Else
		$oConnection.CommandTimeout = $iTimeOut
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.CommandTimeout)
	EndIf
EndFunc   ;==>_ADO_Connection_CommandTimeout

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Connection_Create
; Description ...: Creates ADODB.Connection object
; Syntax.........:  _ADO_Connection_Create()
; Parameters ....: None
; Return values .: On Success - Returns $oConnection Object
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Connection_Create()
	Local $oConnection = ObjCreate("ADODB.Connection")
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection)
EndFunc   ;==>_ADO_Connection_Create

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_OpenConString
; Description ...: Open Connection based on Connection String passed to the function
; Syntax ........: _ADO_Connection_OpenConString(Byref $oConnection, $sConnectionString)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sConnectionString   - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: Description TODO
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_OpenConString(ByRef $oConnection, $sConnectionString)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oConnection.Open($sConnectionString)
	If @error Then Return SetError($ADO_ERR_CONNECTION, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)

EndFunc   ;==>_ADO_Connection_OpenConString

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_OpenMSSQL
; Description ...: Starts a Database Connection to Microsoft SQL Server
; Syntax ........: _ADO_Connection_OpenMSSQL(Byref $oConnection, $sServer, $sDBName, $sUserName, $sPassword[, $sAppName = Default[,
;                  $sWSID = Default[, $bSQLAuth = True [, $bUseProviderInsteadDriver = True]]]])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sServer             - a string value. The server to connect to.
;                  $sDBName             - a string value. The database name to open.
;                  $sUserName           - a string value. Username for database access.
;                  $sPassword           - a string value. Password for database user.
;                  $sAppName            - [optional] a string value. Default is Default.
;                  $sWSID               - [optional] a string value. Default is Default.
;                  $bSQLAuth            - [optional] a boolean value. Default is True.
;                  $bUseProviderInsteadDriver- [optional] a boolean value. Default is True.
; Return values .: On Success - Returns $ADO_RET_SUCCESS
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/pl-pl/library/ms130822(v=sql.110).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_OpenMSSQL(ByRef $oConnection, $sServer, $sDBName, $sUserName, $sPassword, $sAppName = Default, $sWSID = Default, $bSQLAuth = True, $bUseProviderInsteadDriver = True)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	If $oConnection.State = $ADO_adStateOpen Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf

	Local $sConnectionString = ''
	If $bSQLAuth = True Then
		$sConnectionString = __ADO_MSSQL_CONNECTION_STRING_SQLAuth($sServer, $sDBName, $sUserName, $sPassword, $sAppName, $bUseProviderInsteadDriver)
	Else
		$oConnection.Properties("Integrated Security").Value = "SSPI"
		$oConnection.Properties("User ID") = $sUserName
		$oConnection.Properties("Password") = $sPassword
		$sConnectionString = "DRIVER={SQL Server};SERVER=" & $sServer & ";DATABASE=" & $sDBName & ";"
		$sConnectionString = "APP=" & $sAppName & ";"
	EndIf

	If $sWSID <> Default And $sWSID <> "" Then $sConnectionString &= "WSID=" & $sWSID & ";"

	$oConnection.Open($sConnectionString)
	If @error Then Return SetError($ADO_ERR_CONNECTION, @error, $ADO_RET_FAILURE)

	Local $vSQLOpenError_state = @error
	While Sleep(10)
		If Not $vSQLOpenError_state Or $oConnection.State = $ADO_adStateOpen Then
			__ADO_EVENTS_INIT($oConnection)
			Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
		Else
			Return SetError($ADO_ERR_CONNECTION, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
		EndIf
	WEnd

EndFunc   ;==>_ADO_Connection_OpenMSSQL

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_PropertiesToArray
; Description ...: List all Connection Properties
; Syntax ........: _ADO_Connection_PropertiesToArray(Byref $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: On Success - Returns $aProperties
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: water
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........: https://www.autoitscript.com/wiki/ADO_Tools
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_PropertiesToArray(ByRef $oConnection)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	; Property Object (ADO)
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms677577(v=vs.85).aspx
	Local $oProperties_coll = $oConnection.Properties
	Local $aProperties[$oProperties_coll.count][4]
	Local $iIndex = 0

	For $oProperty_enum In $oProperties_coll
		$aProperties[$iIndex][0] = $oProperty_enum.Name
		$aProperties[$iIndex][1] = $oProperty_enum.Type
		$aProperties[$iIndex][2] = $oProperty_enum.Value
		$aProperties[$iIndex][3] = $oProperty_enum.Attributes
		$iIndex += 1
	Next

	$oProperties_coll = Null
	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aProperties)

EndFunc   ;==>_ADO_Connection_PropertiesToArray

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Connection_Timeout
; Description ...: Sets and retrieves SQL ConnectionTimeout
; Syntax ........: _ADO_Connection_Timeout(Byref $oConnection[, $iTimeOut = Default])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $iTimeOut            - [optional] an integer value. Default is Default. The timeout period to set if left blank the current value will be retrieved
; Return values .: On Success - Returns Connection timeout period
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Connection_Timeout(ByRef $oConnection, $iTimeOut = Default)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf $iTimeOut = Default Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.ConnectionTimeout)
	ElseIf Not IsInt($iTimeOut) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_DEFAULT, $ADO_RET_FAILURE)
	Else
		$oConnection.Close
		$oConnection.ConnectionTimeout = $iTimeOut
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $ADO_RET_SUCCESS)
	EndIf

EndFunc   ;==>_ADO_Connection_Timeout

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Execute
; Description ...: Executes an SQL Query
; Syntax ........: _ADO_Execute(Byref $oConnection, $sQuery[, $bReturnAsArray = False[, $bFieldNamesInFirstRow = False]])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
;                  $sQuery              - a string value. SQL Statement to be executed.
;                  $bReturnAsArray      - [optional] a boolean value. Default is False.
;                  $bFieldNamesInFirstRow- [optional] a boolean value. Default is False.
; Return values .: On Success - Returns $oRecordset object or $aRecordsetAsArray
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ===============================================================================================================================
Func _ADO_Execute(ByRef $oConnection, $sQuery, $bReturnAsArray = False, $bFieldNamesInFirstRow = False)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	ElseIf Not IsString($sQuery) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	ElseIf $sQuery = '' Then
		Return SetError($ADO_ERR_INVALIDPARAMETERVALUE, $ADO_EXT_PARAM2, $ADO_RET_FAILURE)
	ElseIf Not IsBool($bReturnAsArray) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM3, $ADO_RET_FAILURE)
	ElseIf Not IsBool($bFieldNamesInFirstRow) Then
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_PARAM4, $ADO_RET_FAILURE)
	EndIf

	Local $oRecordset = $oConnection.Execute($sQuery)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	If $bReturnAsArray Then
		Local $aRecordsetAsArray = _ADO_Recordset_ToArray($oRecordset, $bFieldNamesInFirstRow)
		Return SetError(@error, @extended, $aRecordsetAsArray)
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oRecordset)

EndFunc   ;==>_ADO_Execute

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_GetProvidersList
; Description ...: This tool lists all available providers installed on the computer.
; Syntax ........: _ADO_GetProvidersList()
; Parameters ....: None
; Return values .: $aResult - a list of all available providers installed on the computer
; Author ........: water
; Modified ......: mLipok
; Remarks .......: based on: _ADO_OLEDBProvidersList
; Related .......:
; Link ..........: https://www.autoitscript.com/wiki/ADO_Tools
; Example .......: No
; ===============================================================================================================================
Func _ADO_GetProvidersList()
	Local $sKey = "HKCR\CLSID"
	Local $iIndexReg = 1, $iIndexResult = 0
	Local $iMax = 100000, $iMin = 1, $iPrevious = $iMin, $iCurrent = $iMax / 2
	Local $aResult[200][3]

	ProgressOn("OLE DB Providers", "Processing the Registry", "", Default, Default, $DLG_MOVEABLE)

	; Count the number of keys
	While 1
		RegEnumKey($sKey, $iCurrent)
		If @error = -1 Then ; Requested subkey (key instance) out of range
			$iMax = $iCurrent
			$iCurrent = Int(($iMin + $iMax) / 2)
			$iPrevious = $iMax
		Else
			If $iPrevious <= ($iCurrent + 1) And $iPrevious >= ($iCurrent - 1) Then ExitLoop
			$iMin = $iCurrent
			$iCurrent = Int(($iMin + $iMax) / 2)
			$iPrevious = $iMin
		EndIf
	WEnd

	Local $iPercent = 0
	Local $sKeyValue = '', $sSubKey = ''
	; Process registry
	While 1
		If Mod($iIndexReg, 10) = 0 Then
			$iPercent = Int($iIndexReg * 100 / $iCurrent)
			ProgressSet($iPercent, $iIndexReg & " keys of " & $iCurrent & " processed (" & $iPercent & "%)")
		EndIf
		$sSubKey = RegEnumKey($sKey, $iIndexReg)
		If @error Then ExitLoop

		$sKeyValue = RegRead($sKey & "\" & $sSubKey, "OLEDB_SERVICES")
		If @error = 0 Then
			$aResult[$iIndexResult][0] = $sKey & "\" & $sSubKey
			$aResult[$iIndexResult][1] = RegRead($sKey & "\" & $sSubKey, "")
			$aResult[$iIndexResult][2] = RegRead($sKey & "\" & $sSubKey & "\OLE DB Provider", "")
			$iIndexResult += 1
		EndIf

		$iIndexReg += 1
	WEnd

	ProgressOff()
	ReDim $aResult[$iIndexResult][3]

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $aResult)
	#forceref $sKeyValue
EndFunc   ;==>_ADO_GetProvidersList

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_MSSQL_GetDriverVersion
; Description ...: check for newer DRIVER parameter for CONNECTIONSTRING
; Syntax ........: _ADO_MSSQL_GetDriverVersion()
; Parameters ....: none.
; Return values .: $s_ADO_MSSQL_GetDriverVersion
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_MSSQL_GetDriverVersion()
	Local Static $s_ADO_MSSQL_GetDriverVersion = Default
	If $s_ADO_MSSQL_GetDriverVersion = Default Then
;~ 		Local  $sSQL_NCLI_2014 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2012 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2008 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 10.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2005 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Native Client\CurrentVersion', 'Version') ; For SQL Server 2005
		Select
;~ 			Case  $sSQL_NCLI_2014 <> ''
;~ 				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server Native Client 11.0'
			Case $sSQL_NCLI_2012 <> ''
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server Native Client 11.0'
			Case $sSQL_NCLI_2008 <> ''
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server Native Client 10.0'
			Case $sSQL_NCLI_2005 <> ''
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Native Client'
			Case Else
				$s_ADO_MSSQL_GetDriverVersion = 'SQL Server'
		EndSelect
	EndIf
	Return $s_ADO_MSSQL_GetDriverVersion

EndFunc   ;==>_ADO_MSSQL_GetDriverVersion

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_MSSQL_GetProviderVersion
; Description ...: check for newer PROVIDER parameter for CONNECTIONSTRING
; Syntax ........: _ADO_MSSQL_GetProviderVersion()
; Parameters ....: none.
; Return values .: $s_ADO_MSSQL_GetProviderVersion
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_MSSQL_GetProviderVersion()
	Local Static $s_ADO_MSSQL_GetProviderVersion = Default
	If $s_ADO_MSSQL_GetProviderVersion = Default Then
;~ 		Local  $sSQL_NCLI_2014 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2012 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 11.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2008 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Server Native Client 10.0\CurrentVersion', 'Version') ; For SQL Server 2008/SQL Server 2008 R2
		Local $sSQL_NCLI_2005 = RegRead('HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Microsoft SQL Native Client\CurrentVersion', 'Version') ; For SQL Server 2005
		Select
;~ 			Case  $sSQL_NCLI_2014 <> ''
;~ 				$s_ADO_MSSQL_GetProviderVersion = 'SQL Server Native Client 11.0'
			Case $sSQL_NCLI_2012 <> ''
				$s_ADO_MSSQL_GetProviderVersion = 'SQLNCLI11'
			Case $sSQL_NCLI_2008 <> ''
				$s_ADO_MSSQL_GetProviderVersion = 'SQLNCLI10'
			Case $sSQL_NCLI_2005 <> ''
				$s_ADO_MSSQL_GetProviderVersion = 'SQLNCLI'
			Case Else
				$s_ADO_MSSQL_GetProviderVersion = 'sqloledb'
		EndSelect
	EndIf
	Return $s_ADO_MSSQL_GetProviderVersion

EndFunc   ;==>_ADO_MSSQL_GetProviderVersion

; #FUNCTION# ===================================================================
; Name ..........: _ADO_Recordset_Create
; Description ...: Creates ADODB.Recordset object
; Syntax.........:  _ADO_Recordset_Create()
; Parameters ....: None
; Return values .: On Success - Returns $oRecordset Object
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........;
; Example .......; no
; ==============================================================================
Func _ADO_Recordset_Create()
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	Local $oRecordset = ObjCreate("ADODB.Recordset")
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oRecordset)
EndFunc   ;==>_ADO_Recordset_Create

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_Version
; Description ...:
; Syntax ........: _ADO_Version([ByRef $oConnection])
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: $oConnection.Version
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_Version(ByRef $oConnection)
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	Else
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $oConnection.Version)
	EndIf
EndFunc   ;==>_ADO_Version
#EndRegion ADO.au3 - Functions - Connection & Management

#Region ADO.au3 - Functions - ADDON - COM ERROR HANDLER
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_ComErrorHandler_InternalFunction
; Description ...:
; Syntax ........: __ADO_ComErrorHandler_InternalFunction($oCOMError)
; Parameters ....: $oCOMError           - an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_ComErrorHandler_InternalFunction($oCOMError)
	; Do nothing special, just check @error after suspect functions.
	#forceref $oCOMError
	Local $sUserFunction = _ADO_COMErrorHandler_UserFunction()
	If IsFunc($sUserFunction) Then $sUserFunction($oCOMError)
EndFunc   ;==>__ADO_ComErrorHandler_InternalFunction

; #FUNCTION# ===================================================================
; Name ..........: _ADO_COMErrorHandler
; Description ...: Autoit COM Error handler function
; Syntax ........: _ADO_COMErrorHandler()
; Parameters ....: None.
; Return values .: None
; Author ........: Chris Lambert
; Modified ......: mLipok
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: no
; ================================================================================
Func _ADO_COMErrorHandler($oADO_Error)
	; Error Object
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms677507(v=vs.85).aspx

	; Error Object Properties, Methods, and Events
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms678396(v=vs.85).aspx

	Local $HexNumber = Hex($oADO_Error.number, 8)
	Local $sSQL_ComErrorDescription = ''
	$sSQL_ComErrorDescription &= "ADO.au3 v." & _ADO_UDFVersion() & " (" & $oADO_Error.scriptline & ") : ==> COM Error intercepted !" & @CRLF
	$sSQL_ComErrorDescription &= "$oADO_Error.description is: " & @TAB & $oADO_Error.description & @CRLF
	$sSQL_ComErrorDescription &= "$oADO_Error.windescription: " & @TAB & $oADO_Error.windescription & @CRLF
	$sSQL_ComErrorDescription &= "$oADO_Error.number is: " & @TAB & $HexNumber & @CRLF
	$sSQL_ComErrorDescription &= "$oADO_Error.lastdllerror is: " & @TAB & $oADO_Error.lastdllerror & @CRLF
	$sSQL_ComErrorDescription &= "$oADO_Error.scriptline is: " & @TAB & $oADO_Error.scriptline & @CRLF

	; Source Property (ADO Error)
	; https://msdn.microsoft.com/en-us/library/windows/desktop/ms675830(v=vs.85).aspx
	$sSQL_ComErrorDescription &= "$oADO_Error.source is: " & @TAB & $oADO_Error.source & @CRLF
	$sSQL_ComErrorDescription &= "$oADO_Error.helpfile is: " & @TAB & $oADO_Error.helpfile & @CRLF
	$sSQL_ComErrorDescription &= "$oADO_Error.helpcontext is: " & @TAB & $oADO_Error.helpcontext & @CRLF

	#CS
		; NativeError Property (ADO)
		; https://msdn.microsoft.com/en-us/library/windows/desktop/ms678049(v=vs.85).aspx
		$sSQL_ComErrorDescription &= "$oADO_Error.NativeError is: " & @TAB & $oADO_Error.NativeError & @CRLF

		; SQLState Property
		; https://msdn.microsoft.com/en-us/library/windows/desktop/ms681570(v=vs.85).aspx
		$sSQL_ComErrorDescription &= "$oADO_Error.SQLState is: " & @TAB & $oADO_Error.SQLState & @CRLF

	#CE

	ConsoleWrite("###############################" & @CRLF & $sSQL_ComErrorDescription & "###############################" & @CRLF)
	SetError($ADO_ERR_GENERAL, $ADO_EXT_DEFAULT, $sSQL_ComErrorDescription)
EndFunc   ;==>_ADO_COMErrorHandler

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_COMErrorHandler_UserFunction
; Description ...: Set up user function to get COM Error Handler outside ADO.au3 UDF
; Syntax ........: _ADO_COMErrorHandler_UserFunction([$fnUserFunction = Default])
; Parameters ....: $fnUserFunction      - [optional] a floating point value. Default is Default.
; Return values .: On Success - $fnUserFunction_Static
;                  On Failure - Returns $ADO_RET_FAILURE and set @error to $ADO_ERR_*
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_COMErrorHandler_UserFunction($fnUserFunction = Default)
	; in case when user do not set his own function UDF must use internal function to avoid AutoItError
	Local Static $fnUserFunction_Static = ''

	If $fnUserFunction = Default Then
		; just return stored static variable
		Return $fnUserFunction_Static
	ElseIf IsFunc($fnUserFunction) Then
		; set and return static variable
		$fnUserFunction_Static = $fnUserFunction
		Return $fnUserFunction_Static
	Else
		; reset static variable
		$fnUserFunction_Static = ''
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, $ADO_EXT_DEFAULT, $fnUserFunction_Static)
	EndIf
EndFunc   ;==>_ADO_COMErrorHandler_UserFunction
#EndRegion ADO.au3 - Functions - ADDON - COM ERROR HANDLER

#Region ADO.au3 - Functions - ADDON - COM EVENT HANDLER
; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__BeginTransComplete
; Description ...: BeginTransComplete is called after the BeginTrans operation
; Syntax ........: __ADO_EVENT__BeginTransComplete($iTransactionLevel, Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $iTransactionLevel   - an integer value. A Long value that contains the new transaction level of the BeginTrans that caused this event.
;                  $oError              - [in/out] an object. An Error object. It describes the error that occurred if the value of EventStatusEnum is adStatusErrorsOccurred; otherwise it is not set.
;                  $i_adStatus          - an integer value. An EventStatusEnum status value. When any of these events is called, this parameter is set to adStatusOK if the operation that caused the event was successful, or to adStatusErrorsOccurred if the operation failed.
;											These events can prevent subsequent notifications by setting this parameter to adStatusUnwantedEvent before the event returns.
;                  $oConnection         - [in/out] an object. The Connection object for which this event occurred.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681493%28v=vs.85%29.aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__BeginTransComplete($iTransactionLevel, ByRef $oError, $i_adStatus, ByRef $oConnection)
;~ 	https://msdn.microsoft.com/en-us/library/ms681493(v=vs.85).aspx
	If Not _ADO_EVENTS_SetUp() Then Return

	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__BeginTransComplete:")
	__ADO_ConsoleWrite_Blue("   $iTransactionLevel=" & $iTransactionLevel)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $oConnection
EndFunc   ;==>__ADO_EVENT__BeginTransComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__CommitTransComplete
; Description ...: CommitTransComplete is called after the CommitTrans operation.
; Syntax ........: __ADO_EVENT__CommitTransComplete(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681493%28v=vs.85%29.aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__CommitTransComplete(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__CommitTransComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	If $i_adStatus = $ADO_adStatusErrorsOccurred Then
		__ADO_ConsoleWrite_Red("   $i_adStatus=$ADO_adStatusErrorsOccurred=" & $i_adStatus)
		__ADO_ConsoleWrite_Red("   STARTING:  $oConnection.RollbackTrans")
		$oConnection.RollbackTrans
	EndIf
	#forceref $oError, $oConnection
EndFunc   ;==>__ADO_EVENT__CommitTransComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__ConnectComplete
; Description ...: ConnectComplete Events (ADO)
; Syntax ........: __ADO_EVENT__ConnectComplete(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms676126(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__ConnectComplete(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__ConnectComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $i_adStatus, $oConnection
EndFunc   ;==>__ADO_EVENT__ConnectComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__Disconnect
; Description ...: Disconnect Events (ADO)
; Syntax ........: __ADO_EVENT__Disconnect($i_adStatus, Byref $oConnection)
; Parameters ....: $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms676126(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__Disconnect($i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__Disconnect:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $i_adStatus, $oConnection
EndFunc   ;==>__ADO_EVENT__Disconnect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__FetchComplete
; Description ...: FetchComplete Event (ADO)
; Syntax ........: __ADO_EVENT__FetchComplete(Byref $oError, $i_adStatus, Byref $oRecordset)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms677512(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__FetchComplete(ByRef $oError, $i_adStatus, ByRef $oRecordset)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__FEtchComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $oRecordset
EndFunc   ;==>__ADO_EVENT__FetchComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__FetchProgress
; Description ...: FetchProgress Event (ADO)
; Syntax ........: __ADO_EVENT__FetchProgress($iProgress, $iMaxProgress, $i_adStatus, Byref $oRecordset)
; Parameters ....: $iProgress           - an integer value.
;                  $iMaxProgress        - an integer value.
;                  $i_adStatus          - an integer value.
;                  $oRecordset          - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms675535(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__FetchProgress($iProgress, $iMaxProgress, $i_adStatus, ByRef $oRecordset)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__FetchProgress:")
	__ADO_ConsoleWrite_Blue("   $iProgress=" & $iProgress)
	__ADO_ConsoleWrite_Blue("   $iMaxProgress=" & $iMaxProgress)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	If IsFunc($__g_fnFetchProgress) Then
		$__g_fnFetchProgress($iProgress, $iMaxProgress, $i_adStatus, $oRecordset)
	EndIf
	#forceref $oRecordset
EndFunc   ;==>__ADO_EVENT__FetchProgress

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__InfoMessage
; Description ...:
; Syntax ........: __ADO_EVENT__InfoMessage(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms675859(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__InfoMessage(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__InfoMessage:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $i_adStatus, $oConnection
EndFunc   ;==>__ADO_EVENT__InfoMessage

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__RollbackTransComplete
; Description ...: RollbackTransComplete is called after the RollbackTrans operation.
; Syntax ........: __ADO_EVENT__RollbackTransComplete(Byref $oError, $i_adStatus, Byref $oConnection)
; Parameters ....: $oError              - [in/out] an object.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681493%28v=vs.85%29.aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__RollbackTransComplete(ByRef $oError, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__RollbackTransComplete:")
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oError, $oConnection
EndFunc   ;==>__ADO_EVENT__RollbackTransComplete

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__WillConnect
; Description ...: WillConnect Event (ADO)
; Syntax ........: __ADO_EVENT__WillConnect($sConnection_String, $sUserID, $sPassword, $iOptions, $i_adStatus, Byref $oConnection)
; Parameters ....: $sConnection_String   - a string value.
;                  $sUserID             - a string value.
;                  $sPassword           - a string value.
;                  $iOptions            - an integer value.
;                  $i_adStatus          - an integer value.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms680962(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__WillConnect($sConnection_String, $sUserID, $sPassword, $iOptions, $i_adStatus, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__WillConnect:")
	__ADO_ConsoleWrite_Blue("   $sConnection_String=" & $sConnection_String)
	__ADO_ConsoleWrite_Blue("   $sUserID=" & $sUserID)
	__ADO_ConsoleWrite_Blue("   $sPassword=" & $sPassword)
	__ADO_ConsoleWrite_Blue("   $iOptions=" & $iOptions)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oConnection
EndFunc   ;==>__ADO_EVENT__WillConnect

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENT__WillExecute
; Description ...: WillExecute Event (ADO)
; Syntax ........: __ADO_EVENT__WillExecute($sSource, $iCursorType, $iLockType, $iOptions, $i_adStatus, Byref $oCommand,
;                  Byref $oRecordset, Byref $oConnection)
; Parameters ....: $sSource             - a string value.
;                  $iCursorType         - an integer value.
;                  $iLockType           - an integer value.
;                  $iOptions            - an integer value.
;                  $i_adStatus          - an integer value.
;                  $oCommand            - [in/out] an object.
;                  $oRecordset          - [in/out] an object.
;                  $oConnection         - [in/out] an object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms680993(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENT__WillExecute($sSource, $iCursorType, $iLockType, $iOptions, $i_adStatus, ByRef $oCommand, ByRef $oRecordset, ByRef $oConnection)
	If Not _ADO_EVENTS_SetUp() Then Return
	__ADO_ConsoleWrite_Blue(" ADO EVENT fired function: __ADO_EVENT__WillExecute:")
	__ADO_ConsoleWrite_Blue("   $sSource=" & StringRegExpReplace($sSource, '\R', ' '))
	__ADO_ConsoleWrite_Blue("   $iCursorType=" & $iCursorType)
	__ADO_ConsoleWrite_Blue("   $iLockType=" & $iLockType)
	__ADO_ConsoleWrite_Blue("   $iOptions=" & $iOptions)
	__ADO_ConsoleWrite_Blue("   $i_adStatus=" & $i_adStatus)
	#forceref $oCommand, $oRecordset, $oConnection
EndFunc   ;==>__ADO_EVENT__WillExecute

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_EVENTS_INIT
; Description ...: Function to initialize ADO EVENTs handling
; Syntax ........: __ADO_EVENTS_INIT(Byref $oConnection)
; Parameters ....: $oConnection         - [in/out] an object. ADODB.Connection object.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_EVENTS_INIT(ByRef $oConnection)
	__ADO_Connection_IsValid($oConnection)
	If @error Then
		Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf
	Local Static $oADO_EventHandler = ''
	If $oADO_EventHandler = '' Then
		$oADO_EventHandler = ObjEvent($oConnection, "__ADO_EVENT__", "ConnectionEvents") ; @TODO check with #Au3Stripper_Parameters=/TL /debug
	Else
		$oADO_EventHandler = ''
	EndIf
EndFunc   ;==>__ADO_EVENTS_INIT

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_EVENTS_SetUp
; Description ...: Enable/Disable/Get - ADO EVENTs handling status
; Syntax ........: _ADO_EVENTS_SetUp([$bInitializeEvents = Default])
; Parameters ....: $bInitializeEvents   - [optional] a boolean value. Default is Default.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_EVENTS_SetUp($bInitializeEvents = Default)
	Local Static $bInitializeEvents_static = False

	If $bInitializeEvents = Default Then
		Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $bInitializeEvents_static)
	ElseIf IsBool($bInitializeEvents) Then
		$bInitializeEvents_static = $bInitializeEvents
		Return SetError($ADO_ERR_SUCCESS, 1, $bInitializeEvents_static)
	Else
		Return SetError($ADO_ERR_INVALIDPARAMETERTYPE, 2, $bInitializeEvents_static)
	EndIf
EndFunc   ;==>_ADO_EVENTS_SetUp
#EndRegion ADO.au3 - Functions - ADDON - COM EVENT HANDLER

#Region ADO.au3 - Functions - MISC

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_ConsoleWrite_Blue
; Description ...:
; Syntax ........: __ADO_ConsoleWrite_Blue($sText)
; Parameters ....: $sText               - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_ConsoleWrite_Blue($sText)
	ConsoleWrite(BinaryToString(StringToBinary('>>' & $sText & @CRLF, 4), 1))
EndFunc   ;==>__ADO_ConsoleWrite_Blue

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name ..........: __ADO_ConsoleWrite_Red
; Description ...:
; Syntax ........: __ADO_ConsoleWrite_Red($sText)
; Parameters ....: $sText               - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func __ADO_ConsoleWrite_Red($sText)
	ConsoleWrite(BinaryToString(StringToBinary('!!!!!!!!!' & $sText & @CRLF, 4), 1))
EndFunc   ;==>__ADO_ConsoleWrite_Red

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_ConsoleError
; Description ...:
; Syntax ........: _ADO_ConsoleError([$sDescription = ''[, $iError = @error[, $iExtended = @extended]]])
; Parameters ....: $sDescription        - [optional] a string value. Default is ''.
;                  $iError              - [optional] an integer value. Default is @error.
;                  $iExtended           - [optional] an integer value. Default is @extended.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_ConsoleError($sDescription = '', $iError = @error, $iExtended = @extended)
	Local $sDescription_Result = _ADO_GetErrorDescription($sDescription, True, $iError, $iExtended)
	ConsoleWrite('!!!!!!!!!' & $sDescription_Result & @CRLF)

	Return SetError($iError, $iExtended, $sDescription_Result)
EndFunc   ;==>_ADO_ConsoleError

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_GetErrorDescription
; Description ...:
; Syntax ........: _ADO_GetErrorDescription([$sDescription = ''[, $bShowHumanReadableDescription = True[, $iError = @error[,
;                  $iExtended = @extended]]]])
; Parameters ....: $sDescription        - [optional] a string value. Default is ''.
;                  $bShowHumanReadableDescription- [optional] a boolean value. Default is True.
;                  $iError              - [optional] an integer value. Default is @error.
;                  $iExtended           - [optional] an integer value. Default is @extended.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO Description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_GetErrorDescription($sDescription = '', $bShowHumanReadableDescription = True, $iError = @error, $iExtended = @extended)
	Local $sInfo = ''
	If $iError Then
		$sInfo = '! ADO ERROR  [ ' & $iError & ' / ' & $iExtended & ' ]  ' & $sDescription & '  '
		If $bShowHumanReadableDescription Then
			$sInfo &= @CRLF
			$sInfo &= '!    @ERROR=' & $iError & '='
			Switch $iError
				Case $ADO_ERR_SUCCESS
					$sInfo &= 'No Error'
				Case $ADO_ERR_GENERAL
					$sInfo &= 'General - some ADO Error - Not classified type of error'
				Case $ADO_ERR_COMERROR
					$sInfo &= 'COM Error - check your COM Error Handler'
				Case $ADO_ERR_COMHANDLER
					$sInfo &= 'COM Error Handler Registration'
				Case $ADO_ERR_CONNECTION
					$sInfo &= '$oConection.Open 	- Opening error'
				Case $ADO_ERR_ISNOTOBJECT
					$sInfo &= 'Function Parameters error - Expected/Required Object'
				Case $ADO_ERR_ISCLOSEDOBJECT
					$sInfo &= 'Object state error - Expected/Required state is $ADO_adStateOpen - is $ADO_adStateClosed'
				Case $ADO_ERR_ISNOTREADYOBJECT
					$sInfo &= 'Object state error - Expected/Required state is $ADO_adStateOpen - is $ADO_adStateConnecting or $ADO_adStateExecuting or $ADO_adStateFetching'
				Case $ADO_ERR_INVALIDOBJECTTYPE
					$sInfo &= 'Function Parameters error - Expected/Required different Object Type'
				Case $ADO_ERR_INVALIDPARAMETERTYPE
					$sInfo &= 'Function Parameters error - Invalid Variable type passed to the function'
				Case $ADO_ERR_INVALIDPARAMETERVALUE
					$sInfo &= 'Function Parameters error - Invalid value passed to the function'
				Case $ADO_ERR_INVALIDARRAY
					$sInfo &= 'Function Parameters error - Invalid Recordset Array'
				Case $ADO_ERR_RECORDSETEMPTY
					$sInfo &= 'The Recordset is Empty'
				Case $ADO_ERR_NOCURRENTRECORD
					$sInfo &= 'The Recordset has no current record'
				Case $ADO_ERR_ENUMCOUNTER
					$sInfo &= 'not used in UDF - just for other/future testing'
				Case Else
					$sInfo &= 'UNKNOWN @ERROR'
			EndSwitch

			$sInfo &= @CRLF & '    @EXTENDED=' & $iExtended & '='
			Switch $iExtended
				Case $ADO_EXT_DEFAULT
					$sInfo &= 'default Extended Value'
				Case $ADO_EXT_PARAM1
					$sInfo &= 'Error Occurs in 1-Parameter'
				Case $ADO_EXT_PARAM2
					$sInfo &= 'Error Occurs in 2-Parameter'
				Case $ADO_EXT_PARAM3
					$sInfo &= 'Error Occurs in 3-Parameter'
				Case $ADO_EXT_PARAM4
					$sInfo &= 'Error Occurs in 4-Parameter'
				Case $ADO_EXT_PARAM5
					$sInfo &= 'Error Occurs in 5-Parameter'
				Case $ADO_EXT_PARAM6
					$sInfo &= 'Error Occurs in 6-Parameter'
				Case $ADO_EXT_INTERNALFUNCTION
					$sInfo &= 'Error Related to internal Function - should not happend - UDF Developer make something wrong ???'
				Case $ADO_EXT_ENUMCOUNTER
					$sInfo &= 'not used in UDF - just for other/future testing'
				Case Else
					$sInfo &= 'UNKNOWN @EXTENDED'
			EndSwitch
		EndIf
	EndIf

	Return SetError($iError, $iExtended, $sInfo)
EndFunc   ;==>_ADO_GetErrorDescription

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_MSDNErrorValueEnum_Description
; Description ...: change ErrorValueEnum to Human Readable description
; Syntax ........: _ADO_MSDNErrorValueEnum_Description($iError[, $iErrorMacro = @error[, $iExtendedMacro = @extended]])
; Parameters ....: $iError              - an integer value. ErrorValueEnum
;                  $iErrorMacro         - [optional] an integer value. Default is @error.
;                  $iExtendedMacro      - [optional] an integer value. Default is @extended.
; Return values .: $sDescription
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681549(v=vs.85).aspx
; Example .......: No
; ===============================================================================================================================
Func _ADO_MSDNErrorValueEnum_Description($iError, $iErrorMacro = @error, $iExtendedMacro = @extended)
	Local $sDescription = ''
	If StringLeft($iError, 2) = '0x' Then
		$iError = Number(Dec(StringRight($iError, 4)))
	EndIf

	Switch $iError
		Case $ADO_adErrProviderFailed
			$sDescription = "Provider failed to perform the requested operation."
		Case $ADO_adErrInvalidArgument
			$sDescription = "Arguments are of the wrong type, are out of acceptable range, or are in conflict with one another. This error is often caused by a typographical error in an SQL SELECT statement. For example, a misspelled field name or table name can generate this error. This error can also occur when a field or table named in a SELECT statement does not exist in the data store."
		Case $ADO_adErrOpeningFile
			$sDescription = "File could not be opened. A misspelled file name was specified, or a file has been moved, renamed, or deleted. Over a network, the drive might be temporarily unavailable or network traffic might be preventing a connection."
		Case $ADO_adErrReadFile
			$sDescription = "File could not be read. The name of the file is specified incorrectly, the file might have been moved or deleted, or the file might have become corrupted."
		Case $ADO_adErrWriteFile
			$sDescription = "Write to file failed. You might have closed a file and then tried to write to it, or the file might be corrupted. If the file is located on a network drive, transient network conditions might prevent writing to a network drive."
		Case $ADO_adErrIllegalOperation
			$sDescription = "Operation is not allowed in this context."
		Case $ADO_adErrNoCurrentRecord
			$sDescription = "Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record."
		Case $ADO_adErrCantChangeProvider
			$sDescription = "Supplied provider is different from the one already in use."
		Case $ADO_adErrInTransaction
			$sDescription = "Connection object cannot be explicitly closed while in a transaction. A Recordset or Connection object that is currently participating in a transaction cannot be closed. Call either RollbackTrans or CommitTrans before closing the object."
		Case $ADO_adErrFeatureNotAvailable
			$sDescription = "The object or provider is not capable of performing the requested operation. Some operations depend on a particular provider version."
		Case $ADO_adErrItemNotFound
			$sDescription = "Item cannot be found in the collection corresponding to the requested name or ordinal. An incorrect field or table name has been specified."
		Case $ADO_adErrObjectInCollection
			$sDescription = "Object is already in collection. Cannot append. An object cannot be added to the same collection twice."
		Case $ADO_adErrObjectNotSet
			$sDescription = "Object is no longer valid."
		Case $ADO_adErrDataConversion
			$sDescription = "Application uses a value of the wrong type for the current operation. You might have supplied a string to an operation that expects a stream, for example."
		Case $ADO_adErrObjectClosed
			$sDescription = "Operation is not allowed when the object is closed. TheConnection or Recordset has been closed. For example, some other routine might have closed a global object. You can prevent this error by checking the State property before you attempt an operation."
		Case $ADO_adErrObjectOpen
			$sDescription = "Operation is not allowed when the object is open. An object that is open cannot be opened. Fields cannot be appended to an open Recordset."
		Case $ADO_adErrProviderNotFound
			$sDescription = "Provider cannot be found. It may not be properly installed."
		Case $ADO_adErrBoundToCommand
			$sDescription = "The ActiveConnection property of a Recordset object, which has a Command object as its source, cannot be changed. The application attempted to assign a newConnection object to a Recordset that has a Commandobject as its source."
		Case $ADO_adErrInvalidParamInfo
			$sDescription = "Parameter object is improperly defined. Inconsistent or incomplete information was provided."
		Case $ADO_adErrInvalidConnection
			$sDescription = "The connection cannot be used to perform this operation. It is either closed or invalid in this context."
		Case $ADO_adErrNotReentrant
			$sDescription = "Operation cannot be performed while processing event. An operation cannot be performed within an event handler that causes the event to fire again. For example, navigation methods should not be called from within aWillMove event handler."
		Case $ADO_adErrStillExecuting
			$sDescription = "Operation cannot be performed while executing asynchronously."
		Case $ADO_adErrOperationCancelled
			$sDescription = "Operation has been canceled by the user. The application has called the CancelUpdate or CancelBatch method and the current operation has been canceled."
		Case $ADO_adErrStillConnecting
			$sDescription = "Operation cannot be performed while connecting asynchronously."
		Case $ADO_adErrInvalidTransaction
			$sDescription = "Coordinating transaction is invalid or has not started."
		Case $ADO_adErrNotExecuting
			$sDescription = "Operation cannot be performed while not executing."
		Case $ADO_adErrUnsafeOperation
			$sDescription = "Safety settings on this computer prohibit accessing a data source on another domain."
		Case $ADO_adWrnSecurityDialog
			$sDescription = "For internal use only. Don't use. (Entry was included for the sake of completeness. This error should not appear in your code.)"
		Case $ADO_adWrnSecurityDialogHeader
			$sDescription = "For internal use only. Don't use. (Entry included for the sake of completeness. This error should not appear in your code.)"
		Case $ADO_adErrIntegrityViolation
			$sDescription = "Data value conflicts with the integrity constraints of the field. A new value for a Field would cause a duplicate key. A value that forms one side of a relationship between two records might not be updatable."
		Case $ADO_adErrPermissionDenied
			$sDescription = "Insufficient permission prevents writing to the field. The user named in the connection string does not have the proper permissions to write to a Field."
		Case $ADO_adErrDataOverflow
			$sDescription = "Data value is too large to be represented by the field data type. A numeric value that is too large for the intended field was assigned. For example, a long integer value was assigned to a short integer field."
		Case $ADO_adErrSchemaViolation
			$sDescription = "Data value conflicts with the data type or constraints of the field. The data store has validation constraints that differ from the Field value."
		Case $ADO_adErrSignMismatch
			$sDescription = "Conversion failed because the data value was signed and the field data type used by the provider was unsigned."
		Case $ADO_adErrCantConvertvalue
			$sDescription = "Data value cannot be converted for reasons other than sign mismatch or data overflow. For example, conversion would have truncated data."
		Case $ADO_adErrCantCreate
			$sDescription = "Data value cannot be set or retrieved because the field data type was unknown, or the provider had insufficient resources to perform the operation."
		Case $ADO_adErrColumnNotOnThisRow
			$sDescription = "Record does not contain this field. An incorrect field name was specified or a field not in the Fields collection of the current record was referenced."
		Case $ADO_adErrURLDoesNotExist
			$sDescription = "Either the source URL or the parent of the destination URL does not exist. There is a typographical error in either the source or destination URL. You might havehttp://mysite/photo/myphoto.jpg when you should actually have http://mysite/photos/myphoto.jpginstead. The typographical error in the parent URL (in this case, photo instead of photos) has caused the error."
		Case $ADO_adErrTreePermissionDenied
			$sDescription = "Permissions are insufficient to access tree or subtree. The user named in the connection string does not have the appropriate permissions."
		Case $ADO_adErrInvalidURL
			$sDescription = "URL contains invalid characters. Make sure the URL is typed correctly. The URL follows the scheme registered to the current provider (for example, Internet Publishing Provider is registered for http)."
		Case $ADO_adErrResourceLocked
			$sDescription = "Object represented by the specified URL is locked by one or more other processes. Wait until the process has finished and attempt the operation again. The object you are trying to access has been locked by another user or by another process in your application. This is most likely to arise in a multi-user environment."
		Case $ADO_adErrResourceExists
			$sDescription = "Copy operation cannot be performed. Object named by destination URL already exists. Specify adCopyOverwriteto replace the object. If you do not specifyadCopyOverwrite when copying the files in a directory, the copy fails when you try to copy an item that already exists in the destination location."
		Case $ADO_adErrCannotComplete
			$sDescription = "The server cannot complete the operation. This might be because the server is busy with other operations or it might be low on resources."
		Case $ADO_adErrVolumeNotFound
			$sDescription = "Provider cannot locate the storage device indicated by the URL. Make sure the URL is typed correctly. The URL of the storage device might be incorrect, but this error can occur for other reasons. The device might be offline or a large volume of network traffic might prevent the connection from being made."
		Case $ADO_adErrOutOfSpace
			$sDescription = "Operation cannot be performed. Provider cannot obtain enough storage space. There might not be enough RAM or hard-drive space for temporary files on the server."
		Case $ADO_adErrResourceOutOfScope
			$sDescription = "Source or destination URL is outside the scope of the current record."
		Case $ADO_adErrUnavailable
			$sDescription = "Operation failed to complete and the status is unavailable. The field may be unavailable or the operation was not attempted. Another user might have changed or deleted the field you are trying to access."
		Case $ADO_adErrURLNamedRowDoesNotExist
			$sDescription = "Record named by this URL does not exist. While attempting to open a file using a Record object, either the file name or the path to the file was misspelled."
		Case $ADO_adErrDelResOutOfScope
			$sDescription = "The URL of the object to be deleted is outside the scope of the current record."
		Case $ADO_adErrCatalogNotSet
			$sDescription = "Operation requires a valid ParentCatalog."
		Case $ADO_adErrCantChangeConnection
			$sDescription = "Connection was denied. The new connection you requested has different characteristics than the one already in use."
		Case $ADO_adErrFieldsUpdateFailed
			$sDescription = "Fields update failed. For further information, examine theStatus property of individual field objects. This error can occur in two situations: when changing a Field object's value in the process of changing or adding a record to the database; and when changing the properties of the Fieldobject itself."
		Case $ADO_adErrDenyNotSupported
			$sDescription = "Provider does not support sharing restrictions. An attempt was made to restrict file sharing and your provider does not support the concept."
		Case $ADO_adErrDenyTypeNotSupported
			$sDescription = "Provider does not support the requested kind of sharing restriction. An attempt was made to establish a particular type of file-sharing restriction that is not supported by your provider. See the provider's documentation to determine what file-sharing restrictions are supported."
	EndSwitch
	Return SetError($iErrorMacro, $iExtendedMacro, '[ ' & $iErrorMacro & ' / ' & $iExtendedMacro & ' ] ' & $sDescription)
EndFunc   ;==>_ADO_MSDNErrorValueEnum_Description

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_UDFVersion
; Description ...: Get UDFVersion number
; Syntax ........: _ADO_UDFVersion()
; Parameters ....: none
; Return values .: UDF Version
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_UDFVersion()
	Return '2.1.13 BETA'
EndFunc   ;==>_ADO_UDFVersion

; #FUNCTION# ====================================================================================================================
; Name ..........: _Au3Date_to_SQLDate
; Description ...:
; Syntax ........: _Au3Date_to_SQLDate($sAu3Date)
; Parameters ....: $sAu3Date            - a string value.
; Return values .: None
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _Au3Date_to_SQLDate($sAu3Date)
	; IN:  1970/01/01 12:30:15
	; OUT: 1970-01-01T12:30:15.000

	If Not _DateIsValid($sAu3Date) Then
		Return SetError($ADO_ERR_GENERAL, $ADO_EXT_PARAM1, $ADO_RET_FAILURE)
	EndIf

	; if only date then add time
	If StringRegExpReplace($sAu3Date, '(\d{4}\/\d{2}\/\d{2})', '') = '' Then $sAu3Date &= ' 00:00:00'
	; replace "/" to "-"    and add miliseconds
	Local $sSQLDate = StringReplace($sAu3Date, '/', '-') & '.000'
	; change the space (separator for date and time) for SQL equivalent T char
	$sSQLDate = StringReplace($sSQLDate, ' ', 'T')

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $sSQLDate)
EndFunc   ;==>_Au3Date_to_SQLDate

; #FUNCTION# ====================================================================================================================
; Name ..........: _SQLDate_to_Au3Date
; Description ...:
; Syntax ........: _SQLDate_to_Au3Date($sDate[, $bOnlyYMD = False])
; Parameters ....: $sDate               - a string value.
;                  $bOnlyYMD            - [optional] a boolean value. Default is False.
; Return values .: Au3Date
; Author ........: mLipok
; Modified ......:
; Remarks .......: TODO - description
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _SQLDate_to_Au3Date($sDate, $bOnlyYMD = False)
	Local $sParam = ($bOnlyYMD = True) ? '$1\/$2\/$3' : '$1\/$2\/$3\ $4:$5:$6'
	Return StringRegExpReplace($sDate, '(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})', $sParam)

EndFunc   ;==>_SQLDate_to_Au3Date
#EndRegion ADO.au3 - Functions - MISC

#Region ADO.au3 - Functions - OpenSchema

Func _ADO_OpenSchema_Catalogs(ByRef $oConnection, $s_CATALOG_NAME = '')
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aCriteria_Catalog[1]
	If IsString($s_CATALOG_NAME) And $s_CATALOG_NAME <> '' Then $aCriteria_Catalog[0] = $s_CATALOG_NAME

	Local $oRecordset_catalogs = $oConnection.OpenSchema($ADO_adSchemaCatalogs, $aCriteria_Catalog)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return $oRecordset_catalogs

EndFunc   ;==>_ADO_OpenSchema_Catalogs

Func _ADO_OpenSchema_Columns(ByRef $oConnection, $s_TABLE_CATALOG = '', $s_TABLE_SCHEMA = '', $s_TABLE_NAME = '', $s_COLUMN_NAME = '')
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aCriteria_Column[4]
	If IsString($s_TABLE_CATALOG) And $s_TABLE_CATALOG <> '' Then $aCriteria_Column[0] = $s_TABLE_CATALOG
	If IsString($s_TABLE_SCHEMA) And $s_TABLE_SCHEMA <> '' Then $aCriteria_Column[1] = $s_TABLE_SCHEMA
	If IsString($s_TABLE_NAME) And $s_TABLE_NAME <> '' Then $aCriteria_Column[2] = $s_TABLE_NAME
	If IsString($s_COLUMN_NAME) And $s_COLUMN_NAME <> '' Then $aCriteria_Column[3] = $s_COLUMN_NAME

	Local $oRecordset_columns = $oConnection.OpenSchema($ADO_adSchemaColumns, $aCriteria_Column)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return $oRecordset_columns

EndFunc   ;==>_ADO_OpenSchema_Columns

Func _ADO_OpenSchema_Indexes(ByRef $oConnection, $s_TABLE_CATALOG = '', $s_TABLE_SCHEMA = '', $s_INDEX_NAME = '', $s_TYPE = '', $s_TABLE_NAME = '')
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aCriteria_Index[5]
	If IsString($s_TABLE_CATALOG) And $s_TABLE_CATALOG <> '' Then $aCriteria_Index[0] = $s_TABLE_CATALOG
	If IsString($s_TABLE_SCHEMA) And $s_TABLE_SCHEMA <> '' Then $aCriteria_Index[1] = $s_TABLE_SCHEMA
	If IsString($s_INDEX_NAME) And $s_INDEX_NAME <> '' Then $aCriteria_Index[2] = $s_INDEX_NAME
	If IsString($s_TYPE) And $s_TYPE <> '' Then $aCriteria_Index[3] = $s_TYPE
	If IsString($s_TABLE_NAME) And $s_TABLE_NAME <> '' Then $aCriteria_Index[4] = $s_TABLE_NAME

	Local $oRecordset_Indexes = $oConnection.OpenSchema($ADO_adSchemaIndexes, $aCriteria_Index)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return $oRecordset_Indexes

EndFunc   ;==>_ADO_OpenSchema_Indexes

Func _ADO_OpenSchema_Tables(ByRef $oConnection, $s_TABLE_CATALOG = '', $s_TABLE_SCHEMA = '', $s_TABLE_NAME = '', $s_TABLE_TYPE = '')
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aCriteria_Table[4]
	If IsString($s_TABLE_CATALOG) And $s_TABLE_CATALOG <> '' Then $aCriteria_Table[0] = $s_TABLE_CATALOG
	If IsString($s_TABLE_SCHEMA) And $s_TABLE_SCHEMA <> '' Then $aCriteria_Table[1] = $s_TABLE_SCHEMA
	If IsString($s_TABLE_NAME) And $s_TABLE_NAME <> '' Then $aCriteria_Table[2] = $s_TABLE_NAME
	If IsString($s_TABLE_TYPE) And $s_TABLE_TYPE <> '' Then $aCriteria_Table[3] = $s_TABLE_TYPE

	Local $oRecordset_tables = $oConnection.OpenSchema($ADO_adSchemaTables, $aCriteria_Table)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return $oRecordset_tables

EndFunc   ;==>_ADO_OpenSchema_Tables

Func _ADO_OpenSchema_Views(ByRef $oConnection, $s_TABLE_CATALOG = '', $s_TABLE_SCHEMA = '', $s_TABLE_NAME = '')
	; Error handler, automatic cleanup at end of function
	Local $oADO_COM_ErrorHandler = ObjEvent("AutoIt.Error", __ADO_ComErrorHandler_InternalFunction)
	If @error Then Return SetError($ADO_ERR_COMHANDLER, @error, $ADO_RET_FAILURE)
	#forceref $oADO_COM_ErrorHandler

	__ADO_Connection_IsReady($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aCriteria_View[3]
	If IsString($s_TABLE_CATALOG) And $s_TABLE_CATALOG <> '' Then $aCriteria_View[0] = $s_TABLE_CATALOG
	If IsString($s_TABLE_SCHEMA) And $s_TABLE_SCHEMA <> '' Then $aCriteria_View[1] = $s_TABLE_SCHEMA
	If IsString($s_TABLE_NAME) And $s_TABLE_NAME <> '' Then $aCriteria_View[2] = $s_TABLE_NAME

	Local $oRecordset_Views = $oConnection.OpenSchema($ADO_adSchemaViews, $aCriteria_View)
	If @error Then Return SetError($ADO_ERR_COMERROR, @error, $ADO_RET_FAILURE)

	Return $oRecordset_Views

EndFunc   ;==>_ADO_OpenSchema_Views

Func _ADO_Schema_GetAllCatalogs(ByRef $oConnection)
	Local $oRecordset_catalogs = _ADO_OpenSchema_Catalogs($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aSchema_Catalogs = _ADO_Recordset_ToArray($oRecordset_catalogs)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oRecordset_catalogs.Close
	$oRecordset_catalogs = Null

	Return $aSchema_Catalogs
EndFunc   ;==>_ADO_Schema_GetAllCatalogs

Func _ADO_Schema_GetAllTables(ByRef $oConnection, $s_TABLE_CATALOG)
	Local $oRecordset_tables = _ADO_OpenSchema_Tables($oConnection)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aSchema_Tables = _ADO_Recordset_ToArray($oRecordset_tables, $s_TABLE_CATALOG)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oRecordset_tables.Close
	$oRecordset_tables = Null

	Return $aSchema_Tables
EndFunc   ;==>_ADO_Schema_GetAllTables

Func _ADO_Schema_GetAllViews(ByRef $oConnection, $s_TABLE_CATALOG)
	Local $oRecordset_Views = _ADO_OpenSchema_Views($oConnection, $s_TABLE_CATALOG)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	Local $aSchema_Views = _ADO_Recordset_ToArray($oRecordset_Views)
	If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)

	$oRecordset_Views.Close
	$oRecordset_Views = Null

	Return $aSchema_Views
EndFunc   ;==>_ADO_Schema_GetAllViews
#EndRegion ADO.au3 - Functions - OpenSchema

#Region ADO.au3 - Functions - Connection Strings

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_ConnectionString_Access
; Description ...: Create Connection string for MS Access file
; Syntax ........: _ADO_ConnectionString_Access($sFileFullPath[, $sUser = Default[, $sPassword = Default[, $sDriver = Default]]])
; Parameters ....: $sFileFullPath   - a string value.
;                  $sUser               - [optional] a string value. Default is Default.
;                  $sPassword           - [optional] a string value. Default is Default.
;                  $sDriver             - [optional] a string value. Default is Default.
; Return values .: connection string
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_ConnectionString_Access($sFileFullPath, $sUser = Default, $sPassword = Default, $sDriver = Default)

	If $sUser = Default Then
		$sUser = ''
	Else
		$sUser = 'Uid=' & $sUser & ';'
	EndIf

	If $sPassword = Default Then
		$sPassword = ''
	Else
		$sPassword = 'PWD=' & $sPassword & ';'
	EndIf

	If $sDriver = Default Then $sDriver = 'Microsoft Access Driver (*.mdb)'

	Local $sConnectionString = 'Driver={' & $sDriver & '};Dbq="' & $sFileFullPath & '";' & $sUser & $sPassword

	If Not StringRegExp($sConnectionString, '(?i)(Microsoft Access Driver \(*.mdb\)|Microsoft Access Driver \(*.mdb, *.accdb\))', $STR_REGEXPMATCH) Then
		$sConnectionString = StringReplace($sConnectionString, ';Dbq=', ' ;Data Source=')
	EndIf

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $sConnectionString)
EndFunc   ;==>_ADO_ConnectionString_Access

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_ConnectionString_Excel
; Description ...: Create Connection string for MS Excel file
; Syntax ........: _ADO_ConnectionString_Excel([$sFileFullPath = Default[, $sProvider = Default[, $sExtProperties = Default[,
;                  $HDR = Default[, $IMEX = 0]]]]])
; Parameters ....: $sFileFullPath   - [optional] a string value. Default is Default.
;                  $sProvider           - [optional] a string value. Default is Default.
;                  $sExtProperties        - [optional] a string value. Default is Default.
;                  $HDR                 - [optional] an unknown value. Default is Default.
;                  $IMEX                - [optional] an unknown value. Default is 0.
; Return values .: Connection String
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_ConnectionString_Excel($sFileFullPath = Default, $sProvider = Default, $sExtProperties = Default, $HDR = Default, $IMEX = Default)

	; Parameter #1 Validation
	If $sFileFullPath = Default Then
		$sFileFullPath = FileOpenDialog('Select XLS File', @ScriptDir, 'XLS file (*.xls)', $FD_FILEMUSTEXIST + $FD_PATHMUSTEXIST)
		If @error Then Return SetError(@error, @extended, $ADO_RET_FAILURE)
	EndIf

	; Parameter #2 Validation
	If $sProvider = Default Then $sProvider = 'Microsoft.Jet.OLEDB.4.0'

	; Parameter #3 Validation
	If $sExtProperties = Default Then $sExtProperties = 'Excel 8.0'

	; Parameter #4 Validation
	If $HDR = Default Or $HDR = True Or $HDR = 'yes' Then
		$HDR = 'yes'
	Else
		$HDR = 'no'
	EndIf

	; Parameter #5 Validation
	If $IMEX = Default Then $IMEX = 0

	Local $sXLS_ConnectionString = _
			'Provider=' & $sProvider & ';' & _
			'Data Source="' & $sFileFullPath & '";' & _
			'Extended Properties="' & $sExtProperties & ';' & _
			'HDR=' & $HDR & ';' & _
			'IMEX=' & $IMEX & '";'

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $sXLS_ConnectionString)
EndFunc   ;==>_ADO_ConnectionString_Excel

; #FUNCTION# ====================================================================================================================
; Name ..........: _ADO_ConnectionString_MySQL
; Description ...: Create Connection string for MySQL database
; Syntax ........: _ADO_ConnectionString_MySQL($sUser, $sPassword, $sDatabase[, $sDriver = Default [, $sServer = Default [,
;                  $sPort = Default]]])
; Parameters ....: $sUser               - a string value.
;                  $sPassword           - a string value.
;                  $sDatabase           - a string value.
;                  $sDriver             - [optional] a string value. Default is Default .
;                  $sServer             - [optional] a string value. Default is Default .
;                  $sPort               - [optional] a string value. Default is Default.
; Return values .: Connection String
; Author ........: mLipok
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........:
; Example .......: No
; ===============================================================================================================================
Func _ADO_ConnectionString_MySQL($sUser, $sPassword, $sDataBase, $sDriver = Default, $sServer = Default, $sPort = Default)
	; https://dev.mysql.com/doc/connector-net/en/connector-net-connection-options.html

	If $sDriver = Default Then $sDriver = 'MySQL ODBC 5.3 ANSI Driver'
	If $sServer = Default Then $sServer = 'localhost'
	If $sPort = Default Then $sPort = '3306'

	Local $sConnectionString = 'Driver={' & $sDriver & '};SERVER=' & $sServer & ';PORT=' & $sPort & ';DATABASE=' & $sDataBase & ';User=' & $sUser & ';Password=' & $sPassword & ';'

	Return SetError($ADO_ERR_SUCCESS, $ADO_EXT_DEFAULT, $sConnectionString)
EndFunc   ;==>_ADO_ConnectionString_MySQL
#EndRegion ADO.au3 - Functions - Connection Strings

#Region ADO.au3 - TODO and Help/Docs

#CS
	SQLState Property
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms681570(v=vs.85).aspx

	NativeError Property (ADO)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms678049(v=vs.85).aspx

	Programming ADO SQL Server Applications
	https://technet.microsoft.com/en-us/library/aa905875(v=sql.80).aspx
	https://technet.microsoft.com/en-us/library/aa214053(v=sql.80).aspx

	ADO API Reference
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms678086(v=vs.85).aspx

	ADO Code Examples
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms681484(v=vs.85).aspx

	Microsoft OLE DB Provider for SQL Server
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms677227(v=vs.85).aspx

	OpenSchema Method Example (VB)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms675853(v=vs.85).aspx

	Errors Collection Properties, Methods, and Events
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms676176(v=vs.85).aspx

	ErrorValueEnum
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms677004(v=vs.85).aspx

	ADO Code Examples VBScript
	https://msdn.microsoft.com/en-us/library/ms676589(v=vs.85).aspx

	ADO Code Examples in Visual Basic
	https://msdn.microsoft.com/en-us/library/ms675104(v=vs.85).aspx
#CE

#CS ADO Events some reference

	Handling ADO Events
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms681467(v=vs.85).aspx

	ADO Event Handler Summary
	https://msdn.microsoft.com/en-us/library/ms677579(v=vs.85).aspx

	Handling Errors and Messages in ADO
	https://technet.microsoft.com/en-us/library/aa905919(v=sql.80).aspx

	ExecuteComplete Event (ADO)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms676183(v=vs.85).aspx


	ADO Error Reference
	https://msdn.microsoft.com/en-us/library/ms681549(v=vs.85).aspx

	ADO Collections
	https://msdn.microsoft.com/en-us/library/ms677591(v=vs.85).aspx

	WillChangeRecordset and RecordsetChangeComplete Events (ADO)
	https://msdn.microsoft.com/en-us/library/ms680919(v=vs.85).aspx


	Handling Errors and Messages in ADO
	https://technet.microsoft.com/en-us/library/aa905919(v=sql.80).aspx

	Performing Transactions in ADO
	https://technet.microsoft.com/en-us/library/aa905921(v=sql.80).aspx

	An ADO Transaction
	https://msdn.microsoft.com/en-us/library/aa227162(v=vs.60).aspx

	ADO BeginTrans, CommitTrans, and RollbackTrans Methods
	http://www.w3schools.com/asp/met_conn_begintrans.asp

	BeginTrans, CommitTrans, and RollbackTrans Methods Example (VB)
	https://msdn.microsoft.com/en-us/library/windows/desktop/ms677538%28v=vs.85%29.aspx

#CE

#CS
	View Object (ADOX)
	https://msdn.microsoft.com/en-us/library/ms676503(v=vs.85).aspx

	Views Collection (ADOX)
	https://msdn.microsoft.com/en-us/library/ms677523(v=vs.85).aspx

	Views Collection, CommandText Property Example (VB)
	https://msdn.microsoft.com/en-us/library/ms677503(v=vs.85).aspx

	Views and Fields Collections Example (VB)
	https://msdn.microsoft.com/en-us/library/ms680939(v=vs.85).aspx


	How To Determine Number of Records Affected by an ADO UPDATE
	https://support.microsoft.com/en-us/kb/195048

	ADO Programmer's Guide
	https://msdn.microsoft.com/en-us/library/ms681025(v=vs.85).aspx

	ADO Programmer's Reference
	https://msdn.microsoft.com/en-us/library/ms676539(v=vs.85).aspx

	ADO Objects and Interfaces
	https://msdn.microsoft.com/en-us/library/ms679836(v=vs.85).aspx

	ADOX Programming Code Examples
	http://allenbrowne.com/func-adox.html

	ADO Programming Code Examples
	http://allenbrowne.com/func-ADO.html

	Driver Specification Subkeys
	https://msdn.microsoft.com/en-us/library/ms714538(v=vs.85).aspx
	Local $key = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"

	The SQL Server Native Client...
	https://msdn.microsoft.com/pl-pl/sqlserver/aa937733.aspx

	Get Started Developing with the SQL Server Native Client
	https://msdn.microsoft.com/pl-pl/sqlserver/ff658533

	Building Applications with SQL Server Native Client
	https://msdn.microsoft.com/en-us/library/ms130904.aspx

	When to Use SQL Server Native Client
	https://msdn.microsoft.com/en-us/library/ms130828.aspx

	What's New in SQL Server Native Client
	https://msdn.microsoft.com/en-us/library/cc280510.aspx

	SQL Server Native Client Features
	https://msdn.microsoft.com/en-us/library/ms131456.aspx

	SQL Server Native Client Programming
	https://msdn.microsoft.com/en-us/library/ms130892.aspx

	Native API for SQL Server FAQ
	https://msdn.microsoft.com/en-us/sqlserver/aa937707.aspx

	Examlple of Connection Strings
	https://www.connectionstrings.com/


	HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers

	ADO Command Strategies
	https://msdn.microsoft.com/en-us/library/aa260835(v=vs.60).aspx

	Command Object (ADO)
	https://msdn.microsoft.com/en-us/library/ms677502(v=vs.85).aspx

	CreateParameter Method (ADO)
	https://msdn.microsoft.com/en-us/library/ms677209(v=vs.85).aspx


	; How To Determine Number of Records Affected by an ADO UPDATE
	; https://support.microsoft.com/en-us/kb/195048
	; Use the command object to perform an UPDATE and return the count of affected records.

	XSLT Transformations (Recordset XML >> HTML)
	https://msdn.microsoft.com/en-us/library/ms675135(v=vs.85).aspx

	XML Recordset Persistence Scenario
	https://msdn.microsoft.com/en-us/library/ms675780(v=vs.85).aspx

	Persisting Data
	https://msdn.microsoft.com/en-us/library/ms675273(v=vs.85).aspx

	Saving to the XML DOM Object
	https://msdn.microsoft.com/en-us/library/ms675954(v=vs.85).aspx

	Persisting Records in XML Format
	https://msdn.microsoft.com/en-us/library/ms681538(v=vs.85).aspx




#CE
#EndRegion ADO.au3 - TODO and Help/Docs

#Region ADO.au3 - NEW WIP

#EndRegion ADO.au3 - NEW WIP
