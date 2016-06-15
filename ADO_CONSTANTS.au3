#CS
	Any program that accepts data from a user must include code to validate that data before sending it to the data store. You cannot rely on the data store, the provider, ADO, or even your programming language to notify you of problems. You must check every byte entered by your users, making sure that data is the correct type for its field and that required fields are not empty.
	https://msdn.microsoft.com/en-us/library/ms681470(v=vs.85).aspx
#CE

#include-once
#AutoIt3Wrapper_Au3Check_Parameters=-d -w 1 -w 2 -w 3 -w- 4 -w 5 -w 6 -w 7

#Region ADO_CONSTANTS.au3 - ADO.au3 UDF Constants
Global Enum _
		$ADO_ERR_SUCCESS, _ ;			 	No Error
		$ADO_ERR_GENERAL, _ ;   			General - some ADO Error - Not classified type of error
		$ADO_ERR_COMERROR, _ ;   			COM Error - check your COM Error Handler
		$ADO_ERR_COMHANDLER, _ ;   			COM Error Handler Registration
		$ADO_ERR_CONNECTION, _ ;   			$oConection.Open 	- Opening error
		$ADO_ERR_ISNOTOBJECT, _ ;			Function Parameters error - Expected/Required Object
		$ADO_ERR_ISCLOSEDOBJECT, _ ;		Object state error - Expected/Required state is $ADO_adStateOpen - is $ADO_adStateClosed
		$ADO_ERR_ISNOTREADYOBJECT, _ ;		Object state error - Expected/Required state is $ADO_adStateOpen - is $ADO_adStateConnecting or $ADO_adStateExecuting or $ADO_adStateFetching
		$ADO_ERR_INVALIDOBJECTTYPE, _ ;		Function Parameters error - Expected/Required different Object Type
		$ADO_ERR_INVALIDPARAMETERTYPE, _ ;  Function Parameters error - Invalid Variable type passed to the function
		$ADO_ERR_INVALIDPARAMETERVALUE, _ ; Function Parameters error - Invalid value passed to the function
		$ADO_ERR_INVALIDARRAY, _ ;			Function Parameters error - Invalid Recordset Array
		$ADO_ERR_RECORDSETEMPTY, _ ;		The Recordset is Empty
		$ADO_ERR_NOCURRENTRECORD, _ ;		The Recordset has no current record
		$ADO_ERR_ENUMCOUNTER ;-------------	just for testing

Global Enum _
		$ADO_EXT_DEFAULT, _ ;				default Extended Value
		$ADO_EXT_PARAM1, _ ;				Error Occurs in 1-Parameter
		$ADO_EXT_PARAM2, _ ;				Error Occurs in 2-Parameter
		$ADO_EXT_PARAM3, _ ;				Error Occurs in 3-Parameter
		$ADO_EXT_PARAM4, _ ;				Error Occurs in 4-Parameter
		$ADO_EXT_PARAM5, _ ;				Error Occurs in 5-Parameter
		$ADO_EXT_PARAM6, _ ;				Error Occurs in 6-Parameter
		$ADO_EXT_INTERNALFUNCTION, _ ;		Error Related to internal Function - should not happend - UDF Developer make something wrong ???
		$ADO_EXT_ENUMCOUNTER ;-------------	just for testing

Global Enum _
		$ADO_RET_FAILURE = -1, _ ;			Failure result
		$ADO_RET_SUCCESS = 1, _ ;			Successful result
		$ADO_RET_ENUMCOUNTER ;-------------	just for testing

Global Enum _
		$ADO_RS_ARRAY_GUID, _ ;				Array GUID
		$ADO_RS_ARRAY_FIELDNAMES, _ ;		Array index for inner FileNames Array
		$ADO_RS_ARRAY_RSCONTENT, _ ;		Array index for inner Recordset Array
		$ADO_RS_ARRAY_ENUMCOUNTR ;--------- just for testing

Global Const $ADO_RS_GUID = '{2399DBEE-2450-462D-B102-9094A9EB5D02}'
#EndRegion  ADO_CONSTANTS.au3 - ADO.au3 UDF Constants

#Region ADO_CONSTANTS.au3 - MSDN Enumerated Constants
;~ ADO Enumerated Constants
;~ http://msdn.microsoft.com/en-us/library/windows/desktop/ms678353%28v=vs.85%29.aspx

; ADCPROP_ASYNCTHREADPRIORITY_ENUM
Global Const $ADO_adPriorityLowest = 1
Global Const $ADO_adPriorityBelowNormal = 2
Global Const $ADO_adPriorityNormal = 3
Global Const $ADO_adPriorityAboveNormal = 4
Global Const $ADO_adPriorityHighest = 5

; ADCPROP_AUTORECALC_ENUM
Global Const $ADO_adRecalcUpFront = 0
Global Const $ADO_adRecalcAlways = 1

; ADCPROP_UPDATECRITERIA_ENUM
Global Const $ADO_adCriteriaKey = 0
Global Const $ADO_adCriteriaAllCols = 1
Global Const $ADO_adCriteriaUpdCols = 2
Global Const $ADO_adCriteriaTimeStamp = 3

; ADCPROP_UPDATERESYNC_ENUM
Global Const $ADO_adResyncNone = 0
Global Const $ADO_adResyncAutoIncrement = 1
Global Const $ADO_adResyncConflicts = 2
Global Const $ADO_adResyncUpdates = 4
Global Const $ADO_adResyncInserts = 8
Global Const $ADO_adResyncAll = 15

; AffectEnum
Global Const $ADO_adAffectCurrent = 1
Global Const $ADO_adAffectGroup = 2
Global Const $ADO_adAffectAll = 3
Global Const $ADO_adAffectAllChapters = 4

; BookmarkEnum
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms676118(v=vs.85).aspx
Global Const $ADO_adBookmarkCurrent = 0
Global Const $ADO_adBookmarkFirst = 1
Global Const $ADO_adBookmarkLast = 2

; CommandTypeEnum
Global Const $ADO_adCmdUnspecified = -1
Global Const $ADO_adCmdText = 1
Global Const $ADO_adCmdTable = 2
Global Const $ADO_adCmdStoredProc = 4
Global Const $ADO_adCmdUnknown = 8
Global Const $ADO_adCmdFile = 256
Global Const $ADO_adCmdTableDirect = 512

; CompareEnum
Global Const $ADO_adCompareLessThan = 0
Global Const $ADO_adCompareEqual = 1
Global Const $ADO_adCompareGreaterThan = 2
Global Const $ADO_adCompareNotEqual = 3
Global Const $ADO_adCompareNotComparable = 4

; ConnectModeEnum
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms675792(v=vs.85).aspx
Global Const $ADO_adModeUnknown = 0
Global Const $ADO_adModeRead = 1
Global Const $ADO_adModeWrite = 2
Global Const $ADO_adModeReadWrite = 3
Global Const $ADO_adModeShareDenyRead = 4
Global Const $ADO_adModeShareDenyWrite = 8
Global Const $ADO_adModeShareExclusive = 12
Global Const $ADO_adModeShareDenyNone = 16
Global Const $ADO_adModeRecursive = 0x400000

; ConnectOptionEnum
Global Const $ADO_adConnectUnspecified = -1
Global Const $ADO_adAsyncConnect = 16

; ConnectPromptEnum
Global Const $ADO_adPromptAlways = 1
Global Const $ADO_adPromptComplete = 2
Global Const $ADO_adPromptCompleteRequired = 3
Global Const $ADO_adPromptNever = 4

; CopyRecordOptionsEnum
Global Const $ADO_adCopyUnspecified = -1
Global Const $ADO_adCopyOverWrite = 1
Global Const $ADO_adCopyNonRecursive = 2
Global Const $ADO_adCopyAllowEmulation = 4

; CursorLocationEnum
Global Const $ADO_adUseNone = 1
Global Const $ADO_adUseServer = 2
Global Const $ADO_adUseClient = 3

; CursorOptionEnum
Global Const $ADO_adAddNew = 0x1000400
Global Const $ADO_adApproxPosition = 0x4000
Global Const $ADO_adBookmark = 0x2000
Global Const $ADO_adDelete = 0x1000800
Global Const $ADO_adFind = 0x80000
Global Const $ADO_adHoldRecords = 0x100
Global Const $ADO_adIndex = 0x100000
Global Const $ADO_adMovePrevious = 0x200
Global Const $ADO_adNotify = 0x40000
Global Const $ADO_adResync = 0x20000
Global Const $ADO_adSeek = 0x200000
Global Const $ADO_adUpdate = 0x1008000
Global Const $ADO_adUpdateBatch = 0x10000

; CursorTypeEnum
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms681771(v=vs.85).aspx
Global Const $ADO_adOpenUnspecified = -1 ; Does not specify the type of cursor.
Global Const $ADO_adOpenForwardOnly = 0 ; Default. Uses a forward-only cursor. Identical to a static cursor, except that you can only scroll forward through records. This improves performance when you need to make only one pass through a Recordset.
Global Const $ADO_adOpenKeyset = 1 ; Uses a keyset cursor. Like a dynamic cursor, except that you can't see records that other users add, although records that other users delete are inaccessible from your Recordset. Data changes by other users are still visible.
Global Const $ADO_adOpenDynamic = 2 ; Uses a dynamic cursor. Additions, changes, and deletions by other users are visible, and all types of movement through the Recordset are allowed, except for bookmarks, if the provider doesn't support them.
Global Const $ADO_adOpenStatic = 3 ; Uses a static cursor, which is a static copy of a set of records that you can use to find data or generate reports. Additions, changes, or deletions by other users are not visible.

; DataTypeEnum
; https://msdn.microsoft.com/en-us/library/ms675318(v=vs.85).aspx
Global Const $ADO_adArray = 0x2000
Global Const $ADO_adBigInt = 20
Global Const $ADO_adBinary = 128
Global Const $ADO_adBoolean = 11
Global Const $ADO_adBSTR = 8
Global Const $ADO_adChapter = 136
Global Const $ADO_adChar = 129
Global Const $ADO_adCurrency = 6
Global Const $ADO_adDate = 7
Global Const $ADO_adDBDate = 133
Global Const $ADO_adDBTime = 134
Global Const $ADO_adDBTimeStamp = 135
Global Const $ADO_adDecimal = 14
Global Const $ADO_adDouble = 5
Global Const $ADO_adEmpty = 0
Global Const $ADO_adError = 10
Global Const $ADO_adFileTime = 64
Global Const $ADO_adGUID = 72
Global Const $ADO_adIDispatch = 9
Global Const $ADO_adInteger = 3
Global Const $ADO_adIUnknown = 13
Global Const $ADO_adLongVarBinary = 205
Global Const $ADO_adLongVarChar = 201
Global Const $ADO_adLongVarWChar = 203
Global Const $ADO_adNumeric = 131
Global Const $ADO_adPropVariant = 138
Global Const $ADO_adSingle = 4
Global Const $ADO_adSmallInt = 2
Global Const $ADO_adTinyInt = 16
Global Const $ADO_adUnsignedBigInt = 21
Global Const $ADO_adUnsignedInt = 19
Global Const $ADO_adUnsignedSmallInt = 18
Global Const $ADO_adUnsignedTinyInt = 17
Global Const $ADO_adUserDefined = 132
Global Const $ADO_adVarBinary = 204
Global Const $ADO_adVarChar = 200
Global Const $ADO_adVariant = 12
Global Const $ADO_adVarNumeric = 139
Global Const $ADO_adVarWChar = 202
Global Const $ADO_adWChar = 130

; EditModeEnum
; https://msdn.microsoft.com/en-us/library/ms675856(v=vs.85).aspx
Global Const $ADO_adEditNone = 0
Global Const $ADO_adEditInProgress = 1
Global Const $ADO_adEditAdd = 2
Global Const $ADO_adEditDelete = 4

; ErrorValueEnum ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms681549(v=vs.85).aspx

Global Const $ADO_adErrProviderFailed = 3000
Global Const $ADO_adErrInvalidArgument = 3001
Global Const $ADO_adErrOpeningFile = 3002
Global Const $ADO_adErrReadFile = 3003
Global Const $ADO_adErrWriteFile = 3004
Global Const $ADO_adErrNoCurrentRecord = 3021
Global Const $ADO_adErrIllegalOperation = 3219
Global Const $ADO_adErrCantChangeProvider = 3220
Global Const $ADO_adErrInTransaction = 3246
Global Const $ADO_adErrFeatureNotAvailable = 3251
Global Const $ADO_adErrItemNotFound = 3265
Global Const $ADO_adErrObjectInCollection = 3367
Global Const $ADO_adErrObjectNotSet = 3420
Global Const $ADO_adErrDataConversion = 3421
Global Const $ADO_adErrObjectClosed = 3704
Global Const $ADO_adErrObjectOpen = 3705
Global Const $ADO_adErrProviderNotFound = 3706
Global Const $ADO_adErrBoundToCommand = 3707
Global Const $ADO_adErrInvalidParamInfo = 3708
Global Const $ADO_adErrInvalidConnection = 3709
Global Const $ADO_adErrNotReentrant = 3710
Global Const $ADO_adErrStillExecuting = 3711
Global Const $ADO_adErrOperationCancelled = 3712
Global Const $ADO_adErrStillConnecting = 3713
Global Const $ADO_adErrInvalidTransaction = 3714
Global Const $ADO_adErrNotExecuting = 3715
Global Const $ADO_adErrUnsafeOperation = 3716
Global Const $ADO_adWrnSecurityDialog = 3717
Global Const $ADO_adWrnSecurityDialogHeader = 3718
Global Const $ADO_adErrIntegrityViolation = 3719
Global Const $ADO_adErrPermissionDenied = 3720
Global Const $ADO_adErrDataOverflow = 3721
Global Const $ADO_adErrSchemaViolation = 3722
Global Const $ADO_adErrSignMismatch = 3723
Global Const $ADO_adErrCantConvertvalue = 3724
Global Const $ADO_adErrCantCreate = 3725
Global Const $ADO_adErrColumnNotOnThisRow = 3726
Global Const $ADO_adErrURLDoesNotExist = 3727
Global Const $ADO_adErrTreePermissionDenied = 3728
Global Const $ADO_adErrInvalidURL = 3729
Global Const $ADO_adErrResourceLocked = 3730
Global Const $ADO_adErrResourceExists = 3731
Global Const $ADO_adErrCannotComplete = 3732
Global Const $ADO_adErrVolumeNotFound = 3733
Global Const $ADO_adErrOutOfSpace = 3734
Global Const $ADO_adErrResourceOutOfScope = 3735
Global Const $ADO_adErrUnavailable = 3736
Global Const $ADO_adErrURLNamedRowDoesNotExist = 3737
Global Const $ADO_adErrDelResOutOfScope = 3738
Global Const $ADO_adErrCatalogNotSet = 3747
Global Const $ADO_adErrCantChangeConnection = 3748
Global Const $ADO_adErrFieldsUpdateFailed = 3749
Global Const $ADO_adErrDenyNotSupported = 3750
Global Const $ADO_adErrDenyTypeNotSupported = 3751

; EventReasonEnum
; https://msdn.microsoft.com/en-us/library/ms676681(v=vs.85).aspx
Global Const $ADO_adRsnAddNew = 1
Global Const $ADO_adRsnDelete = 2
Global Const $ADO_adRsnUpdate = 3
Global Const $ADO_adRsnUndoUpdate = 4
Global Const $ADO_adRsnUndoAddNew = 5
Global Const $ADO_adRsnUndoDelete = 6
Global Const $ADO_adRsnRequery = 7
Global Const $ADO_adRsnResynch = 8
Global Const $ADO_adRsnClose = 9
Global Const $ADO_adRsnMove = 10
Global Const $ADO_adRsnFirstChange = 11
Global Const $ADO_adRsnMoveFirst = 12
Global Const $ADO_adRsnMoveNext = 13
Global Const $ADO_adRsnMovePrevious = 14
Global Const $ADO_adRsnMoveLast = 15

; EventStatusEnum
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms681491(v=vs.85).aspx
Global Const $ADO_adStatusOK = 1 ; Indicates that the operation that caused the event was successful.
Global Const $ADO_adStatusErrorsOccurred = 2 ; Indicates that the operation that caused the event failed due to an error or errors.
Global Const $ADO_adStatusCantDeny = 3 ; Indicates that the operation cannot request cancellation of the pending operation.
Global Const $ADO_adStatusCancel = 4 ; Requests cancellation of the operation that caused the event to occur.
Global Const $ADO_adStatusUnwantedEvent = 5 ; Prevents subsequent notifications before the event method has finished executing.

; ExecuteOptionEnum
; https://msdn.microsoft.com/en-us/library/ms676517(v=vs.85).aspx
Global Const $ADO_adAsyncExecute = 0x10
Global Const $ADO_adAsyncFetch = 0x20
Global Const $ADO_adAsyncFetchNonBlocking = 0x40
Global Const $ADO_adExecuteNoRecords = 0x80
Global Const $ADO_adExecuteStream = 0x400
Global Const $ADO_adExecuteRecord = 2048
Global Const $ADO_adOptionUnspecified = -1

; FieldEnum
Global Const $ADO_adDefaultStream = -1
Global Const $ADO_adRecordURL = -2

; FieldStatusEnum
Global Const $ADO_adFieldOK = 0
Global Const $ADO_adFieldCantConvertValue = 2
Global Const $ADO_adFieldIsNull = 3
Global Const $ADO_adFieldTruncated = 4
Global Const $ADO_adFieldSignMismatch = 5
Global Const $ADO_adFieldDataOverflow = 6
Global Const $ADO_adFieldCantCreate = 7
Global Const $ADO_adFieldUnavailable = 8
Global Const $ADO_adFieldIntegrityViolation = 10
Global Const $ADO_adFieldSchemaViolation = 11
Global Const $ADO_adFieldBadStatus = 12
Global Const $ADO_adFieldDefault = 13
Global Const $ADO_adFieldIgnore = 15
Global Const $ADO_adFieldDoesNotExist = 16
Global Const $ADO_adFieldInvalidURL = 17
Global Const $ADO_adFieldResourceLocked = 18
Global Const $ADO_adFieldResourceExists = 19
Global Const $ADO_adFieldCannotComplete = 20
Global Const $ADO_adFieldVolumeNotFound = 21
Global Const $ADO_adFieldOutOfSpace = 22
Global Const $ADO_adFieldCannotDeleteSource = 23
Global Const $ADO_adFieldResourceOutOfScope = 25
Global Const $ADO_adFieldAlreadyExists = 26
Global Const $ADO_adFieldPendingChange = 0x40000
Global Const $ADO_adFieldPendingDelete = 0x20000
Global Const $ADO_adFieldPendingInsert = 0x10000
Global Const $ADO_adFieldPendingUnknown = 0x80000
Global Const $ADO_adFieldPendingUnknownDelete = 0x100000
Global Const $ADO_adFieldPermissionDenied = 0x9
Global Const $ADO_adFieldReadOnly = 0x24

; FilterGroupEnum
Global Const $ADO_adFilterNone = 0
Global Const $ADO_adFilterPendingRecords = 1
Global Const $ADO_adFilterAffectedRecords = 2
Global Const $ADO_adFilterFetchedRecords = 3
Global Const $ADO_adFilterConflictingRecords = 5

; GetRowsOptionEnum
Global Const $ADO_adGetRowsRest = -1

; IsolationLevelEnum
Global Const $ADO_adXactUnspecified = -1
Global Const $ADO_adXactChaos = 16
Global Const $ADO_adXactBrowse = 256
Global Const $ADO_adXactReadUncommitted = 256
Global Const $ADO_adXactCursorStability = 4096
Global Const $ADO_adXactReadCommitted = 4096
Global Const $ADO_adXactRepeatableRead = 65536
Global Const $ADO_adXactIsolated = 1048576
Global Const $ADO_adXactSerializable = 1048576

; LineSeparatorsEnum
Global Const $ADO_adCRLF = -1
Global Const $ADO_adLF = 10
Global Const $ADO_adCR = 13

; LockTypeEnum
Global Const $ADO_adLockUnspecified = -1
Global Const $ADO_adLockReadOnly = 1
Global Const $ADO_adLockPessimistic = 2
Global Const $ADO_adLockOptimistic = 3
Global Const $ADO_adLockBatchOptimistic = 4

; MarshalOptionsEnum
Global Const $ADO_adMarshalAll = 0
Global Const $ADO_adMarshalModifiedOnly = 1

; MoveRecordOptionsEnum
Global Const $ADO_adMoveUnspecified = -1
Global Const $ADO_adMoveOverWrite = 1
Global Const $ADO_adMoveDontUpdateLinks = 2
Global Const $ADO_adMoveAllowEmulation = 4

;~ ObjectStateEnum
;~ https://msdn.microsoft.com/en-us/library/windows/desktop/ms675546(v=vs.85).aspx
Global Const $ADO_adStateClosed = 0 ;   The object is closed
Global Const $ADO_adStateOpen = 1 ;   The object is open
Global Const $ADO_adStateConnecting = 2 ;   The object is connecting
Global Const $ADO_adStateExecuting = 4 ;   The object is executing a command
Global Const $ADO_adStateFetching = 8 ;   The rows of the object are being retrieved

; ParameterDirectionEnum
Global Const $ADO_adParamUnknown = 0
Global Const $ADO_adParamInput = 1
Global Const $ADO_adParamOutput = 2
Global Const $ADO_adParamInputOutput = 3
Global Const $ADO_adParamReturnValue = 4

; PersistFormatEnum
Global Const $ADO_adPersistADTG = 0 ; Indicates Microsoft Advanced Data TableGram (ADTG) format.
Global Const $ADO_adPersistXML = 1 ; Indicates Extensible Markup Language (XML) format.
Global Const $ADO_adPersistADO = 1 ; Indicates that ADO's own Extensible Markup Language (XML) format will be used. This value is the same as adPersistXML and is included for backwards compatibility.
Global Const $ADO_adPersistProviderSpecific = 2 ; Indicates that the provider will persist the Recordset using its own format.

; PositionEnum
Global Const $ADO_adPosEOF = -3
Global Const $ADO_adPosBOF = -2
Global Const $ADO_adPosUnknown = -1

; PropertyAttributesEnum
Global Const $ADO_adPropNotSupported = 0
Global Const $ADO_adPropRequired = 1
Global Const $ADO_adPropOptional = 2
Global Const $ADO_adPropRead = 512
Global Const $ADO_adPropWrite = 1024

; RecordCreateOptionsEnum
Global Const $ADO_adFailIfNotExists = -1
Global Const $ADO_adCreateNonCollection = 0
Global Const $ADO_adCreateCollection = 0x2000
Global Const $ADO_adCreateOverwrite = 0x4000000
Global Const $ADO_adCreateStructDoc = 0x80000000
Global Const $ADO_adOpenIfExists = 0x2000000

; RecordOpenOptionsEnum
Global Const $ADO_adDelayFetchFields = 0x8000
Global Const $ADO_adDelayFetchStream = 0x4000
Global Const $ADO_adOpenAsync = 0x1000
Global Const $ADO_adOpenExecuteCommand = 0x10000
Global Const $ADO_adOpenRecordUnspecified = -1
Global Const $ADO_adOpenOutput = 0x800000

; RecordStatusEnum
Global Const $ADO_adRecCanceled = 0x100
Global Const $ADO_adRecCantRelease = 0x400
Global Const $ADO_adRecConcurrencyViolation = 0x800
Global Const $ADO_adRecDBDeleted = 0x40000
Global Const $ADO_adRecDeleted = 0x4
Global Const $ADO_adRecIntegrityViolation = 0x1000
Global Const $ADO_adRecInvalid = 0x10
Global Const $ADO_adRecMaxChangesExceeded = 0x2000
Global Const $ADO_adRecModified = 0x2
Global Const $ADO_adRecMultipleChanges = 0x40
Global Const $ADO_adRecNew = 0x1
Global Const $ADO_adRecObjectOpen = 0x4000
Global Const $ADO_adRecOK = 0
Global Const $ADO_adRecOutOfMemory = 0x8000
Global Const $ADO_adRecPendingChanges = 0x80
Global Const $ADO_adRecPermissionDenied = 0x10000
Global Const $ADO_adRecSchemaViolation = 0x20000
Global Const $ADO_adRecUnmodified = 0x8

; RecordTypeEnum
Global Const $ADO_adSimpleRecord = 0
Global Const $ADO_adCollectionRecord = 1
Global Const $ADO_adRecordUnknown = -1
Global Const $ADO_adStructDoc = 2

; ResyncEnum
Global Const $ADO_adResyncUnderlyingValues = 1
Global Const $ADO_adResyncAllValues = 2

; SaveOptionsEnum
Global Const $ADO_adSaveCreateNotExist = 1
Global Const $ADO_adSaveCreateOverWrite = 2

; SchemaEnum
Global Const $ADO_adSchemaProviderSpecific = -1
Global Const $ADO_adSchemaAsserts = 0
Global Const $ADO_adSchemaCatalogs = 1
Global Const $ADO_adSchemaCharacterSets = 2
Global Const $ADO_adSchemaCollations = 3
Global Const $ADO_adSchemaCheckConstraints = 5
Global Const $ADO_adSchemaColumns = 4
Global Const $ADO_adSchemaConstraintColumnUsage = 6
Global Const $ADO_adSchemaConstraintTableUsage = 7
Global Const $ADO_adSchemaKeyColumnUsage = 8
Global Const $ADO_adSchemaReferentialConstraints = 9
Global Const $ADO_adSchemaTableConstraints = 10
Global Const $ADO_adSchemaColumnsDomainUsage = 11
Global Const $ADO_adSchemaIndexes = 12
Global Const $ADO_adSchemaColumnPrivileges = 13
Global Const $ADO_adSchemaTablePrivileges = 14
Global Const $ADO_adSchemaUsagePrivileges = 15
Global Const $ADO_adSchemaProcedures = 16
Global Const $ADO_adSchemaSchemata = 17
Global Const $ADO_adSchemaSQLLanguages = 18
Global Const $ADO_adSchemaStatistics = 19
Global Const $ADO_adSchemaTables = 20
Global Const $ADO_adSchemaTranslations = 21
Global Const $ADO_adSchemaProviderTypes = 22
Global Const $ADO_adSchemaViews = 23
Global Const $ADO_adSchemaViewColumnUsage = 24
Global Const $ADO_adSchemaViewTableUsage = 25
Global Const $ADO_adSchemaProcedureParameters = 26
Global Const $ADO_adSchemaForeignKeys = 27
Global Const $ADO_adSchemaPrimaryKeys = 28
Global Const $ADO_adSchemaProcedureColumns = 29
Global Const $ADO_adSchemaDBInfoKeywords = 30
Global Const $ADO_adSchemaDBInfoLiterals = 31
Global Const $ADO_adSchemaCubes = 32
Global Const $ADO_adSchemaDimensions = 33
Global Const $ADO_adSchemaHierarchies = 34
Global Const $ADO_adSchemaLevels = 35
Global Const $ADO_adSchemaMeasures = 36
Global Const $ADO_adSchemaProperties = 37
Global Const $ADO_adSchemaMembers = 38
Global Const $ADO_adSchemaTrustees = 39

; SearchDirectionEnum
Global Const $ADO_adSearchBackward = -1
Global Const $ADO_adSearchForward = 1

; SeekEnum
Global Const $ADO_adSeekFirstEQ = 1
Global Const $ADO_adSeekLastEQ = 2
Global Const $ADO_adSeekAfterEQ = 4
Global Const $ADO_adSeekAfter = 8
Global Const $ADO_adSeekBeforeEQ = 16
Global Const $ADO_adSeekBefore = 32

; StreamOpenOptionsEnum
Global Const $ADO_adOpenStreamUnspecified = -1
Global Const $ADO_adOpenStreamAsync = 1
Global Const $ADO_adOpenStreamFromRecord = 4

; StreamReadEnum
Global Const $ADO_adReadLine = -2
Global Const $ADO_adReadAll = -1

; StreamTypeEnum
Global Const $ADO_adTypeBinary = 1
Global Const $ADO_adTypeText = 2

; StreamWriteEnum
Global Const $ADO_adWriteChar = 0
Global Const $ADO_adWriteLine = 1

; StringFormatEnum
Global Const $ADO_adClipString = 2

; Attributes Property (ADO)
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms677543(v=vs.85).aspx

; XactAttributeEnum ; DEFAULT is 0;
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms681457(v=vs.85).aspx
Global Const $ADO_adXactCommitRetaining = 131072
Global Const $ADO_adXactAbortRetaining = 262144

; ParameterAttributesEnum ; DEFAULT is $ADO_adParamSigned
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms676687(v=vs.85).aspx
Global Const $ADO_adParamSigned = 16
Global Const $ADO_adParamNullable = 64
Global Const $ADO_adParamLong = 128

; FieldAttributeEnum
; https://msdn.microsoft.com/en-us/library/windows/desktop/ms676553(v=vs.85).aspx
Global Const $ADO_adFldCacheDeferred = 0x1000
Global Const $ADO_adFldFixed = 0x10
Global Const $ADO_adFldIsChapter = 0x2000
Global Const $ADO_adFldIsCollection = 0x40000
Global Const $ADO_adFldKeyColumn = 0x8000
Global Const $ADO_adFldIsDefaultStream = 0x20000
Global Const $ADO_adFldIsNullable = 0x20
Global Const $ADO_adFldIsRowURL = 0x10000
Global Const $ADO_adFldLong = 0x80
Global Const $ADO_adFldMayBeNull = 0x40
Global Const $ADO_adFldMayDefer = 0x2
Global Const $ADO_adFldNegativeScalem = 0x4000
Global Const $ADO_adFldRowID = 0x100
Global Const $ADO_adFldRowVersion = 0x200
Global Const $ADO_adFldUnknownUpdatable = 0x8
Global Const $ADO_adFldUnspecified = -1
Global Const $ADO_adFldUpdatable = 0x4

#EndRegion  ADO_CONSTANTS.au3 - MSDN Enumerated Constants

#Region ADO_CONSTANTS.au3 - TODO

; #FUNCTION# ====================================================================================================================
; Name ..........: _SQLState_Description
; Description ...:
; Syntax ........: _SQLState_Description($vState)
; Parameters ....: $vState              - a variant value.
; Return values .: None
; Author ........: Your Name
; Modified ......:
; Remarks .......:
; Related .......:
; Link ..........: http://www.rosam.se/doc/atsdoc/Server%20-%20Messages%20and%20Codes/ADC5906BBD514757BAEE546DC6F7A4FA/0.htm
; Link ..........: https://support.ca.com/cadocs/0/CA%20IDMS%20%20Server%20Option%2017-ENU/Bookshelf_Files/Javadoc/ca/idms/qcli/SQLState.html
; Link ..........: http://www.postgresql.org/docs/8.2/static/errcodes-appendix.html
; Example .......: No
; ===============================================================================================================================
Func _SQLState_Description($vState)
	Local $sDescription = ''

	Switch $vState
		Case ""
	EndSwitch
	Return $sDescription
EndFunc    ;==>_SQLState_Description

#EndRegion  ADO_CONSTANTS.au3 - TODO
