VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "FileReader"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileSpec                               As ITickfileSpecifier

Private mTs                                         As Scripting.TextStream
Private mAtEnd                                      As Boolean

Private mSizeBytes                                  As Long
Private mBytesRead                                  As Long

Private mFirstTickFileTime                          As Date

Private mTickFileVersion                            As TickFileVersions
Private mTickfileFormat                             As TickfileFormats

Private mEventCode                                  As TickfileEventCodes
Private mEventMessage                               As String

Private mContractSpecifier                          As IContractSpecifier

Private mBuffer                                     As String

Private mHeaderLineFields()                         As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get AtEnd()
AtEnd = mAtEnd
End Property

Friend Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

Friend Property Get CurrentRecord() As String
CurrentRecord = mBuffer
End Property

Friend Property Get EventCode() As TickfileEventCodes
EventCode = mEventCode
End Property

Friend Property Get EventMessage() As String
EventMessage = mEventMessage
End Property

Friend Property Get FirstTickFileTime() As Date
FirstTickFileTime = mFirstTickFileTime
End Property

Friend Property Get SizeBytes() As Long
SizeBytes = mSizeBytes
End Property

Friend Property Get BytesRead() As Long
BytesRead = mBytesRead
End Property

Friend Property Get TickfileFormat() As TickfileFormats
TickfileFormat = mTickfileFormat
End Property

Friend Property Get TickfileVersion() As TickFileVersions
TickfileVersion = mTickFileVersion
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Finish()
If Not mTs Is Nothing Then mTs.Close
Set mTs = Nothing
End Function

Friend Function Initialise( _
                ByVal pTickfileSpec As ITickfileSpecifier) As Boolean
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTickfileSpec = pTickfileSpec

gLogger.Log "Opening tickfile " & mTickfileSpec.Filename, ProcName, ModuleName
    
If Not openFile(mTickfileSpec.Filename) Then Exit Function
If fileIsEmpty(mTickfileSpec.Filename) Then Exit Function
If isESignalFile(mTickfileSpec.Filename) Then Initialise = True: Exit Function
If Not getHeaderLineFields(mHeaderLineFields, mTickfileSpec.Filename) Then Exit Function

recogniseFileFormat mHeaderLineFields
If Not getContractInfo(mHeaderLineFields, mTickfileSpec.Filename) Then Exit Function

If Not getNextRecord(mBuffer) Then
    NotifyEvent TickfileEventCodes.TickfileEventFileIsEmpty, "Empty tickfile"
    Initialise = False
    Exit Function
End If

repositionFile mTickfileSpec.Filename

Initialise = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function MoveNext() As Boolean
Const ProcName As String = "MoveNext"
On Error GoTo Err

If mAtEnd Then
    MoveNext = False
Else
    MoveNext = getNextRecord(mBuffer)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function fileIsEmpty(ByVal pFilename As String) As Boolean
If Not mTs.AtEndOfStream Then Exit Function

mTs.Close
Set mTs = Nothing
mSizeBytes = 0
mTickFileVersion = TickFileVersions.UnknownVersion
mTickfileFormat = TickfileFormats.TickfileUnknown
NotifyEvent TickfileEventCodes.TickfileEventFileIsInvalid, "Tickfile mTickfileFormat cannot be determined: " & pFilename
fileIsEmpty = True
End Function

Private Function getCrescendoV1ContractInfo(ByRef pHeaderLineFields() As String, ByVal pFilename As String) As Boolean
Const ProcName As String = "getCrescendoV1ContractInfo"
On Error GoTo Err

On Error Resume Next
Set mContractSpecifier = CreateContractSpecifier( _
                                , _
                                pHeaderLineFields(TickfileFieldsV1.Symbol), _
                                pHeaderLineFields(TickfileFieldsV1.Exchange), _
                                SecTypeFuture, _
                                , _
                                pHeaderLineFields(TickfileFieldsV1.Expiry))

If Err.Number <> 0 Then
    mTs.Close
    NotifyEvent TickfileEventCodes.TickfileEventFileIsInvalid, "Tickfile mTickfileFormat/mTickFileVersion not determined: " & pFilename
    getCrescendoV1ContractInfo = False
    Exit Function
End If

On Error GoTo Err

mFirstTickFileTime = CDate(pHeaderLineFields(TickfileFieldsV1.TimestampString))

getCrescendoV1ContractInfo = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCrescendoV2ContractInfo(ByRef pHeaderLineFields() As String) As Boolean
Const ProcName As String = "getCrescendoV2ContractInfo"
On Error GoTo Err

Set mContractSpecifier = CreateContractSpecifier( _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV2.Symbol), _
                                pHeaderLineFields(TickfileHeaderFieldsV2.Exchange), _
                                SecTypeFuture, _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV2.Expiry))

mFirstTickFileTime = CDate(pHeaderLineFields(TickfileHeaderFieldsV2.StartTime))

getCrescendoV2ContractInfo = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getContractInfo( _
                ByRef pHeaderLineDataFields() As String, _
                ByVal pFilename As String) As Boolean
Const ProcName As String = "getContractInfo"
On Error GoTo Err

Select Case mTickFileVersion
Case TickFileVersions.CrescendoV1
    getContractInfo = getCrescendoV1ContractInfo(pHeaderLineDataFields, pFilename)
Case TickFileVersions.CrescendoV2
    getContractInfo = getCrescendoV2ContractInfo(pHeaderLineDataFields)
Case TickFileVersions.TradeBuildV3, _
        TickFileVersions.TradeBuildV4, _
        TickFileVersions.TradeBuildV5
    getContractInfo = getTradeBuildContractInfo(pHeaderLineDataFields, pFilename)
Case Else
    Assert False, "Invalid tickfile version"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getHeaderLineFields(ByRef pAr() As String, ByVal pFilename As String) As Boolean
Dim lBuffer As String
If Not getNextRecord(lBuffer) Then
    NotifyEvent TickfileEventCodes.TickfileEventFileIsInvalid, "Tickfile mTickfileFormat cannot be determined: " & pFilename
    getHeaderLineFields = False
    Exit Function
End If

mHeaderLineFields = Split(lBuffer, ",", -1, vbTextCompare)
getHeaderLineFields = True
End Function

Private Function getNextRecord(ByRef pBuffer As String) As Boolean
Const ProcName As String = "getNextLine"
On Error GoTo Err

If mAtEnd Then
    pBuffer = ""
    getNextRecord = False
    Exit Function
End If

Do While Not mTs.AtEndOfStream
    pBuffer = mTs.ReadLine
    mBytesRead = mBytesRead + Len(pBuffer) + 2
    If Not isCommentOrBlank(pBuffer) Then
        getNextRecord = True
        Exit Do
    End If
Loop

If mTs.AtEndOfStream Then
    mTs.Close
    Set mTs = Nothing
    mAtEnd = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTradeBuildContractInfo( _
                ByRef pHeaderLineFields() As String, _
                ByVal pFilename As String) As Boolean
Const ProcName As String = "getTradeBuildContractInfo"
On Error GoTo Err

Set mContractSpecifier = CreateContractSpecifier( _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV3.Symbol), _
                                pHeaderLineFields(TickfileHeaderFieldsV3.Exchange), _
                                SecTypeFuture, _
                                , _
                                pHeaderLineFields(TickfileHeaderFieldsV3.Expiry))

mFirstTickFileTime = CDate(pHeaderLineFields(TickfileHeaderFieldsV3.StartTime))

' read the serialised contract details
Dim lBuffer As String
If Not getNextRecord(lBuffer) Then
    NotifyEvent TickfileEventCodes.TickfileEventFileIsInvalid, "Invalid tickfile - no contract details line: " & pFilename
    getTradeBuildContractInfo = False
    Exit Function
End If

If Left$(lBuffer, Len(CONTRACT_DETAILS_MARKER)) <> CONTRACT_DETAILS_MARKER Then
    NotifyEvent TickfileEventCodes.TickfileEventFileIsInvalid, "Invalid tickfile - no contract details line: " & pFilename
    getTradeBuildContractInfo = False
    Exit Function
End If

If mTickFileVersion = TickFileVersions.TradeBuildV3 Then
        ' can no longer get contract details from this mTickFileVersion so do nothing
Else
    Dim xmlContract As String
    xmlContract = Right$(lBuffer, Len(lBuffer) - Len(CONTRACT_DETAILS_MARKER))
    Set mTickfileSpec.ContractFuture = CreateContractFuture(CreateContractFromXML(xmlContract))
End If

getTradeBuildContractInfo = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isESignalFile(ByVal pFilename As String) As Boolean
Const ProcName As String = "isESignalFile"
On Error GoTo Err

Dim lDataFields() As String

lDataFields = Split(pFilename, ".", -1, vbTextCompare)
If UCase$(lDataFields(UBound(lDataFields))) = "EPF" Then
    mTickfileFormat = TickfileFormats.TickfileESignal
    mTickFileVersion = TickFileVersions.ESignal
    isESignalFile = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isCommentOrBlank(ByVal pBuffer As String) As Boolean
If pBuffer = "" Then
    isCommentOrBlank = True
ElseIf mTickfileFormat = TickfileCrescendo And Left$(pBuffer, 2) = "//" Then
    isCommentOrBlank = True
ElseIf mTickfileFormat = TickfileTradeBuild And Left$(pBuffer, 2) = "//" Then
    isCommentOrBlank = True
ElseIf mTickfileFormat = TickfileESignal And Left$(pBuffer, 1) = ";" Then
    isCommentOrBlank = True
Else
    isCommentOrBlank = False
End If
End Function

Private Sub NotifyEvent(ByVal pEventCode As TickfileEventCodes, pMessage As String)
mEventCode = pEventCode
mEventMessage = pMessage
End Sub

Private Function openFile(ByVal pFilename As String) As Boolean
Const ProcName As String = "openFile"
On Error GoTo Err

Dim lFileSys As Scripting.FileSystemObject
Set lFileSys = New FileSystemObject

Set mTs = lFileSys.OpenTextFile(pFilename, ForReading, False)

Dim f As File
Set f = lFileSys.GetFile(pFilename)
mSizeBytes = f.Size

openFile = True

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrFileNotFound Then
    NotifyEvent TickfileEventCodes.TickfileEventFileDoesNotExist, "Tickfile does not exist: " & pFilename
ElseIf Err.Number = VBErrorCodes.VbErrFileAlreadyOpen Then
    NotifyEvent TickfileEventCodes.TickfileEventFileDoesNotExist, "File is already open: " & pFilename
ElseIf Err.Number = VBErrorCodes.VbErrPathNotFound Then
    NotifyEvent TickfileEventCodes.TickfileEventFileDoesNotExist, "Tickfile path not found: " & pFilename
ElseIf Err.Number = VBErrorCodes.VbErrPermissionDenied Then
    NotifyEvent TickfileEventCodes.TickfileEventFileDoesNotExist, "Tickfile access denied: " & pFilename
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Sub recogniseFileFormat(ByRef pHeaderLineFields() As String)
Const ProcName As String = "recogniseFileFormat"
On Error GoTo Err

If pHeaderLineFields(TickfileHeaderFieldsV3.ContentDeclarer) = TICKFILE_DECLARER Then
    Select Case pHeaderLineFields(TickfileHeaderFieldsV3.Version)
    Case 2
        mTickfileFormat = TickfileFormats.TickfileCrescendo
        mTickFileVersion = TickFileVersions.CrescendoV2
    Case 3
        mTickfileFormat = TickfileFormats.TickfileTradeBuild
        mTickFileVersion = TickFileVersions.TradeBuildV3
    Case 4
        mTickfileFormat = TickfileFormats.TickfileTradeBuild
        mTickFileVersion = TickFileVersions.TradeBuildV4
    Case 5
        mTickfileFormat = TickfileFormats.TickfileTradeBuild
        mTickFileVersion = TickFileVersions.TradeBuildV5
    End Select
Else
    mTickfileFormat = TickfileFormats.TickfileCrescendo
    mTickFileVersion = TickFileVersions.CrescendoV1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub repositionFile(ByVal pFilename As String)
Const ProcName As String = "repositionFile"
On Error GoTo Err

' reopen the file so it's correctly positioned to start reading data
mTs.Close
mBytesRead = 0

Assert openFile(pFilename), "File cannot be re-opened"

Dim lBuffer As String

Select Case mTickFileVersion
Case TradeBuildV3, TradeBuildV4, TradeBuildV5
    ' skip header record
    Assert getNextRecord(lBuffer), "Failed to get record"
    ' skip contract details record
    Assert getNextRecord(lBuffer), "Failed to get record"
Case CrescendoV1
    ' No preliminary records for this format
Case CrescendoV2
    ' skip header record
    Assert getNextRecord(lBuffer), "Failed to get record"
Case ESignal
    ' No preliminary records for this format
Case Else
    Assert False, "Invalid tickfile version"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
