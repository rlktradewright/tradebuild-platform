VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPersistenceDataStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements Enumerable
Implements IOrderPersistenceDataStore

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

Private Const ModuleName                            As String = "OrderPersistenceDataStore"

Private Const KeyDataSeparator                      As String = ":"
Private Const OrderInfoSession                      As String = "SESSION"
Private Const TimestampSeparator                    As String = "  "

'@================================================================================
' Member variables
'@================================================================================

Private mName                                       As String

Private mRecoveryFilePath                           As String
Private mScopeName                                  As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Const ProcName As String = "Enumerable_Enumerator"
On Error GoTo Err

Set Enumerable_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Function Enumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
Dim lTS As TextStream

Const ProcName As String = "Enumerable_NextItem"
On Error GoTo Err

Set lTS = cookie
If lTS Is Nothing Then
    Enumerable_NextItem = False
ElseIf getNextLine(lTS, lLine) Then
    Item = lLine
    Enumerable_NextItem = True
Else
    Enumerable_NextItem = False
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub Enumerable_Remove(cookie As Variant)
Const ProcName As String = "Enumerable_Remove"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Remove not supported"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Const ProcName As String = "Enumerable_Skip"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Skip not supported"

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Get Enumerable_UpdateSequence() As Long

End Property

'@================================================================================
' IOrderPersistenceDataStore Interface Members
'@================================================================================

Private Function IOrderPersistenceDataStore_GetEnumeratorForScope(ByVal pScopeName As String) As Enumerator
Const ProcName As String = "IOrderPersistenceDataStore_GetEnumeratorForScope"
On Error GoTo Err

On Error GoTo Err

mScopeName = scopeName
Set IOrderPersistenceDataStore_GetEnumeratorForScope = getEnumerator

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub IOrderPersistenceDataStore_SaveOrderInfo(ByVal pOrderInfo As String)
Const ProcName As String = "IOrderPersistenceDataStore_SaveOrderInfo"
On Error GoTo Err

mTS.WriteLine orderInfo

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub IOrderPersistenceDataStore_StartSession(ByVal pSessionId As String)
Const ProcName As String = "IOrderPersistenceDataStore_StartSession"
On Error GoTo Err

createPersistenceFile SessionId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get RecoveryFilePath() As String
RecoveryFilePath = mRecoveryFilePath
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise(ByVal pRecoveryFilePath As String)
mRecoveryFilePath = pRecoveryFilePath
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createPersistenceFile(SessionId As String)
Const ProcName As String = "createPersistenceFile"
Dim fso As New FileSystemObject

On Error GoTo Err

If Not mTS Is Nothing Then mTS.Close

Set mTS = fso.CreateTextFile(getRecoveryFileName(), True)
 
mTS.Write OrderInfoSession
mTS.Write ": "
mTS.WriteLine SessionId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Dim lTS As TextStream
Dim lLine As String

If mScopeName = "" Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Scopename has not been set"

Set lTS = getPersistenceFile
If Not lTS Is Nothing Then getNextLine lTS, lLine   ' first line is session id line so ignore it
Set getEnumerator = CreateEnumerator(Me, lTS)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getNextLine(ByVal pTs As TextStream, ByRef result As String) As Boolean
Const ProcName As String = "getNextLine"

On Error GoTo Err

Do While Not mTS.AtEndOfStream
    result = mTS.ReadLine
    If result <> "" And Left$(result, 2) <> "//" Then
        getNextLine = True
        Exit Do
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getPersistenceFile() As TextStream
Dim fso As New FileSystemObject
Const ProcName As String = "getPersistenceFile"
On Error GoTo Err

On Error Resume Next
Set getPersistenceFile = fso.OpenTextFile(getRecoveryFileName(), _
                        IOMode.ForReading, _
                        False)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getRecoveryFileName() As String
If mScopeName <> "" Then
    getRecoveryFileName = mRecoveryFilePath & "\OrderInfo(" & mScopeName & ").txt"
Else
    getRecoveryFileName = mRecoveryFilePath & "\OrderInfo.txt"
End If
End Function

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHandle
End Sub



