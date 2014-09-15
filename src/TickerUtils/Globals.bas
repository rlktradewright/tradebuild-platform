Attribute VB_Name = "Globals"
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

Public Const ProjectName                            As String = "TickerUtils27"
Private Const ModuleName                            As String = "Globals"

'@================================================================================
' Member variables
'@================================================================================

Private mListenersCollection                        As New EnumerableCollection

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

Public Property Get gLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("tickerutils", ProjectName)
Set gLogger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

'Public Sub gAddListener( _
'                ByVal pEventName As String, _
'                ByVal pSource As Object, _
'                ByVal pListener As Object)
'Const ProcName As String = "gAddListener"
'On Error GoTo Err
'
'Dim lKey As String
'lKey = generateKey(pEventName, pSource)
'
'Dim lListeners As Listeners
'
'If mListenersCollection.Contains(lKey) Then
'    Set lListeners = mListenersCollection(lKey)
'Else
'    Set lListeners = New Listeners
'    mListenersCollection.Add lListeners, lKey
'End If
'
'lListeners.Add pListener
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub
'
'Public Function gGetListeners( _
'                ByVal pEventName As String, _
'                ByVal pSource As Object) As Listeners
'Const ProcName As String = "gGetListeners"
'On Error GoTo Err
'
'Static sEmptyListeners As New Listeners
'
'Set gGetListeners = mListenersCollection(generateKey(pEventName, pSource))
'If Not gGetListeners Is Nothing Then Exit Function
'Set gGetListeners = sEmptyListeners
'
'Exit Function
'
'Err:
'If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
'gHandleUnexpectedError ProcName, ModuleName
'End Function
'
'Public Sub gRemoveListener( _
'                ByVal pEventName As String, _
'                ByVal pSource As Object, _
'                ByVal pListener As Object)
'Const ProcName As String = "gRemoveListener"
'On Error GoTo Err
'
'Dim lKey As String
'lKey = generateKey(pEventName, pSource)
'
'If mListenersCollection.Contains(lKey) Then
'    Dim lListeners As Listeners
'    Set lListeners = mListenersCollection.Item(lKey)
'    lListeners.Remove pListener
'    If lListeners.Count = 0 Then mListenersCollection.Remove lKey
'End If
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateKey(ByVal pEventName As String, ByVal pSource As Object) As String
generateKey = pEventName & "$" & GetObjectKey(pSource)
End Function


