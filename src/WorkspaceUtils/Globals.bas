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

Public Const ProjectName                            As String = "WorkspaceUtils"
Private Const ModuleName                            As String = "Globals"

Public Const ConfigSectionTickers                   As String = "Tickers"
Public Const ConfigSectionWorkspace                 As String = "Workspace"

'@================================================================================
' Member variables
'@================================================================================

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

'@================================================================================
' Methods
'@================================================================================

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

Public Sub gLog(ByRef pMsg As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err

Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("workspaceutils", ProjectName)

sLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Function gNotifyExistingCollectionMembers( _
'                ByVal pCollection As Variant, _
'                ByVal pListener As ICollectionChangeListener, _
'                ByVal pSource As Object)
'Const ProcName As String = "gNotifyExistingCollectionMembers"
'On Error GoTo Err
'
'Dim lItem As Variant
'If VarType(pCollection) And vbArray = vbArray Then
'    For Each lItem In pCollection
'        notifyCollectionMember lItem, pSource, pListener
'    Next
'ElseIf Not IsObject(pCollection) Then
'    AssertArgument False, "pCollection argument must be an array or a VB6 collection or must implement Enumerable"
'ElseIf TypeOf pCollection Is Collection Then
'    For Each lItem In pCollection
'        notifyCollectionMember lItem, pSource, pListener
'    Next
'ElseIf TypeOf pCollection Is IEnumerable Then
'    Dim enColl As Enumerable
'    Set enColl = pCollection
'
'    Dim en As Enumerator
'    Set en = enColl.Enumerator
'
'    Do While en.MoveNext
'        notifyCollectionMember en.Current, pSource, pListener
'    Loop
'Else
'     AssertArgument False, "pCollection argument must be an array or a VB6 collection or must implement Enumerable"
'End If
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function

'@================================================================================
' Helper Functions
'@================================================================================




