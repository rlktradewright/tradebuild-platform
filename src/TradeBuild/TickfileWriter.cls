VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "TickfileWriter"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mWriter                         As TradeBuildSP.ITickfileWriter
Private mServiceProviderWriterReady As Boolean

Private mPendingWrites As Collection

'Private WithEvents mTimer As IntervalTimer

'Private mSelfRef As TickfileWriter

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Collection
End Sub

Private Sub Class_Terminate()
Debug.Print "TradeBuild26.TickfileWriter - object terminated"
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get NumberOfPendingWrites() As Long
Const ProcName As String = "NumberOfPendingWrites"
On Error GoTo Err

NumberOfPendingWrites = mPendingWrites.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let ServiceProviderWriter(ByVal value As TradeBuildSP.ITickfileWriter)
Const ProcName As String = "ServiceProviderWriter"
On Error GoTo Err

Set mWriter = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get ready() As Boolean
Const ProcName As String = "ready"
On Error GoTo Err

ready = mServiceProviderWriterReady

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub CloseOutputFile()
Const ProcName As String = "CloseOutputFile"
On Error GoTo Err

GLogLogger.Log pMsg:="Closing output tickfile", pProcName:=ProcName, pModName:=ModuleName
Debug.Print "TradeBuild26.TickfileWriter - closing output tickfile"
mWriter.CloseOutputFile

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ServiceProviderWriterNotReady()
Const ProcName As String = "ServiceProviderWriterNotReady"
On Error GoTo Err

mServiceProviderWriterReady = False

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ServiceProviderWriterReady()
Dim Item As Variant
Dim pendingTick As GenericTick

Const ProcName As String = "ServiceProviderWriterReady"
On Error GoTo Err

mServiceProviderWriterReady = True

For Each Item In mPendingWrites
    pendingTick = Item
    mWriter.writeTick pendingTick
Next

Set mPendingWrites = New Collection

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub writeTick( _
                pTick As GenericTick)
Const ProcName As String = "writeTick"
On Error GoTo Err

If Not mServiceProviderWriterReady Then
    mPendingWrites.Add pTick
Else
    mWriter.writeTick pTick
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


