VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommandList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Private Const ModuleName                            As String = "CommandList"

'@================================================================================
' Member variables
'@================================================================================

Private mCommands()                                 As Command
Private mSB                                         As StringBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSB = CreateStringBuilder
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get CommandNames() As String
CommandNames = mSB.ToString
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ParamArray pCommands() As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

ReDim mCommands(UBound(pCommands)) As Command

Dim i As Long
For i = 0 To UBound(pCommands)
    Dim lCommand As Command: Set lCommand = pCommands(i)
    If i <> 0 Then mSB.Append " "
    mSB.Append lCommand.Names
    Set mCommands(i) = lCommand
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function IsValidCommandName(ByVal pCommandName As String) As Boolean
Dim i As Long
For i = 0 To UBound(mCommands)
    If mCommands(i).parse(pCommandName) Then
        IsValidCommandName = True
        Exit Function
    End If
Next
End Function

'@================================================================================
' Helper Functions
'@================================================================================




