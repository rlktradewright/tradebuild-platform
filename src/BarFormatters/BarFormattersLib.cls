VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFormattersLib"
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

Implements IBarFormatterLibrary

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

Private Const ModuleName                            As String = "BarFormattersLib"

'@================================================================================
' Member variables
'@================================================================================

Private mName As String

Private mFactoryNames() As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = "TradeWright Bar Formatters Library"

ReDim mFactoryNames(0) As String
addFactoryName NameDunnigan
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Function IBarFormatterLibrary_CreateFactory(ByVal pName As String) As IBarFormatterFactory
Const ProcName As String = "IBarFormatterLibrary_CreateFactory"
On Error GoTo Err

Select Case UCase$(pName)
Case UCase$(NameDunnigan)
    Set IBarFormatterLibrary_CreateFactory = New DunniganFactory
Case Else
    AssertArgument False, "Invalid BarFormatterFactory name"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IBarFormatterLibrary_GetFactoryNames() As String()
IBarFormatterLibrary_GetFactoryNames = mFactoryNames
End Function

Private Property Let IBarFormatterLibrary_Name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get IBarFormatterLibrary_Name() As String
IBarFormatterLibrary_Name = mName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addFactoryName(ByVal Name As String)
Const ProcName As String = "addFactoryName"
On Error GoTo Err

If mFactoryNames(0) <> "" Then
    ReDim Preserve mFactoryNames(UBound(mFactoryNames) + 1) As String
End If
mFactoryNames(UBound(mFactoryNames)) = Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




