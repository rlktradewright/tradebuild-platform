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

Implements BarFormatterLibrary

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

Private Function BarFormatterLibrary_CreateFactory(ByVal pName As String) As BarFormatterFactory
Const ProcName As String = "BarFormatterLibrary_CreateFactory"
On Error GoTo Err

Select Case UCase$(pName)
Case UCase$(NameDunnigan)
    Set BarFormatterLibrary_CreateFactory = New DunniganFactory
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid BarFormatterFactory name"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function BarFormatterLibrary_GetFactoryNames() As String()
BarFormatterLibrary_GetFactoryNames = mFactoryNames
End Function

Private Property Let BarFormatterLibrary_Name(ByVal RHS As String)
mName = RHS
End Property

Private Property Get BarFormatterLibrary_Name() As String
BarFormatterLibrary_Name = mName
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

Private Sub addFactoryName(ByVal name As String)
Const ProcName As String = "addFactoryName"
On Error GoTo Err

If mFactoryNames(0) <> "" Then
    ReDim Preserve mFactoryNames(UBound(mFactoryNames) + 1) As String
End If
mFactoryNames(UBound(mFactoryNames)) = name

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub




