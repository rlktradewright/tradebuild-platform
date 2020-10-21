VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScanResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IScanResult

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

Private Const ModuleName                            As String = "ScanResult"

'@================================================================================
' Member variables
'@================================================================================

Private mAttributes                                 As Parameters

Private mContract                                   As IContract

Private mRank                                       As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mAttributes = New Parameters
End Sub

'@================================================================================
' IScanResult Interface Members
'@================================================================================

Private Property Get IScanResult_Attributes() As Parameters
Set IScanResult_Attributes = mAttributes
End Property

Private Property Get IScanResult_Contract() As IContract
Set IScanResult_Contract = mContract
End Property

Private Property Get IScanResult_Rank() As Long
IScanResult_Rank = mRank
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Contract(ByVal Value As IContract)
Set mContract = Value
End Property

Friend Property Let Rank(ByVal Value As Long)
mRank = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddAttribute( _
                ByVal pName As String, _
                ByVal pValue As String)
mAttributes.SetParameterValue pName, pValue
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




