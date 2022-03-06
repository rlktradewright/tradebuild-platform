VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PriceIncrementRuleset"
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type Rule
    PriceBandLow    As Double
    Increment       As Double
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "PriceIncrementRuleset"

'@================================================================================
' Member variables
'@================================================================================

Private mRulesID                                    As Long

Private mRules()                                    As Rule
Private mNextRulesIndex                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mRules(0) As Rule
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

Friend Property Get RulesID() As Long
RulesID = mRulesID
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add( _
                ByVal pPriceBandLow As Double, _
                ByVal pIncrement As Double)
If mNextRulesIndex > UBound(mRules) Then
    ReDim Preserve mRules(2 * (UBound(mRules) + 1) - 1)
End If
mRules(mNextRulesIndex).PriceBandLow = pPriceBandLow
mRules(mNextRulesIndex).Increment = pIncrement
mNextRulesIndex = mNextRulesIndex + 1
End Sub

Friend Sub Initialise( _
                ByVal pRulesID As Long)
mRulesID = pRulesID
End Sub

Friend Function GetIncrement( _
                ByVal pPrice As Double) As Double
If pPrice = 0 Then
    If mNextRulesIndex = 1 Then
        GetIncrement = mRules(0).Increment
    Else
        GetIncrement = 0
    End If
    Exit Function
End If

Dim i As Long
For i = 0 To mNextRulesIndex - 1
    If pPrice >= mRules(i).PriceBandLow Then
        GetIncrement = mRules(i).Increment
    Else
        Exit For
    End If
Next
End Function

'@================================================================================
' Helper Functions
'@================================================================================




