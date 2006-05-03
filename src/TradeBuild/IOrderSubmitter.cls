VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrderSubmitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Public Function cancelOrderPlex(ByVal pOrderPlex As orderPlex) As Boolean
End Function

Public Sub cancelStopOrder(ByVal pOrderPlex As orderPlex)
End Sub

Public Sub cancelTargetOrder(ByVal pOrderPlex As orderPlex)
End Sub

Public Sub executeOrderPlex(ByVal pOrderPlex As orderPlex)
End Sub

Public Sub modifyOrderPlex( _
                ByVal pOrderPlex As orderPlex, _
                ByVal entryOrderChanged As Boolean, _
                ByVal stopOrderChanged As Boolean, _
                ByVal targetOrderChanged As Boolean, _
                ByVal closeoutOrderChanged As Boolean)
End Sub

Public Sub resubmitStopAndTargetOrders(ByVal pOrderPlex As orderPlex)
End Sub

Public Sub resubmitStopOrder(ByVal pOrderPlex As orderPlex)
End Sub

Public Sub resubmitTargetOrder(ByVal pOrderPlex As orderPlex)
End Sub

'================================================================================
' Helper Functions
'================================================================================

