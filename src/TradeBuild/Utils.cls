VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Utils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'' A GlobalMultiUse class that provides a variety of useful functions.

'@================================================================================
' Description
'@================================================================================
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

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

''
'
' Converts a member of the EntryOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
'
' @param pEntryOrderType the EntryOrderTypes value to be converted
'
'@/
Public Function EntryOrderTypeToOrderType( _
                ByVal pEntryOrderType As TradeBuild25.EntryOrderTypes) As TradeBuild25.OrderTypes
EntryOrderTypeToOrderType = gEntryOrderTypeToOrderType(pEntryOrderType)
End Function

Public Function EntryOrderTypeToString( _
                ByVal value As EntryOrderTypes) As String
EntryOrderTypeToString = gEntryOrderTypeToString(value)
End Function

Public Function EntryOrderTypeToShortString( _
                ByVal value As EntryOrderTypes) As String
EntryOrderTypeToShortString = gEntryOrderTypeToShortString(value)
End Function

Public Function LegOpenCloseFromString( _
                ByVal value As String) As TradeBuild25.LegOpenClose
LegOpenCloseFromString = gLegOpenCloseFromString(value)
End Function

Public Function LegOpenCloseToString( _
                ByVal value As TradeBuild25.LegOpenClose) As String
LegOpenCloseToString = gLegOpenCloseToString(value)
End Function

Public Function OptionRightFromString( _
                ByVal value As String) As OptionRights
OptionRightFromString = gOptionRightFromString(value)
End Function

Public Function OptionRightToString( _
                ByVal value As OptionRights) As String
OptionRightToString = gOptionRightToString(value)
End Function

Public Function OrderActionFromString( _
                ByVal value As String) As OrderActions
OrderActionFromString = gOrderActionFromString(value)
End Function

Public Function OrderActionToString( _
                ByVal value As OrderActions) As String
OrderActionToString = gOrderActionToString(value)
End Function

Public Function OrderStatusToString( _
                ByVal value As OrderStatuses) As String
OrderStatusToString = gOrderStatusToString(value)
End Function

Public Function OrderStopTriggerMethodToString( _
                ByVal value As StopTriggerMethods) As String
OrderStopTriggerMethodToString = gOrderStopTriggerMethodToString(value)
End Function

Public Function OrderTIFToString( _
                ByVal value As OrderTifs) As String
OrderTIFToString = gOrderTIFToString(value)
End Function

Public Function OrderTypeToString( _
                ByVal value As OrderTypes) As String
OrderTypeToString = gOrderTypeToString(value)
End Function

Public Function SecTypeFromString( _
                ByVal value As String) As SecurityTypes
SecTypeFromString = gSecTypeFromString(value)
End Function

Public Function SecTypeToShortString( _
                ByVal value As SecurityTypes) As String
SecTypeToShortString = gSecTypeToShortString(value)
End Function

Public Function SecTypeToString( _
                ByVal value As SecurityTypes) As String
SecTypeToString = gSecTypeToString(value)
End Function

''
' Converts a member of the StopOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pStopOrderType the StopOrderTypes value to be converted
' @see
'
'@/
Public Function StopOrderTypeToOrderType( _
                ByVal pStopOrderType As TradeBuild25.StopOrderTypes) As TradeBuild25.OrderTypes
StopOrderTypeToOrderType = gStopOrderTypeToOrderType(pStopOrderType)
End Function

Public Function StopOrderTypeToString( _
                ByVal value As StopOrderTypes)
StopOrderTypeToString = gStopOrderTypeToString(value)
End Function

''
' Converts a member of the <code>TargetOrderTypes</code> enumeration to the equivalent <code>OrderTypes</code> value.
'
' @return           the <code>OrderTypes</code> value corresponding to the parameter
' @param pTargetOrderType the <code>TargetOrderTypes</code> value to be converted
' @see
'
'@/
Public Function TargetOrderTypeToOrderType( _
                ByVal pTargetOrderType As TradeBuild25.TargetOrderTypes) As TradeBuild25.OrderTypes
TargetOrderTypeToOrderType = gTargetOrderTypeToOrderType(pTargetOrderType)
End Function

Public Function TargetOrderTypeToString( _
                ByVal value As TargetOrderTypes) As String
TargetOrderTypeToString = gTargetOrderTypeToString(value)
End Function

Public Function TickfileSpecifierToString( _
                TickfileSpec As TradeBuild25.TickfileSpecifier) As String
TickfileSpecifierToString = gTickfileSpecifierToString(TickfileSpec)
End Function

'@================================================================================
' Helper Functions
'@================================================================================



