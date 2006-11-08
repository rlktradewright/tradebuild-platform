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

'================================================================================
' Description
'================================================================================
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

'/**
' Converts a member of the EntryOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pEntryOrderType the EntryOrderTypes value to be converted
' @ see
'
'*/
Public Function EntryOrderTypeToOrderType( _
                ByVal pEntryOrderType As TradeBuild.EntryOrderTypes) As TradeBuild.OrderTypes
EntryOrderTypeToOrderType = gEntryOrderTypeToOrderType(pEntryOrderType)
End Function

Public Function EntryOrderTypeToString(ByVal value As EntryOrderTypes) As String
EntryOrderTypeToString = gEntryOrderTypeToString(value)
End Function

Public Function EntryOrderTypeToShortString(ByVal value As EntryOrderTypes) As String
EntryOrderTypeToShortString = gEntryOrderTypeToShortString(value)
End Function

Public Function FormatTimestamp(ByVal timestamp As Date, _
                                Optional ByVal formatOption As TimestampFormats = TimestampDateAndTime, _
                                Optional ByVal formatString As String = "yyyymmddhhnnss") As String
gFormatTimestamp timestamp, formatOption, formatString
End Function

Public Function GenerateGUID() As GUIDStruct
GenerateGUID = gGenerateGUID
End Function

Public Function GenerateGUIDString() As String
GenerateGUIDString = gGenerateGUIDString
End Function

Public Function GenerateID() As Long
GenerateID = gGenerateID
End Function

Public Function GenerateTextID() As String
GenerateTextID = gGenerateIDString
End Function

Public Function GetTimestamp() As Date
GetTimestamp = gCurrentTime
End Function

Public Function GUIDToString(ByRef pGUID As GUIDStruct) As String
GUIDToString = gGUIDToString(pGUID)
End Function

Public Function LegOpenCloseFromString(ByVal value As String) As TradeBuild.LegOpenClose
LegOpenCloseFromString = gLegOpenCloseFromString(value)
End Function

Public Function LegOpenCloseToString(ByVal value As TradeBuild.LegOpenClose) As String
LegOpenCloseToString = gLegOpenCloseToString(value)
End Function

Public Function newWeakReference(ByVal target As Object) As WeakReference
Set newWeakReference = gNewWeakReference(target)
End Function

Public Function OptionRightFromString(ByVal value As String) As OptionRights
OptionRightFromString = gOptionRightFromString(value)
End Function

Public Function OptionRightToString(ByVal value As OptionRights) As String
OptionRightToString = gOptionRightToString(value)
End Function

Public Function OrderActionFromString(ByVal value As String) As OrderActions
OrderActionFromString = gOrderActionFromString(value)
End Function

Public Function OrderActionToString(ByVal value As OrderActions) As String
OrderActionToString = gOrderActionToString(value)
End Function

Public Function OrderStatusToString(ByVal value As OrderStatuses) As String
OrderStatusToString = gOrderStatusToString(value)
End Function

Public Function OrderStopTriggerMethodToString(ByVal value As StopTriggerMethods) As String
OrderStopTriggerMethodToString = gOrderStopTriggerMethodToString(value)
End Function

Public Function OrderTIFToString(ByVal value As OrderTifs) As String
OrderTIFToString = gOrderTIFToString(value)
End Function

Public Function OrderTypeToString(ByVal value As OrderTypes) As String
OrderTypeToString = gOrderTypeToString(value)
End Function

Public Function SecTypeFromString(ByVal value As String) As SecurityTypes
SecTypeFromString = gSecTypeFromString(value)
End Function

Public Function SecTypeToShortString(ByVal value As SecurityTypes) As String
SecTypeToShortString = gSecTypeToShortString(value)
End Function

Public Function SecTypeToString(ByVal value As SecurityTypes) As String
SecTypeToString = gSecTypeToString(value)
End Function

Public Sub SortObjects(data() As SortEntryStruct)
gSortObjects data, LBound(data), UBound(data)
End Sub

'/**
' Converts a member of the StopOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pStopOrderType the StopOrderTypes value to be converted
' @ see
'
'*/
Public Function StopOrderTypeToOrderType( _
                ByVal pStopOrderType As TradeBuild.StopOrderTypes) As TradeBuild.OrderTypes
StopOrderTypeToOrderType = gStopOrderTypeToOrderType(pStopOrderType)
End Function

Public Function StopOrderTypeToString(ByVal value As StopOrderTypes)
StopOrderTypeToString = gStopOrderTypeToString(value)
End Function

'/**
' Converts a member of the TargetOrderTypes enumeration to the equivalent OrderTypes value.
'
' @return           the OrderTypes value corresponding to the parameter
' @param pTargetOrderType the TargetOrderTypes value to be converted
' @ see
'
'*/
Public Function TargetOrderTypeToOrderType( _
                ByVal pTargetOrderType As TradeBuild.TargetOrderTypes) As TradeBuild.OrderTypes
TargetOrderTypeToOrderType = gTargetOrderTypeToOrderType(pTargetOrderType)
End Function

Public Function TargetOrderTypeToString(ByVal value As TargetOrderTypes) As String
TargetOrderTypeToString = gTargetOrderTypeToString(value)
End Function

Public Function TickfileSpecifierToString(TickfileSpec As TradeBuild.TickfileSpecifier) As String
TickfileSpecifierToString = gTickfileSpecifierToString(TickfileSpec)
End Function

Public Function ToBytes(inString As String) As Byte()
ToBytes = gToBytes(inString)
End Function

Public Function ToHex(inAr() As Byte) As String
ToHex = gToHex(inAr)
End Function

'================================================================================
' Helper Functions
'================================================================================


