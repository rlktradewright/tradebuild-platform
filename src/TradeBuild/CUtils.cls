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

Public Function FormatTimestamp(ByVal Timestamp As Date, _
                                Optional ByVal formatOption As TimestampFormats = TimestampDateAndTime, _
                                Optional ByVal formatString As String = "yyyymmddhhnnss") As String
gFormatTimestamp Timestamp, formatOption, formatString
End Function

Public Function GenerateGUID() As GUID
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

Public Function GUIDToString(ByRef pGUID As GUID) As String
GUIDToString = gGUIDToString(pGUID)
End Function

Public Sub SortObjects(data() As SortEntry)
gSortObjects data, LBound(data), UBound(data)
End Sub

Public Function ToBytes(inString As String) As Byte()
ToBytes = gToBytes(inString)
End Function

Public Function ToHex(inAr() As Byte) As String
ToHex = gToHex(inAr)
End Function


