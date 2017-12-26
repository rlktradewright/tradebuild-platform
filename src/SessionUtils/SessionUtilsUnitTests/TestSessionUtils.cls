VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestSessionUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Setup Functions
'@================================================================================

Public Sub FixtureSetup()
End Sub

Public Sub FixtureTeardown()
End Sub

Public Sub Setup()
End Sub

Public Sub TearDown()
LogMessage "Test Completed -----------------------------------------------"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Test Methods
'@================================================================================

Public Sub GetOffsetSessionTimes_0100()
Assert.AreEqual "2016-01-18 08:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("15/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0110()
Assert.AreEqual "2016-01-18 17:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("15/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0200()
Assert.AreEqual "2016-01-18 08:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("16/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0210()
Assert.AreEqual "2016-01-18 17:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("16/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0300()
Assert.AreEqual "2016-01-19 08:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0310()
Assert.AreEqual "2016-01-19 17:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0400()
Assert.AreEqual "2016-01-22 08:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("21/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0410()
Assert.AreEqual "2016-01-22 17:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("21/1/16 08:00:00"), _
                                1, _
                                CDate("08:00"), _
                                CDate("17:30")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub




Public Sub GetOffsetSessionTimes_0500()
Assert.AreEqual "2016-01-17 16:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("15/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0510()
Assert.AreEqual "2016-01-18 16:15:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("15/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0600()
Assert.AreEqual "2016-01-17 16:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("16/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0610()
Assert.AreEqual "2016-01-18 16:15:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("16/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0700()
Assert.AreEqual "2016-01-18 16:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0710()
Assert.AreEqual "2016-01-19 16:15:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0800()
Assert.AreEqual "2016-01-21 16:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("21/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_0810()
Assert.AreEqual "2016-01-22 16:15:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("21/1/16 08:00:00"), _
                                1, _
                                CDate("16:30"), _
                                CDate("16:15")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub





Public Sub GetOffsetSessionTimes_1100()
Assert.AreEqual "2016-01-18 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("15/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1110()
Assert.AreEqual "2016-01-19 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("15/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1200()
Assert.AreEqual "2016-01-18 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("16/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1210()
Assert.AreEqual "2016-01-19 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("16/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1300()
Assert.AreEqual "2016-01-19 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1310()
Assert.AreEqual "2016-01-20 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1400()
Assert.AreEqual "2016-01-22 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("21/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1410()
Assert.AreEqual "2016-01-23 00:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("21/1/16 08:00:00"), _
                                1, _
                                CDate("00:00"), _
                                CDate("00:00")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub




Public Sub GetOffsetSessionTimes_1500()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                -1, _
                                CDate("08:00"), _
                                CDate("17:30")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1510()
Assert.AreEqual "2016-01-15 17:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("18/1/16 08:00:00"), _
                                -1, _
                                CDate("08:00"), _
                                CDate("17:30")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1600()
Assert.AreEqual "2016-12-30 08:00:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("02/1/17 08:00:00"), _
                                -1, _
                                CDate("08:00"), _
                                CDate("17:30")).StartTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub GetOffsetSessionTimes_1610()
Assert.AreEqual "2016-12-30 17:30:00.000", _
                FormatTimestamp( _
                    GetOffsetSessionTimes(CDate("02/1/17 08:00:00"), _
                                -1, _
                                CDate("08:00"), _
                                CDate("17:30")).EndTime, _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





