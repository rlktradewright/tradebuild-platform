VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestBarUtils"
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

Public Sub BarEndTime_SevenSeconds_0100()
Assert.AreEqual "2016-01-15 08:00:07.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_SevenSeconds_0200()
Assert.AreEqual "2016-01-15 17:29:55.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:29:48"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_SevenSeconds_0300()
Assert.AreEqual "2016-01-15 17:36:06.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:36:00"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_SevenSeconds_0400()
Assert.AreEqual "2016-01-15 13:01:32.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 13:01:25"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_SevenSeconds_0500()
Assert.AreEqual "2016-01-15 17:30:05.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_SevenSeconds_0600()
Assert.AreEqual "2016-01-15 16:15:05.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 16:14:59"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub





Public Sub BarEndTime_OneMinute_0100()
Assert.AreEqual "2016-01-15 08:01:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneMinute_0200()
Assert.AreEqual "2016-01-15 17:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneMinute_0300()
Assert.AreEqual "2016-01-15 17:37:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:36:00"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneMinute_0400()
Assert.AreEqual "2016-01-15 13:02:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 13:01:25"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneMinute_0500()
Assert.AreEqual "2016-01-15 17:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneMinute_0600()
Assert.AreEqual "2016-01-15 16:15:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 16:14:59"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub




Public Sub BarEndTime_ThirteenMinute_0100()
Assert.AreEqual "2016-01-15 08:13:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThirteenMinute_0110()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 07:59:00"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThirteenMinute_0200()
Assert.AreEqual "2016-01-15 17:32:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThirteenMinute_0300()
Assert.AreEqual "2016-01-15 17:45:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:36:00"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThirteenMinute_0400()
Assert.AreEqual "2016-01-15 13:05:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 13:01:25"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThirteenMinute_0500()
Assert.AreEqual "2016-01-15 17:35:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThirteenMinute_0600()
Assert.AreEqual "2016-01-15 16:20:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 16:14:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThirteenMinute_0700()
Assert.AreEqual "2016-01-15 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 16:29:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub





Public Sub BarEndTime_ThreeHour_0100()
Assert.AreEqual "2016-01-15 11:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThreeHour_0200()
Assert.AreEqual "2016-01-15 17:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 15:15:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThreeHour_0300()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 07:15:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThreeHour_0400()
Assert.AreEqual "2016-01-14 19:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("14/1/16 17:32:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_ThreeHour_0500()
Assert.AreEqual "2016-01-15 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 16:14:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub






Public Sub BarEndTime_OneDay_0100()
Assert.AreEqual "2016-01-18 08:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneDay_0110()
Assert.AreEqual "2016-02-18 08:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("17/2/16 11:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneDay_0120()
Assert.AreEqual "2016-02-16 08:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/2/16 11:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneDay_0200()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 07:50:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("08:00"), _
                                CDate("17:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneDay_0300()
Assert.AreEqual "2016-01-17 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 11:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneDay_0400()
Assert.AreEqual "2016-01-17 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("14/1/16 19:21:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneDay_0500()
Assert.AreEqual "2016-01-11 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("10/1/16 16:42:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_OneDay_0600()
Assert.AreEqual "2016-01-17 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 16:31:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub






Public Sub BarEndTime_TwoDay_0100()
Assert.AreEqual "2016-01-18 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("15/1/16 11:57:00"), _
                                GetTimePeriod(2, TimePeriodDay), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarEndTime_TwoDay_0200()
Assert.AreEqual "2016-01-10 16:30:00.000", _
                FormatTimestamp( _
                    BarEndTime(CDate("08/1/16 11:57:00"), _
                                GetTimePeriod(2, TimePeriodDay), _
                                CDate("16:30"), _
                                CDate("16:15")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub







Public Sub BarStartTime_SevenSeconds_0100()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_SevenSeconds_0200()
Assert.AreEqual "2016-01-15 17:29:48.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:29:48"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_SevenSeconds_0300()
Assert.AreEqual "2016-01-15 17:35:59.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:36:00"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_SevenSeconds_0400()
Assert.AreEqual "2016-01-15 13:01:25.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 13:01:25"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_SevenSeconds_0500()
Assert.AreEqual "2016-01-15 17:29:58.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_SevenSeconds_0600()
Assert.AreEqual "2016-01-15 16:14:58.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 16:14:59"), _
                                GetTimePeriod(7, TimePeriodSecond), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub





Public Sub BarStartTime_OneMinute_0100()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneMinute_0200()
Assert.AreEqual "2016-01-15 17:29:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneMinute_0300()
Assert.AreEqual "2016-01-15 17:36:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:36:00"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneMinute_0400()
Assert.AreEqual "2016-01-15 13:01:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 13:01:25"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneMinute_0500()
Assert.AreEqual "2016-01-15 17:29:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneMinute_0600()
Assert.AreEqual "2016-01-15 16:14:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 16:14:59"), _
                                GetTimePeriod(1, TimePeriodMinute), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub





Public Sub BarStartTime_ThirteenMinute_0100()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThirteenMinute_0110()
Assert.AreEqual "2016-01-15 07:50:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 07:59:00"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThirteenMinute_0200()
Assert.AreEqual "2016-01-15 17:19:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThirteenMinute_0300()
Assert.AreEqual "2016-01-15 17:32:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:36:00"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThirteenMinute_0400()
Assert.AreEqual "2016-01-15 12:52:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 13:01:25"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThirteenMinute_0500()
Assert.AreEqual "2016-01-15 17:22:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 17:29:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThirteenMinute_0600()
Assert.AreEqual "2016-01-15 16:07:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 16:14:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThirteenMinute_0700()
Assert.AreEqual "2016-01-15 16:20:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 16:29:59"), _
                                GetTimePeriod(13, TimePeriodMinute), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub





Public Sub BarStartTime_ThreeHour_0100()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThreeHour_0200()
Assert.AreEqual "2016-01-15 14:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 15:15:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThreeHour_0300()
Assert.AreEqual "2016-01-15 05:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 07:15:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThreeHour_0400()
Assert.AreEqual "2016-01-14 16:30:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("14/1/16 17:32:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_ThreeHour_0500()
Assert.AreEqual "2016-01-15 13:30:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 16:14:00"), _
                                GetTimePeriod(3, TimePeriodHour), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub






Public Sub BarStartTime_OneDay_0100()
Assert.AreEqual "2016-01-15 08:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneDay_0110()
Assert.AreEqual "2016-02-16 08:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("16/2/16 11:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneDay_0200()
Assert.AreEqual "2016-01-14 08:00:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 07:50:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneDay_0300()
Assert.AreEqual "2016-01-14 16:30:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 11:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneDay_0400()
Assert.AreEqual "2016-01-14 16:30:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("14/1/16 19:21:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneDay_0500()
Assert.AreEqual "2016-01-10 16:30:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("10/1/16 16:42:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub

Public Sub BarStartTime_OneDay_0600()
Assert.AreEqual "2016-01-14 16:30:00.000", _
                FormatTimestamp( _
                    BarStartTime(CDate("15/1/16 16:31:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                CDate("16:30")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub





Public Sub OffsetBarStartTime_OneDay_0100()
Assert.AreEqual "2016-01-18 08:00:00.000", _
                FormatTimestamp( _
                    OffsetBarStartTime(CDate("15/1/16 08:00:00"), _
                                GetTimePeriod(1, TimePeriodDay), _
                                1, _
                                CDate("08:00")), _
                    TimestampFormats.TimestampDateAndTimeISO8601)
End Sub


'@================================================================================
' Helper Functions
'@================================================================================





