VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartSpecifier"
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "ChartSpecifier"

Private Const ConfigSettingCustomSessionEndTime         As String = "&CustomSessionEndTime"
Private Const ConfigSettingCustomSessionStartTime       As String = "&CustomSessionStartTime"
Private Const ConfigSettingFromTime                     As String = "&FromTime"
Private Const ConfigSettingToTime                       As String = "&ToTime"
Private Const ConfigSettingIncludeBarsOutsideSession    As String = "&IncludeBarsOutsideSession"
Private Const ConfigSettingInitialNumberOfBars          As String = "&InitialNumberOfBars"

'@================================================================================
' Member variables
'@================================================================================

Private mInitialNumberOfBars                As Long
Private mIncludeBarsOutsideSession          As Boolean
Private mFromTime                           As Date
Private mToTime                             As Date
Private mCustomSessionStartTime             As Date
Private mCustomSessionEndTime               As Date

Private mConfig                             As ConfigurationSection


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

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is mConfig Then Exit Property
Set mConfig = Value
If mConfig Is Nothing Then Exit Property

mConfig.SetSetting ConfigSettingFromTime, FormatTimestamp(mFromTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
mConfig.SetSetting ConfigSettingToTime, FormatTimestamp(mToTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
mConfig.SetSetting ConfigSettingInitialNumberOfBars, mInitialNumberOfBars
mConfig.SetBooleanSetting ConfigSettingIncludeBarsOutsideSession, mIncludeBarsOutsideSession
mConfig.SetSetting ConfigSettingCustomSessionStartTime, Format(mCustomSessionStartTime, "hh:mm")
mConfig.SetSetting ConfigSettingCustomSessionEndTime, Format(mCustomSessionEndTime, "hh:mm")

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CustomSessionEndTime() As Date
CustomSessionEndTime = mCustomSessionEndTime
End Property

Public Property Get CustomSessionStartTime() As Date
CustomSessionStartTime = mCustomSessionStartTime
End Property

Public Property Get FromTime() As Date
FromTime = mFromTime
End Property

Public Property Get ToTime() As Date
ToTime = mToTime
End Property

Public Property Get InitialNumberOfBars() As Long
Const ProcName As String = "InitialNumberOfBars"
On Error GoTo Err

InitialNumberOfBars = mInitialNumberOfBars

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeBarsOutsideSession() As Boolean
Const ProcName As String = "IncludeBarsOutsideSession"
On Error GoTo Err

IncludeBarsOutsideSession = mIncludeBarsOutsideSession

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pInitialNumberOfBars As Long, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pFromTime As Date, _
                ByVal pToTime As Date, _
                ByVal pCustomSessionStartTime As Date, _
                ByVal pCustomSessionEndTime As Date)
Const ProcName As String = "Initialise"
On Error GoTo Err

mInitialNumberOfBars = pInitialNumberOfBars
mIncludeBarsOutsideSession = pIncludeBarsOutsideSession
mFromTime = pFromTime
mToTime = pToTime
mCustomSessionStartTime = pCustomSessionStartTime
mCustomSessionEndTime = pCustomSessionEndTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim ar() As String
Dim i As Long

Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

mInitialNumberOfBars = mConfig.GetSetting(ConfigSettingInitialNumberOfBars, 200)
mIncludeBarsOutsideSession = CBool(mConfig.GetSetting(ConfigSettingIncludeBarsOutsideSession, "False"))
mFromTime = CDate(mConfig.GetSetting(ConfigSettingFromTime, CDate(0)))
mToTime = CDate(mConfig.GetSetting(ConfigSettingToTime, CDate(0)))
mCustomSessionStartTime = CDate(mConfig.GetSetting(ConfigSettingCustomSessionStartTime, CDate(0)))
mCustomSessionEndTime = CDate(mConfig.GetSetting(ConfigSettingCustomSessionEndTime, CDate(0)))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

