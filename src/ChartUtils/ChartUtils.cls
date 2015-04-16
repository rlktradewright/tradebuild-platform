VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

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


Private Const ModuleName                As String = "ChartUtils"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type RegionEntry
    Region              As ChartRegion
    usageCount          As Long
End Type

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

Public Property Get ChartRegionNameCustom() As String
ChartRegionNameCustom = RegionNameCustom
End Property

Public Property Get ChartRegionNameDefault() As String
ChartRegionNameDefault = RegionNameDefault
End Property

Public Property Get ChartRegionNamePrice() As String
ChartRegionNamePrice = RegionNamePrice
End Property

Public Property Get ChartRegionNameVolume() As String
ChartRegionNameVolume = RegionNameVolume
End Property

Public Property Get ChartRegionNameUnderlying() As String
ChartRegionNameUnderlying = RegionNameUnderlying
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateAbstractChartTool( _
                ByVal pChart As ChartController, _
                ByVal tool As IChartTool, _
                Optional ByVal pointerIcon As IPictureDisp) As AbstractChartTool
Const ProcName As String = "CreateAbstractChartTool"
On Error GoTo Err

Set CreateAbstractChartTool = New AbstractChartTool
CreateAbstractChartTool.Initialise pChart, tool, pointerIcon

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateChartManager( _
                ByVal pChart As ChartController, _
                ByVal pStudyManager As StudyManager, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pIncludeBarsOutsideSession As Boolean) As ChartManager
Const ProcName As String = "CreateChartManager"
On Error GoTo Err

AssertArgument Not pChart Is Nothing, "pChart Is Nothing"
AssertArgument Not pStudyManager Is Nothing, "pStudyManager Is Nothing"

Set CreateChartManager = New ChartManager
CreateChartManager.Initialise pChart, pStudyManager, pBarFormatterLibManager, pIncludeBarsOutsideSession

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateChartSpecifier( _
                Optional ByVal pInitialNumberOfBars As Long = 200, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean = False, _
                Optional ByVal pFromTime As Date, _
                Optional ByVal pTotime As Date) As ChartSpecifier
Const ProcName As String = "CreateChartSpecifier"
On Error GoTo Err

Set CreateChartSpecifier = New ChartSpecifier
CreateChartSpecifier.Initialise pInitialNumberOfBars, _
                pIncludeBarsOutsideSession, _
                pFromTime, _
                pTotime

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadChartSpecifierFromConfig( _
                ByVal config As ConfigurationSection) As ChartSpecifier
Const ProcName As String = "LoadChartSpecifierFromConfig"
On Error GoTo Err

AssertArgument Not config Is Nothing, "config argument is Nothing"


Set LoadChartSpecifierFromConfig = New ChartSpecifier
LoadChartSpecifierFromConfig.LoadFromConfig config

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LoadDefaultStudyConfigurationsFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadDefaultStudyConfigurationsFromConfig"
On Error GoTo Err

gLoadDefaultStudyConfigurationsFromConfig config

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetDefaultStudyConfiguration( _
                ByVal Value As StudyConfiguration)
Const ProcName As String = "SetDefaultStudyConfiguration"
On Error GoTo Err

gSetDefaultStudyConfiguration Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


