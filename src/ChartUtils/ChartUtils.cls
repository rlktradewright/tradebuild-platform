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
Dim failpoint As String
On Error GoTo Err

Set CreateAbstractChartTool = New AbstractChartTool
CreateAbstractChartTool.Initialise pChart, tool, pointerIcon

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function CreateChartManager( _
                ByVal studyManager As studyManager, _
                ByVal pChart As ChartController) As ChartManager
Const ProcName As String = "CreateChartManager"
Dim failpoint As String
On Error GoTo Err

Set CreateChartManager = New ChartManager
CreateChartManager.Initialise studyManager, pChart

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function CreateChartSpecifier( _
                Optional ByVal pInitialNumberOfBars As Long = 200, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean = False, _
                Optional ByVal pFromTime As Date, _
                Optional ByVal pTotime As Date) As ChartSpecifier
Const ProcName As String = "CreateChartSpecifier"
Dim failpoint As String
On Error GoTo Err

Set CreateChartSpecifier = New ChartSpecifier
CreateChartSpecifier.Initialise pInitialNumberOfBars, _
                pIncludeBarsOutsideSession, _
                pFromTime, _
                pTotime

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function GetDefaultStudyConfiguration( _
                ByVal Name As String, _
                ByVal studyLibName As String) As StudyConfiguration
Const ProcName As String = "GetDefaultStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

Set GetDefaultStudyConfiguration = gGetDefaultStudyConfiguration(Name, studyLibName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function LoadChartSpecifierFromConfig( _
                ByVal config As ConfigurationSection) As ChartSpecifier
Const ProcName As String = "LoadChartSpecifierFromConfig"
Dim failpoint As String
On Error GoTo Err

If config Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "config argument is Nothing"
End If

Set LoadChartSpecifierFromConfig = New ChartSpecifier
LoadChartSpecifierFromConfig.LoadFromConfig config

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub LoadDefaultStudyConfigurationsFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadDefaultStudyConfigurationsFromConfig"
Dim failpoint As String
On Error GoTo Err

gLoadDefaultStudyConfigurationsFromConfig config

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetDefaultStudyConfiguration( _
                ByVal value As StudyConfiguration)
Const ProcName As String = "SetDefaultStudyConfiguration"
Dim failpoint As String
On Error GoTo Err

gSetDefaultStudyConfiguration value

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


