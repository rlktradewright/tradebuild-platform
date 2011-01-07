VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StudyValueConfigurations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
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

Private Const ModuleName                As String = "StudyValueConfigurations"

Private Const ConfigSectionStudyValueConfig             As String = "StudyValueConfiguration"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStudyValueConfigurations As Collection

Private mConfig As ConfigurationSection

Private mBars As Bars

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStudyValueConfigurations = New Collection
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

Friend Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get count() As Long
Const ProcName As String = "count"
On Error GoTo Err

count = mStudyValueConfigurations.count

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal ValueName As String) As StudyValueConfiguration
Const ProcName As String = "Add"
On Error GoTo Err

Set Add = New StudyValueConfiguration
Add.ValueName = ValueName
mStudyValueConfigurations.Add Add, ValueName
storeStudyValueSettings Add

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Clone() As StudyValueConfigurations
Dim svc As StudyValueConfiguration
Dim newSvc As StudyValueConfiguration

Const ProcName As String = "Clone"
On Error GoTo Err

Set Clone = New StudyValueConfigurations

For Each svc In mStudyValueConfigurations
    newSvc = Clone.Add(svc.ValueName)
    newSvc.BarStyle = svc.BarStyle
    newSvc.ChartRegionName = svc.ChartRegionName
    newSvc.DataPointStyle = svc.DataPointStyle
    newSvc.IncludeInChart = svc.IncludeInChart
    newSvc.Layer = svc.Layer
    newSvc.LineStyle = svc.LineStyle
    newSvc.TextStyle = svc.TextStyle
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Item( _
                ByVal index As Variant) As StudyValueConfiguration
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mStudyValueConfigurations.Item(index)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub LoadFromConfig( _
                ByVal config As ConfigurationSection, _
                ByVal pBars As Bars)
Dim svcSect As ConfigurationSection
Dim svc As StudyValueConfiguration

Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mBars = pBars
Set mConfig = config
If mConfig Is Nothing Then Exit Sub

For Each svcSect In mConfig
    Set svc = New StudyValueConfiguration
    svc.LoadFromConfig svcSect
    mStudyValueConfigurations.Add svc, svc.ValueName
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = mStudyValueConfigurations.[_NewEnum]

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub storeSettings()
Dim svc As StudyValueConfiguration
Const ProcName As String = "storeSettings"
On Error GoTo Err

For Each svc In mStudyValueConfigurations
    storeStudyValueSettings svc
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub storeStudyValueSettings( _
                ByVal studyValueConfig As StudyValueConfiguration)
Const ProcName As String = "storeStudyValueSettings"
On Error GoTo Err

If Not mConfig Is Nothing Then
    studyValueConfig.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionStudyValueConfig & "(" & GenerateGUIDString & ")")
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



