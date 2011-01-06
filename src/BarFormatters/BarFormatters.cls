VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFormatters"
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

Private Const ModuleName                            As String = "BarFormatters"

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

Public Function AddBarFormatterLibrary( _
                ByVal ProgId As String, _
                ByVal Enabled As Boolean, _
                Optional ByVal Name As String) As BarFormatterLibrary
Const ProcName As String = "AddBarFormatterLibrary"
On Error GoTo Err

If Name = "" Then Name = GenerateGUIDString
Set AddBarFormatterLibrary = Globals.gBarFormatterLibManager.AddBarFormatterLibrary( _
                                ProgId, _
                                Enabled, _
                                Name)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function GetAvailableBarFormatterFactories() As BarFormatterFactoryListEntry()
GetAvailableBarFormatterFactories = gBarFormatterLibManager.AvailableBarFormatterFactories
End Function

Public Sub LoadBarFormatterLibraryConfiguration( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadBarFormatterLibraryConfiguration"
On Error GoTo Err

Globals.gBarFormatterLibManager.LoadBarFormatterLibraryConfiguration config

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveBarFormatterLibrary( _
                ByVal library As BarFormatterLibrary)
Const ProcName As String = "RemoveBarFormatterLibrary"
On Error GoTo Err

gBarFormatterLibManager.Remove library

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveAllBarFormatterLibraries()
Const ProcName As String = "RemoveAllBarFormatterLibraries"
On Error GoTo Err

gBarFormatterLibManager.RemoveAll

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SetDefaultBarFormatterLibraryConfig( _
                ByVal configdata As ConfigurationSection)
Dim currSLsList As ConfigurationSection
Dim currSL As ConfigurationSection

Const ProcName As String = "SetDefaultBarFormatterLibraryConfig"
On Error GoTo Err

On Error Resume Next
Set currSLsList = configdata.GetConfigurationSection(ConfigNameBarFormatterLibraries)
On Error GoTo Err

If Not currSLsList Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "BarFormatter libraries list is not empty"
End If

Set currSLsList = configdata.AddConfigurationSection(ConfigNameBarFormatterLibraries, , BarFormatterLibrariesRenderer)

Set currSL = currSLsList.AddConfigurationSection(ConfigNameBarFormatterLibrary & "(" & BuiltInBarFormatterLibName & ")")

currSL.SetAttribute AttributeNameEnabled, "True"
currSL.SetAttribute AttributeNameBarFormatterLibraryBuiltIn, "True"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




