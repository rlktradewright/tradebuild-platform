VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFormatters"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

Public Type BarFormatterFactoryListEntry
    Name                As String
    LibraryName         As String
End Type

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

Public Property Get BuiltInBarFormatterLibProgId() As String
BuiltInBarFormatterLibProgId = Globals.BuiltInBarFormatterLibProgId
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub SetDefaultBarFormatterLibraryConfig( _
                ByVal configdata As ConfigurationSection)
Const ProcName As String = "SetDefaultBarFormatterLibraryConfig"
On Error GoTo Err

On Error Resume Next
Dim currSLsList As ConfigurationSection
Set currSLsList = configdata.GetConfigurationSection(ConfigNameBarFormatterLibraries)
On Error GoTo Err

Assert currSLsList Is Nothing, "BarFormatter libraries list is not empty"

Set currSLsList = configdata.AddConfigurationSection(ConfigNameBarFormatterLibraries, , BarFormatterLibrariesRenderer)

Dim currSL As ConfigurationSection
Set currSL = currSLsList.AddConfigurationSection(ConfigNameBarFormatterLibrary & "(" & BuiltInBarFormatterLibName & ")")

currSL.SetAttribute AttributeNameEnabled, "True"
currSL.SetAttribute AttributeNameBarFormatterLibraryBuiltIn, "True"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




