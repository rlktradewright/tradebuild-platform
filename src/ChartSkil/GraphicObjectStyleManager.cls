VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicObjectStyleManager"
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

Private Const ModuleName                            As String = "GraphicObjectStyleManager"

Private Const ConfigSectionStyleTypes               As String = "StyleTypes"
Private Const ConfigSectionStyle                    As String = "Style"

Private Const ConfigAttributeProgId                 As String = "ProgId"

'@================================================================================
' Member variables
'@================================================================================

Private mDict As SortedDictionary

Private mConfig As ConfigurationSection

Private mStyleTypesSection As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mDict = CreateSortedDictionary
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

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)

Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Not setConfig(Value) Then Exit Property

Dim lStyleType
For Each lStyleType In StyleTypeEnumerator
    addStylesToConfig lStyleType
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get EnumeratorForStyleType( _
                ByVal pStyleType As String) As Enumerator
Dim lCol As SortedDictionary
Const ProcName As String = "EnumeratorForStyleType"
On Error GoTo Err

Set lCol = mDict.Item(pStyleType)
Set EnumeratorForStyleType = lCol.Enumerator

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get StyleTypeEnumerator() As Enumerator
Const ProcName As String = "StyleTypeEnumerator"
On Error GoTo Err

Set StyleTypeEnumerator = mDict.KeyEnumerator

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal pStyle As IGraphicObjectStyle) As IGraphicObjectStyle
Dim lCol As SortedDictionary
Const ProcName As String = "Add"
On Error GoTo Err

Set Add = addStyle(pStyle)
addStyleToConfig pStyle

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Clear()
mDict.Clear
Set mConfig = Nothing
Set mStyleTypesSection = Nothing
End Function

Public Function Item( _
                ByVal pName As String, _
                ByVal pStyleType As String) As IGraphicObjectStyle
Dim lCol As SortedDictionary

Const ProcName As String = "Item"
On Error GoTo Err

Set lCol = mDict.Item(pStyleType)
Set Item = lCol.Item(pName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

If Not setConfig(pConfig) Then Exit Sub

Dim lStyleTypeCs As ConfigurationSection
Dim lStyleCs As ConfigurationSection

For Each lStyleTypeCs In mStyleTypesSection
    For Each lStyleCs In lStyleTypeCs
        Dim lStyle As IGraphicObjectStyle
        Set lStyle = CreateObject(lStyleCs.GetAttribute(ConfigAttributeProgId))
        lStyle.LoadFromConfig lStyleCs
        addStyle lStyle
    Next
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Remove( _
                ByVal pStyle As IGraphicObjectStyle)
Dim lCol As SortedDictionary
Const ProcName As String = "Remove"
On Error GoTo Err

Set lCol = mDict.Item(pStyle.StyleType)
lCol.Remove pStyle.Name

removeStyleFromConfig pStyle
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addStyle( _
                ByVal pStyle As IGraphicObjectStyle) As IGraphicObjectStyle
Dim lCol As SortedDictionary
Const ProcName As String = "addStyle"
On Error GoTo Err

If mDict.Contains(pStyle.StyleType) Then
    Set lCol = mDict.Item(pStyle.StyleType)
Else
    Set lCol = CreateSortedDictionary
    mDict.Add lCol, pStyle.StyleType
End If
lCol.Add pStyle, pStyle.Name
Set addStyle = pStyle

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub addStyleToConfig(ByVal pStyle As IGraphicObjectStyle)
Const ProcName As String = "addStyleToConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

Dim cs As ConfigurationSection
Set cs = addStyleTypeToConfig(pStyle.StyleType).AddConfigurationSection(ConfigSectionStyle & "(" & pStyle.Name & ")")
cs.SetAttribute ConfigAttributeProgId, pStyle.ProgId
pStyle.ConfigurationSection = cs

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addStylesToConfig(ByVal pStyleType As String)
Dim lStyle As IGraphicObjectStyle

Const ProcName As String = "addStylesToConfig"
On Error GoTo Err

For Each lStyle In EnumeratorForStyleType(pStyleType)
    addStyleToConfig lStyle
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function addStyleTypeToConfig(ByVal pStyleType As String) As ConfigurationSection
Const ProcName As String = "addStyleTypeToConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Function

Set addStyleTypeToConfig = mStyleTypesSection.AddConfigurationSection(ConfigSectionStyleTypes & "(" & pStyleType & ")")

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub removeStyleFromConfig(ByVal pStyle As IGraphicObjectStyle)
Const ProcName As String = "removeStyleFromConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

mStyleTypesSection.GetConfigurationSection(ConfigSectionStyleTypes & "(" & pStyle.StyleType & ")").RemoveConfigurationSection (ConfigSectionStyle & "(" & pStyle.Name & ")")

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


Private Function setConfig(ByVal Value As ConfigurationSection) As Boolean
Const ProcName As String = "setConfig"
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    setConfig = False
Else
    Set mConfig = Value
    Set mStyleTypesSection = mConfig.AddPrivateConfigurationSection(ConfigSectionStyleTypes)
    setConfig = True
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function
