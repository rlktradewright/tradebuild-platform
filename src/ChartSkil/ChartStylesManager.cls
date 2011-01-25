VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartStylesManager"
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

Implements Enumerable

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

Private Const ModuleName                            As String = "ChartStylesManager"

Private Const ConfigSectionChartStyle               As String = "ChartStyle"

'@================================================================================
' Member variables
'@================================================================================

Private mChartStyles                                As SortedDictionary

Private mDefaultStyle                               As ChartStyle

Private mConfig                                     As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mChartStyles = CreateSortedDictionary
Set mDefaultStyle = addDefaultStyle

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Set Enumerable_Enumerator = getEnumerator
End Property

Private Function Enumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
End Function

Private Sub Enumerable_Remove(cookie As Variant)
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
End Function

Private Property Get Enumerable_UpdateSequence() As Long
End Property

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

Dim lStyle As ChartStyle

If Value Is mConfig Then Exit Property
Set mConfig = Value

If mConfig Is Nothing Then Exit Property

For Each lStyle In mChartStyles
    addToConfig lStyle
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultStyle() As ChartStyle
Set DefaultStyle = mDefaultStyle
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal pName As String, _
                ByVal pBasedOn As ChartStyle, _
                Optional ByVal pDefaultRegionStyle As ChartRegionStyle, _
                Optional ByVal pXAxisRegionStyle As ChartRegionStyle, _
                Optional ByVal pDefaultYAxisRegionStyle As ChartRegionStyle, _
                Optional ByVal pCrosshairLineStyle As LineStyle, _
                Optional ByVal pXCursorTextStyle As TextStyle, _
                Optional ByVal pTemporary As Boolean = False) As ChartStyle
Const ProcName As String = "Add"
On Error GoTo Err

If pBasedOn Is Nothing Then Set pBasedOn = DefaultStyle

Set Add = New ChartStyle
Add.Initialise pName, pBasedOn, pDefaultRegionStyle, pXAxisRegionStyle, pDefaultYAxisRegionStyle, pCrosshairLineStyle, pXCursorTextStyle

If Not pTemporary Then
    addChartStyle Add
    addToConfig Add
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Public Function Contains(ByVal pName As String) As Boolean
Const ProcName As String = "Item"
On Error GoTo Err

Contains = mChartStyles.Contains(pName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mChartStyles.Count

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function Item(ByVal pName As String) As ChartStyle
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mChartStyles.Item(pName)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Dim cs As ConfigurationSection
Dim lStyle As ChartStyle

mChartStyles.Remove DefaultStyleName
Set mDefaultStyle = Nothing

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

For Each cs In mConfig
    Set lStyle = New ChartStyle
    lStyle.LoadFromConfig cs
    addChartStyle lStyle
Next

If Contains(DefaultStyleName) Then
    Set mDefaultStyle = Item(DefaultStyleName)
Else
    Set mDefaultStyle = addDefaultStyle
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = getEnumerator
End Function

Public Sub Remove(ByVal pName As String)
Dim lStyle As ChartStyle
Const ProcName As String = "Remove"
On Error GoTo Err

If pName = DefaultStyleName Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Default chart style cannot be removed"
End If

Set lStyle = Item(pName)
lStyle.RemoveFromConfig
mChartStyles.Remove pName

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addChartStyle( _
                ByVal pStyle As ChartStyle) As ChartStyle
Const ProcName As String = "addChartStyle"
On Error GoTo Err

Set addChartStyle = pStyle
mChartStyles.Add pStyle, pStyle.Name

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function addDefaultStyle() As ChartStyle
Const ProcName As String = "addDefaultStyle"
On Error GoTo Err

Set addDefaultStyle = addChartStyle(gDefaultChartStyle)
addToConfig addDefaultStyle

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub addToConfig(ByVal pStyle As ChartStyle)
Const ProcName As String = "addToConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

pStyle.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartStyle & "(" & pStyle.Name & ")")

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
                
Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Set getEnumerator = mChartStyles.Enumerator

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function


