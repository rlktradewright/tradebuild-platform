VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#49.0#0"; "TWControls10.ocx"
Begin VB.UserControl BarFormatterPicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   ScaleHeight     =   345
   ScaleWidth      =   3645
   Begin TWControls10.TWImageCombo Combo1 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "BarFormatterPicker.ctx":0000
      Text            =   ""
   End
End
Attribute VB_Name = "BarFormatterPicker"
Attribute VB_GlobalNameSpace = False
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BarFormatterPicker"

Private Const NoBarFormatter                        As String = "(None)"

'@================================================================================
' Member variables
'@================================================================================

Private mBarFormatters()                            As BarFormatterFactoryListEntry

Private WithEvents mChartManager                    As ChartManager
Attribute mChartManager.VB_VarHelpID = -1
Private WithEvents mTradeBuildChart                 As TradeBuildChart
Attribute mTradeBuildChart.VB_VarHelpID = -1
Private WithEvents mMultiChart                      As MultiChart
Attribute mMultiChart.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

Const ProcName As String = "UserControl_ReadProperties"
On Error GoTo Err

Combo1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
Combo1.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
Combo1.ListWidth = PropBag.ReadProperty("ListWidth", 0)
Combo1.Locked = PropBag.ReadProperty("Locked", False)
Combo1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()
Combo1.Move 0, 0, UserControl.Width
UserControl.Height = Combo1.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

Const ProcName As String = "UserControl_WriteProperties"
On Error GoTo Err

Call PropBag.WriteProperty("BackColor", Combo1.BackColor, &H80000005)
Call PropBag.WriteProperty("CausesValidation", Combo1.CausesValidation, True)
Call PropBag.WriteProperty("Enabled", Combo1.Enabled, True)
Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000008)
Call PropBag.WriteProperty("ListWidth", Combo1.ListWidth, 0)
Call PropBag.WriteProperty("Locked", Combo1.Locked, False)
Call PropBag.WriteProperty("ToolTipText", Combo1.ToolTipText, "")
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub Combo1_Click()
Dim lEntry As BarFormatterFactoryListEntry
Dim lBaseStudyConfig As StudyConfiguration
Dim lBarsValueConfig As StudyValueConfiguration

Const ProcName As String = "Combo1_Click"
On Error GoTo Err

Set lBaseStudyConfig = mChartManager.BaseStudyConfiguration
Set lBarsValueConfig = lBaseStudyConfig.StudyValueConfigurations("Bar")

If Combo1.SelectedItem.index = 1 Then
    lBarsValueConfig.BarFormatterFactoryName = ""
    lBarsValueConfig.BarFormatterLibraryName = ""
Else
    lEntry = mBarFormatters(Combo1.SelectedItem.index - 2)
    lBarsValueConfig.BarFormatterFactoryName = lEntry.name
    lBarsValueConfig.BarFormatterLibraryName = lEntry.LibraryName
End If

mChartManager.BaseStudyConfiguration = lBaseStudyConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mChartManager Event Handlers
'@================================================================================

Private Sub mChartManager_BaseStudyConfigurationChanged(ByVal studyConfig As ChartUtils26.StudyConfiguration)
Const ProcName As String = "mChartManager_BaseStudyConfigurationChanged"
On Error GoTo Err

SelectItem

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mMultiChart Event Handlers
'@================================================================================

Private Sub mMultiChart_Change(ev As TWUtilities30.ChangeEventData)
Dim changeType As MultiChartChangeTypes
Const ProcName As String = "mMultiChart_Change"
Dim failpoint As String
On Error GoTo Err

changeType = ev.changeType
Select Case changeType
Case MultiChartSelectionChanged
    attachToCurrentChart
Case MultiChartAdd

Case MultiChartRemove

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTradeBuildChart Event Handlers
'@================================================================================

Private Sub mTradeBuildChart_StateChange(ev As TWUtilities30.StateChangeEventData)
Dim State As ChartStates
Const ProcName As String = "mTradeBuildChart_StateChange"
Dim failpoint As String
On Error GoTo Err

State = ev.State
Select Case State
Case ChartStateBlank

Case ChartStateCreated

Case ChartStateInitialised

Case ChartStateLoaded
    setChartManager
    SelectItem
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BackColor() As OLE_COLOR
Const ProcName As String = "BackColor"
On Error GoTo Err

    BackColor = Combo1.BackColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Const ProcName As String = "BackColor"
On Error GoTo Err

    Combo1.BackColor() = New_BackColor
    PropertyChanged "BackColor"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CausesValidation() As Boolean
Const ProcName As String = "CausesValidation"
On Error GoTo Err

    CausesValidation = Combo1.CausesValidation

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
Const ProcName As String = "CausesValidation"
On Error GoTo Err

    Combo1.CausesValidation() = New_CausesValidation
    PropertyChanged "CausesValidation"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
Dim failpoint As String
On Error GoTo Err

UserControl.Enabled = value
Combo1.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Font() As Font
Const ProcName As String = "Font"
On Error GoTo Err

    Set Font = Combo1.Font

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Set Font(ByVal New_Font As Font)
Const ProcName As String = "Font"
On Error GoTo Err

    Set Combo1.Font = New_Font
    PropertyChanged "Font"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Const ProcName As String = "ForeColor"
On Error GoTo Err

    ForeColor = Combo1.ForeColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Const ProcName As String = "ForeColor"
On Error GoTo Err

    Combo1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Locked() As Boolean
Const ProcName As String = "Locked"
On Error GoTo Err

    Locked = Combo1.Locked

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
Const ProcName As String = "Locked"
On Error GoTo Err

    Combo1.Locked() = New_Locked
    PropertyChanged "Locked"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ListWidth(ByVal value As Long)
Const ProcName As String = "ListWidth"
On Error GoTo Err

Combo1.ListWidth = value
PropertyChanged "ListWidth"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ListWidth() As Long
ListWidth = Combo1.ListWidth
End Property

Public Property Get ToolTipText() As String
Const ProcName As String = "ToolTipText"
On Error GoTo Err

    ToolTipText = Combo1.ToolTipText

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
Const ProcName As String = "ToolTipText"
On Error GoTo Err

    Combo1.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                Optional ByVal pChart As TradeBuildChart, _
                Optional ByVal pMultiChart As MultiChart)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pChart Is Nothing And pMultiChart Is Nothing Or _
    (Not pChart Is Nothing And Not pMultiChart Is Nothing) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Either a Chart or a Multichart (but not both) must be supplied"
End If

loadCombo

If Not pChart Is Nothing Then
    attachToChart pChart
ElseIf Not pMultiChart Is Nothing Then
    Set mMultiChart = pMultiChart
    attachToCurrentChart
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SelectBarFormatter(ByVal pBarFormatterName As String)
Const ProcName As String = "SelectBarFormatter"
On Error GoTo Err

Combo1.ComboItems.item(pBarFormatterName).Selected = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attachToChart(ByVal pChart As TradeBuildChart)
Const ProcName As String = "attachToChart"
Dim failpoint As String
On Error GoTo Err

Set mTradeBuildChart = pChart
If mTradeBuildChart.State = ChartStateLoaded Then
    setChartManager
    SelectItem
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub attachToCurrentChart()
Const ProcName As String = "attachToCurrentChart"
Dim failpoint As String
On Error GoTo Err

If mMultiChart.Count > 0 Then
    attachToChart mMultiChart.Chart
Else
    Set mTradeBuildChart = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function itemKey(ByVal pFormatterName As String, ByVal pLibraryName As String) As String
itemKey = pFormatterName & "  (" & pLibraryName & ")"
End Function

Private Sub loadCombo()
Dim i As Long
Dim lItemKey As String
Dim lMaxIndex As Long

Const ProcName As String = "loadCombo"
On Error GoTo Err

Combo1.ComboItems.Clear

Combo1.ComboItems.Add , NoBarFormatter, NoBarFormatter


mBarFormatters = GetAvailableBarFormatterFactories

lMaxIndex = -1
On Error Resume Next
lMaxIndex = UBound(mBarFormatters)
On Error GoTo Err

For i = 0 To lMaxIndex
    lItemKey = itemKey(mBarFormatters(i).name, mBarFormatters(i).LibraryName)
    Combo1.ComboItems.Add , lItemKey, lItemKey
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub SelectItem()
Dim lBarsValueConfig As StudyValueConfiguration

Const ProcName As String = "selectItem"
On Error GoTo Err

Set lBarsValueConfig = mChartManager.BaseStudyConfiguration.StudyValueConfigurations("Bar")

If lBarsValueConfig.BarFormatterFactoryName = "" Then
    SelectBarFormatter NoBarFormatter
Else
    SelectBarFormatter itemKey(lBarsValueConfig.BarFormatterFactoryName, lBarsValueConfig.BarFormatterLibraryName)
End If

Combo1.SelStart = 0
Combo1.SelLength = 0

Combo1.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setChartManager()
Set mChartManager = mTradeBuildChart.ChartManager
End Sub


