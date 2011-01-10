VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#48.0#0"; "TWControls10.ocx"
Begin VB.UserControl BarFormatterPicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   ScaleHeight     =   345
   ScaleWidth      =   3645
   Begin TWControls10.TWImageCombo BarFormattersCombo 
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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
BarFormattersCombo.Move 0, 0, UserControl.Width
UserControl.Height = BarFormattersCombo.Height
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub BarFormattersCombo_Click()
Dim lEntry As BarFormatterFactoryListEntry
Dim lBaseStudyConfig As StudyConfiguration
Dim lBarsValueConfig As StudyValueConfiguration

Const ProcName As String = "BarFormattersCombo_Click"
On Error GoTo Err

Set lBaseStudyConfig = mChartManager.BaseStudyConfiguration
Set lBarsValueConfig = lBaseStudyConfig.StudyValueConfigurations("Bar")

If BarFormattersCombo.SelectedItem.Index = 1 Then
    lBarsValueConfig.BarFormatterFactoryName = ""
    lBarsValueConfig.BarFormatterLibraryName = ""
Else
    lEntry = mBarFormatters(BarFormattersCombo.SelectedItem.Index - 2)
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
selectItem
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise(ByVal pChartMgr As ChartManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mChartManager = pChartMgr

loadCombo

selectItem

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function itemKey(ByVal pFormatterName As String, ByVal pLibraryName As String) As String
itemKey = pFormatterName & "  (" & pLibraryName & ")"
End Function

Private Sub loadCombo()
Dim i As Long
Dim lItemKey As String
Dim lMaxIndex As Long

Const ProcName As String = "loadCombo"
On Error GoTo Err

BarFormattersCombo.ComboItems.Clear

BarFormattersCombo.ComboItems.Add , NoBarFormatter, NoBarFormatter


mBarFormatters = GetAvailableBarFormatterFactories

lMaxIndex = -1
On Error Resume Next
lMaxIndex = UBound(mBarFormatters)
On Error GoTo Err

For i = 0 To lMaxIndex
    lItemKey = itemKey(mBarFormatters(i).name, mBarFormatters(i).LibraryName)
    BarFormattersCombo.ComboItems.Add , lItemKey, lItemKey
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub selectItem()
Dim lBarsValueConfig As StudyValueConfiguration

Const ProcName As String = "selectItem"
On Error GoTo Err

Set lBarsValueConfig = mChartManager.BaseStudyConfiguration.StudyValueConfigurations("Bar")

If lBarsValueConfig.BarFormatterFactoryName = "" Then
    BarFormattersCombo.ComboItems.item(NoBarFormatter).selected = True
Else
    BarFormattersCombo.ComboItems.item(itemKey(lBarsValueConfig.BarFormatterFactoryName, lBarsValueConfig.BarFormatterLibraryName)).selected = True
End If

BarFormattersCombo.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

