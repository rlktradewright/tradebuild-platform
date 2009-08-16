VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ContractsConfigurer 
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   ScaleHeight     =   4305
   ScaleWidth      =   7800
   Begin MSComctlLib.TreeView ContractsTV 
      Height          =   3735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   6588
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   0
   End
   Begin VB.CommandButton RemoveButton 
      Caption         =   "&Delete"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      ToolTipText     =   "Delete"
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton EditButton 
      Caption         =   "&Edit"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6600
      Picture         =   "ContractsConfigurer.ctx":0000
      TabIndex        =   1
      ToolTipText     =   "Move up"
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "&Add"
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      ToolTipText     =   "Add new"
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape OutlineBox 
      Height          =   4000
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "ContractsConfigurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ProjectName                   As String = "DataCollector26"
Private Const ModuleName                    As String = "ContractsConfigurer"

'@================================================================================
' Member variables
'@================================================================================

Private mContractsConfig                    As ConfigurationSection
Private WithEvents mContractSpecForm        As fContractSpec
Attribute mContractSpecForm.VB_VarHelpID = -1

Private mActionAdd                          As Boolean

Private mReadOnly                           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Width = OutlineBox.Width
UserControl.Height = OutlineBox.Height
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddButton_Click()
mActionAdd = True
showContractSpecForm Nothing, True, False, False
End Sub

Private Sub ContractsTV_DblClick()
If Not ContractsTV.SelectedItem Is Nothing Then editItem
End Sub

Private Sub ContractsTV_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim cs As ConfigurationSection
Set cs = Node.Tag
cs.setAttribute AttributeNameEnabled, IIf(Node.Checked, "True", "False")
End Sub

Private Sub ContractsTV_NodeClick(ByVal Node As MSComctlLib.Node)
If Not mReadOnly Then EditButton.enabled = True
If Not mReadOnly Then RemoveButton.enabled = True
End Sub

Private Sub EditButton_Click()
editItem
End Sub

Private Sub RemoveButton_Click()
mContractsConfig.RemoveConfigurationSection ContractsTV.SelectedItem.Tag
ContractsTV.Nodes.Remove ContractsTV.SelectedItem.index
End Sub

'@================================================================================
' mContractSpecForm Event Handlers
'@================================================================================

Private Sub mContractSpecForm_ContractSpecReady( _
                ByVal contractSpec As ContractUtils26.contractSpecifier, _
                ByVal enabled As Boolean, _
                ByVal writeBidAskBars As Boolean, _
                ByVal includeMktDepth As Boolean)
Dim cs As ConfigurationSection

If mActionAdd Then
    Set cs = addConfigurationSection
    updateConfigurationSection cs, contractSpec, enabled, writeBidAskBars, includeMktDepth
    updateListItem addListItem(cs)
Else
    Set cs = ContractsTV.SelectedItem.Tag
    updateConfigurationSection cs, contractSpec, enabled, writeBidAskBars, includeMktDepth
    updateListItem ContractsTV.SelectedItem
    Unload mContractSpecForm
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub initialise( _
                ByVal contractsConfig As ConfigurationSection, _
                ByVal readonly As Boolean)
Dim contractConfig As ConfigurationSection

mReadOnly = readonly
If mReadOnly Then AddButton.enabled = False

Set mContractsConfig = contractsConfig

ContractsTV.Nodes.Clear

For Each contractConfig In mContractsConfig
    updateListItem addListItem(contractConfig)
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addConfigurationSection() As ConfigurationSection
Set addConfigurationSection = mContractsConfig.addConfigurationSection(ConfigSectionContract & "(" & GenerateGUIDString & ")")
addConfigurationSection.addConfigurationSection ConfigSectionContractspecifier
End Function

Private Function addListItem( _
                ByVal contractCS As ConfigurationSection) As Node
Dim n As Node
Set n = ContractsTV.Nodes.Add
Set n.Tag = contractCS
Set addListItem = n
End Function

Private Function ConfigurationSectionToContractSpec( _
                ByVal contractCS As ConfigurationSection) As contractSpecifier
Dim localSymbol As String
Dim symbol As String
Dim exchange As String
Dim sectype As SecurityTypes
Dim currencyCode As String
Dim expiry As String
Dim strikePrice As Double
Dim optRight As OptionRights

With contractCS
    localSymbol = .GetSetting(ConfigSettingContractSpecLocalSYmbol, "")
    symbol = .GetSetting(ConfigSettingContractSpecSymbol, "")
    exchange = .GetSetting(ConfigSettingContractSpecExchange, "")
    sectype = SecTypeFromString(.GetSetting(ConfigSettingContractSpecSecType, ""))
    currencyCode = .GetSetting(ConfigSettingContractSpecCurrency, "")
    expiry = .GetSetting(ConfigSettingContractSpecExpiry, "")
    strikePrice = CDbl("0" & .GetSetting(ConfigSettingContractSpecStrikePrice, "0.0"))
    optRight = OptionRightFromString(.GetSetting(ConfigSettingContractSpecRight, ""))
    
    Set ConfigurationSectionToContractSpec = CreateContractSpecifier(localSymbol, _
                                                            symbol, _
                                                            exchange, _
                                                            sectype, _
                                                            currencyCode, _
                                                            expiry, _
                                                            strikePrice, _
                                                            optRight)
End With
                
End Function

Private Sub editItem()
Dim cs As ConfigurationSection
mActionAdd = False
Set cs = ContractsTV.SelectedItem.Tag
showContractSpecForm ConfigurationSectionToContractSpec(cs), _
                       CBool(cs.getAttribute(AttributeNameEnabled, "False")), _
                       CBool(cs.getAttribute(AttributeNameBidAskBars, "False")), _
                       CBool(cs.getAttribute(AttributeNameIncludeMktDepth, "False"))
End Sub

Private Sub showContractSpecForm( _
                ByVal contractSpec As contractSpecifier, _
                ByVal enabled As Boolean, _
                ByVal writeBidAskBars As Boolean, _
                ByVal includeMktDepth As Boolean)
Set mContractSpecForm = New fContractSpec
mContractSpecForm.initialise contractSpec, enabled, writeBidAskBars, includeMktDepth
mContractSpecForm.Show vbModal
End Sub

Private Sub updateConfigurationSection( _
                ByVal contractCS As ConfigurationSection, _
                ByVal contractSpec As contractSpecifier, _
                ByVal enabled As Boolean, _
                ByVal writeBidAskBars As Boolean, _
                ByVal includeMktDepth As Boolean)
contractCS.setAttribute AttributeNameEnabled, IIf(enabled, "True", "False")
contractCS.setAttribute AttributeNameBidAskBars, IIf(writeBidAskBars, "True", "False")
contractCS.setAttribute AttributeNameIncludeMktDepth, IIf(includeMktDepth, "True", "False")
With contractCS
    .SetSetting ConfigSettingContractSpecLocalSYmbol, contractSpec.localSymbol
    .SetSetting ConfigSettingContractSpecSymbol, contractSpec.symbol
    .SetSetting ConfigSettingContractSpecExchange, contractSpec.exchange
    .SetSetting ConfigSettingContractSpecSecType, SecTypeToString(contractSpec.sectype)
    .SetSetting ConfigSettingContractSpecCurrency, contractSpec.currencyCode
    .SetSetting ConfigSettingContractSpecExpiry, contractSpec.expiry
    .SetSetting ConfigSettingContractSpecStrikePrice, contractSpec.strike
    .SetSetting ConfigSettingContractSpecRight, OptionRightToString(contractSpec.Right)
End With
                
End Sub

Private Sub updateListItem( _
                ByVal pNode As Node)
Dim contractCS As ConfigurationSection
Set contractCS = pNode.Tag
pNode.Text = ConfigurationSectionToContractSpec(contractCS).toString & _
                                    IIf(CBool(contractCS.getAttribute(AttributeNameBidAskBars, "False")), _
                                        "Bid/Ask bars;", _
                                        "") & _
                                    IIf(CBool(contractCS.getAttribute(AttributeNameIncludeMktDepth, "False")), _
                                        "Mkt depth;", _
                                        "")
pNode.Checked = CBool(contractCS.getAttribute(AttributeNameEnabled, "False"))
End Sub

