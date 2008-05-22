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

Private mContractsConfig                    As ConfigItem
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
Dim ci As ConfigItem
Set ci = Node.Tag
ci.setAttribute AttributeNameEnabled, IIf(Node.Checked, "True", "False")
End Sub

Private Sub ContractsTV_NodeClick(ByVal Node As MSComctlLib.Node)
If Not mReadOnly Then EditButton.enabled = True
If Not mReadOnly Then RemoveButton.enabled = True
End Sub

Private Sub EditButton_Click()
editItem
End Sub

Private Sub RemoveButton_Click()
mContractsConfig.childItems.Remove ContractsTV.SelectedItem.Tag
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
Dim ci As ConfigItem

If mActionAdd Then
    Set ci = addConfigItem
    updateConfigItem ci, contractSpec, enabled, writeBidAskBars, includeMktDepth
    updateListItem addListItem(ci)
Else
    Set ci = ContractsTV.SelectedItem.Tag
    updateConfigItem ci, contractSpec, enabled, writeBidAskBars, includeMktDepth
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
                ByVal contractsConfig As ConfigItem, _
                ByVal readonly As Boolean)
Dim contractConfig As ConfigItem

mReadOnly = readonly
If mReadOnly Then AddButton.enabled = False

Set mContractsConfig = contractsConfig

ContractsTV.Nodes.Clear

For Each contractConfig In mContractsConfig.childItems
    updateListItem addListItem(contractConfig)
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addConfigItem() As ConfigItem
Set addConfigItem = mContractsConfig.childItems.AddItem(ConfigNameContract)
addConfigItem.childItems.AddItem ConfigNameContractSpecifier
End Function

Private Function addListItem( _
                ByVal contractCi As ConfigItem) As Node
Dim n As Node
Set n = ContractsTV.Nodes.Add
Set n.Tag = contractCi
Set addListItem = n
End Function

Private Function configItemToContractSpec( _
                ByVal contractCi As ConfigItem) As contractSpecifier
Dim localSymbol As String
Dim symbol As String
Dim exchange As String
Dim sectype As SecurityTypes
Dim currencyCode As String
Dim expiry As String
Dim strikePrice As Double
Dim optRight As OptionRights

With contractCi.childItems.Item(ConfigNameContractSpecifier)
    localSymbol = .getDefaultableAttribute(AttributeNameLocalSYmbol, "")
    symbol = .getDefaultableAttribute(AttributeNameSymbol, "")
    exchange = .getDefaultableAttribute(AttributeNameExchange, "")
    sectype = SecTypeFromString(.getDefaultableAttribute(AttributeNameSecType, ""))
    currencyCode = .getDefaultableAttribute(AttributeNameCurrency, "")
    expiry = .getDefaultableAttribute(AttributeNameExpiry, "")
    strikePrice = CDbl("0" & .getDefaultableAttribute(AttributeNameStrikePrice, "0.0"))
    optRight = OptionRightFromString(.getDefaultableAttribute(AttributeNameRight, ""))
    
    Set configItemToContractSpec = CreateContractSpecifier(localSymbol, _
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
Dim ci As ConfigItem
mActionAdd = False
Set ci = ContractsTV.SelectedItem.Tag
showContractSpecForm configItemToContractSpec(ci), _
                       CBool(ci.getDefaultableAttribute(AttributeNameEnabled, "False")), _
                       CBool(ci.getDefaultableAttribute(AttributeNameBidAskBars, "False")), _
                       CBool(ci.getDefaultableAttribute(AttributeNameIncludeMktDepth, "False"))
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

Private Sub updateConfigItem( _
                ByVal contractCi As ConfigItem, _
                ByVal cs As contractSpecifier, _
                ByVal enabled As Boolean, _
                ByVal writeBidAskBars As Boolean, _
                ByVal includeMktDepth As Boolean)
contractCi.setAttribute AttributeNameEnabled, IIf(enabled, "True", "False")
contractCi.setAttribute AttributeNameBidAskBars, IIf(writeBidAskBars, "True", "False")
contractCi.setAttribute AttributeNameIncludeMktDepth, IIf(includeMktDepth, "True", "False")
With contractCi.childItems.Item(ConfigNameContractSpecifier)
    .setAttribute AttributeNameLocalSYmbol, cs.localSymbol
    .setAttribute AttributeNameSymbol, cs.symbol
    .setAttribute AttributeNameExchange, cs.exchange
    .setAttribute AttributeNameSecType, SecTypeToString(cs.sectype)
    .setAttribute AttributeNameCurrency, cs.currencyCode
    .setAttribute AttributeNameExpiry, cs.expiry
    .setAttribute AttributeNameStrikePrice, cs.strike
    .setAttribute AttributeNameRight, OptionRightToString(cs.Right)
End With
                
End Sub

Private Sub updateListItem( _
                ByVal pNode As Node)
Dim contractCi As ConfigItem
Set contractCi = pNode.Tag
pNode.Text = configItemToContractSpec(contractCi).toString & _
                                    IIf(CBool(contractCi.getDefaultableAttribute(AttributeNameBidAskBars, "False")), _
                                        "Bid/Ask bars;", _
                                        "") & _
                                    IIf(CBool(contractCi.getDefaultableAttribute(AttributeNameIncludeMktDepth, "False")), _
                                        "Mkt depth;", _
                                        "")
pNode.Checked = CBool(contractCi.getDefaultableAttribute(AttributeNameEnabled, "False"))
End Sub

