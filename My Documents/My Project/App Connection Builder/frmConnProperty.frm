VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmConnProperty 
   Caption         =   "Connection Properties"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5130
   Icon            =   "frmConnProperty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   5130
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar SttBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   6405
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   556
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdButton 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   315
      Left            =   3930
      TabIndex        =   1
      Top             =   5970
      Width           =   1035
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5535
      Left            =   150
      TabIndex        =   0
      Top             =   300
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9763
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Connection Properties:"
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   90
      Width           =   1605
   End
End
Attribute VB_Name = "frmConnProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private RightMargin As Integer
Private BottomMargin As Integer
Private InitFormWidth As Integer

Private Sub cmdButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    InitFormWidth = Me.ScaleWidth
    BottomMargin = Me.ScaleHeight - (cmdButton.Top + cmdButton.Height)
    RightMargin = Me.ScaleWidth - (cmdButton.Left + cmdButton.Width)
    Call InitGrid
    Call MSHFlxAddItem(MSHFlexGrid1, "Connection String|" & CN.ActiveConnection.ConnectionString)
    For i = 0 To CN.ActiveConnection.Properties.Count - 1
        Call MSHFlxAddItem(MSHFlexGrid1, CN.ActiveConnection.Properties(i).Name & "(" & i & ")|" & CN.ActiveConnection.Properties(i))
    Next i
End Sub

Sub InitGrid()
    MSHFlexGrid1.Cols = 2
    Call MSHFlxSetColFitSize(MSHFlexGrid1, 35, 65)
    Call MSHFlxSetColCaption(MSHFlexGrid1, "Name", "Value")
    Call MSHFlxSetColAlign(MSHFlexGrid1, flexAlignLeftCenter, flexAlignLeftCenter)
    Call MSHFlxHighLight(MSHFlexGrid1)
End Sub

Private Sub Form_Resize()
    Dim frmDiff As Integer
    
    cmdButton.Left = Me.ScaleWidth - RightMargin - cmdButton.Width
    cmdButton.Top = Me.ScaleHeight - BottomMargin - cmdButton.Height
    
    MSHFlexGrid1.Width = Me.ScaleWidth - RightMargin - 100
    MSHFlexGrid1.Height = Me.ScaleHeight - BottomMargin - cmdButton.Height - 400
    
    Call InitGrid
    
End Sub

Private Sub MSHFlexGrid1_Click()
    SttBar.SimpleText = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.RowSel, 1)
End Sub
