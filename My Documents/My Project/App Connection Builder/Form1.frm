VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Database Connection Builder"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Connect"
      Height          =   315
      Left            =   5880
      TabIndex        =   3
      Top             =   1020
      Width           =   1365
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Conn. Properties"
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   1020
      Width           =   1365
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Exit"
      Height          =   315
      Index           =   1
      Left            =   7350
      TabIndex        =   7
      Top             =   4560
      Width           =   1035
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "About"
      Height          =   315
      Index           =   0
      Left            =   7350
      TabIndex        =   6
      Top             =   4140
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Build Connection"
      Height          =   315
      Left            =   4440
      TabIndex        =   2
      Top             =   1020
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Query"
      Height          =   315
      Left            =   5880
      TabIndex        =   5
      Top             =   2610
      Width           =   1365
   End
   Begin MSComctlLib.StatusBar SttBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   8
      Top             =   4980
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   "Rizka Connection Builder (R) email:rizka.arifianto@april.com.sg"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   1065
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1470
      Width           =   7125
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   330
      Width           =   7125
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   1785
      Left            =   120
      TabIndex        =   9
      Top             =   3060
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   3149
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Connection String:"
      Height          =   195
      Left            =   150
      TabIndex        =   11
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SQL String:"
      Height          =   195
      Left            =   150
      TabIndex        =   10
      Top             =   1260
      Width           =   810
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdButton_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
    Case 0
    Case 1: CN.DestroyConnection: Unload Me
    End Select
End Sub

Private Sub Command1_Click()
    Dim rs As ADODB.Recordset
    Dim rsSQL As String
    Dim rsCount As Long
    
    If Text2.Text = "" Then
        Call MsgBox("SQL String is null. Query aborted !", vbOKOnly, "Error")
        Exit Sub
    End If
    TmQueryStart = Timer
    Set rs = New ADODB.Recordset
    rsSQL = Text2.Text
    rs.Open rsSQL, CN.ActiveConnection, adOpenStatic, adLockReadOnly
    rs.MoveLast
    rsCount = rs.RecordCount
    Set MSHFlexGrid1.DataSource = rs
    Set rs = Nothing
    TmQueryEnd = Timer
    SttBar.SimpleText = "Query Succesfully executed " & MSHFlexGrid1.Rows - 1 & " row(s), in " & Format(TmQueryEnd - TmQueryStart, "##0.######0") & " second"
End Sub

Private Sub Command2_Click()
    'CN.ConnectionString = "Provider=MSDAORA.1;Password=etsowner;User ID=etsuser;Data Source=etstemp;Locale Identifier=1033"
    'CN.ConnectionString = "Provider=MSDASQL.1;Password=etsowner;Persist Security Info=True;User ID=etsuser;Data Source=etstest;Connect Timeout=15;Extended Properties='DSN=etstest;UID=etsuser;PWD=etsowner;SERVER=etstemp;';Locale Identifier=1033"
    Text1.Text = CN.ConnectionWizard
End Sub

Private Sub Command3_Click()
    If CN.ActiveConnection Is Nothing Then
        Call MsgBox("There is no connection established. Query aborted !", vbOKOnly, "Error")
        Exit Sub
    End If
    Debug.Print "Provider=" & CN.ActiveConnection.Provider
    Debug.Print "ConnectionString=" & CN.ActiveConnection.ConnectionString
    Debug.Print "CursorLocation=" & CN.ActiveConnection.CursorLocation
    'Debug.Print "DefaultDatabase=" & CN.ActiveConnection.DefaultDatabase
    For i = 0 To CN.ActiveConnection.Properties.Count - 1
        Debug.Print CN.ActiveConnection.Properties(i).Name & "(" & i & ")=" & CN.ActiveConnection.Properties(i).Value
    Next i
End Sub

Private Sub Command4_Click()
    Dim dd As Integer
    
    If Text1.Text = "" Then
        Call MsgBox("Connection String is null. Aborted !", vbOKOnly, "Error")
        Exit Sub
    End If
    
    TmConnectStart = Timer
    CN.ConnectionString = Text1.Text
    If CN.CreateConnection <> 0 Then
        Text1.Text = CN.ConnectionWizard
        CN.ConnectionString = Text1.Text
        TmConnectEnd = Timer
        If CN.CreateConnection <> 0 Then
            MsgBox ("Failed to establish connection.")
            End
        Else
            TmConnectEnd = Timer
            SttBar.SimpleText = "Query Succesfully executed in " & Format(TmQueryEnd - TmQueryStart, "##0.######0") & " second"
            dd = MsgBox("Connection established." & vbNewLine & "Do you want to save Connection String to registry ?", vbYesNoCancel, "Confirmation")
            Select Case dd
            Case vbYes
            Case vbNo
            Case vbCancel
            End Select
        End If
    Else
        TmConnectEnd = Timer
        SttBar.SimpleText = "Query Succesfully executed in " & Format(TmQueryEnd - TmQueryStart, "##0.######0") & " second"
        dd = MsgBox("Connection established." & vbNewLine & "Do you want to save Connection String to registry ?", vbYesNoCancel, "Confirmation")
        Select Case dd
        Case vbYes
        Case vbNo
        Case vbCancel
        End Select
    End If
End Sub
