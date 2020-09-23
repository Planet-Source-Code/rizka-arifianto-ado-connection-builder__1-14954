VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rizka Connection Builder"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4305
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "frmAbout.frx":0442
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      Caption         =   "I build this application to help me creating Proper Connection string to any database."
      ForeColor       =   &H00000000&
      Height          =   1410
      Left            =   1050
      TabIndex        =   1
      Top             =   990
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Database Connection Builder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version 1.0"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   630
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "This application is Public domain. You may download and distribute this. "
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   270
      TabIndex        =   2
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    frmMe.Show vbModal
    Unload Me
End Sub

Private Sub Form_Load()
    'Me.Caption = "About " & App.Title
    'lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    'lblTitle.Caption = App.Title
    lblDescription = lblDescription & vbNewLine & vbNewLine & _
        "Copyrights(c) 2001 - Rizka Arifianto" & vbNewLine & _
        "E-mail: Rizka.Arifianto@april.com.sg" & vbNewLine & _
        "            rizkaarifianto@hotmail.com"
        
    lblDisclaimer = lblDisclaimer & vbNewLine & vbNewLine & _
        "Please Vote Me !!!"
End Sub

