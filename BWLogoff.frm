VERSION 5.00
Begin VB.Form FrmLogoff 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   1695
   ClientTop       =   2400
   ClientWidth     =   3645
   Icon            =   "BWLogoff.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton No 
      Caption         =   "&No"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2025
      TabIndex        =   2
      Top             =   2112
      Width           =   1095
   End
   Begin VB.CommandButton Yes 
      Caption         =   "&Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   585
      TabIndex        =   1
      Top             =   2112
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Are you sure you want to Log Off?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Image Logo 
      Height          =   1620
      Left            =   720
      Top             =   105
      Width           =   2175
   End
End
Attribute VB_Name = "FrmLogoff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    logo.Left = (Width - logo.Width) / 2
    Label1.Left = (Width - Label1.Width) / 2
    logo.Picture = LoadPicture(Logo_Image)
End Sub
Private Sub No_Click()
    Unload FrmLogoff
End Sub
Private Sub Yes_Click()
    Unload FrmLogoff
    Call GUI.Activate_Icons(False)
    Call GUI.Visible_Icons(False)
    Call GUI.Load_Form(FrmLogin, FrmDesktop)
End Sub
