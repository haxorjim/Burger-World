VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Burger World Splash!"
   ClientHeight    =   4785
   ClientLeft      =   2295
   ClientTop       =   2025
   ClientWidth     =   6165
   Icon            =   "BWSplash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrSplash 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   5460
      Top             =   4185
   End
   Begin VB.Image Background 
      Height          =   4800
      Left            =   -15
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6180
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright ©2001 Lakeland CIS Class.  All Rights Reserved"
      Height          =   270
      Left            =   150
      TabIndex        =   0
      Top             =   4485
      Width           =   4590
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Displays the Splash Screen
    Call GUI.Center_Form(Me)
    Background.Picture = LoadPicture(Splash_Image)
    Call GUI.Fit_Background(Me)
    Call GUI.Window_Border(Me)
    Me.Visible = True
    tmrSplash.Enabled = True
End Sub
Private Sub tmrSplash_Timer()
    'Splash Screen Closes after two seconds
    Unload FrmSplash
    FrmDesktop.Show
End Sub
