VERSION 5.00
Begin VB.Form Payroll_menu 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Payroll Main Menu"
   ClientHeight    =   3780
   ClientLeft      =   3510
   ClientTop       =   3285
   ClientWidth     =   7440
   Icon            =   "Payroll_menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "Payroll_menu.frx":030A
      Height          =   1812
      Index           =   2
      Left            =   4680
      MouseIcon       =   "Payroll_menu.frx":F5A4
      MousePointer    =   99  'Custom
      Picture         =   "Payroll_menu.frx":F6F6
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2052
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "Payroll_menu.frx":1E990
      Height          =   1860
      Index           =   1
      Left            =   2580
      MouseIcon       =   "Payroll_menu.frx":23F56
      MousePointer    =   99  'Custom
      Picture         =   "Payroll_menu.frx":240A8
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1932
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "Payroll_menu.frx":33342
      Height          =   1848
      Index           =   0
      Left            =   360
      MouseIcon       =   "Payroll_menu.frx":38908
      MousePointer    =   99  'Custom
      Picture         =   "Payroll_menu.frx":38A5A
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1872
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Payroll System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   576
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   6456
   End
End
Attribute VB_Name = "Payroll_menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
    title.Left = 100
    title.Width = Me.Width - 200
    title.Top = 200
    title.BackColor = &HFFFFFF
End Sub
Private Sub Form_Load()
    Call GUI.Place_Icons(Me, 3, 75, (1 / 3))
End Sub
Private Sub Icons_Click(Index As Integer)
    If Index = 0 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(Payroll_Calculator, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 1 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(Payroll_Records, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 2 Then
        Unload Me
    End If
End Sub
Public Sub Activate_Icons(TrueFalse As Boolean)
    For Index = 0 To 2
        Icons(Index).Enabled = TrueFalse
    Next
    If title.BackColor = &HC0C0C0 Then
        title.BackColor = &HFFFFFF
    ElseIf title.BackColor = &HFFFFFF Then
        title.BackColor = &HC0C0C0
    End If
End Sub
