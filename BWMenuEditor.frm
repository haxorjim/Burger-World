VERSION 5.00
Begin VB.Form frmMenuEditor 
   BorderStyle     =   0  'None
   ClientHeight    =   4092
   ClientLeft      =   1356
   ClientTop       =   3408
   ClientWidth     =   9456
   LinkTopic       =   "Form1"
   ScaleHeight     =   4092
   ScaleWidth      =   9456
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWMenuEditor.frx":0000
      Height          =   1812
      Index           =   3
      Left            =   6945
      MouseIcon       =   "BWMenuEditor.frx":F29A
      MousePointer    =   99  'Custom
      Picture         =   "BWMenuEditor.frx":F3EC
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   885
      Width           =   2052
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWMenuEditor.frx":1E686
      Height          =   1848
      Index           =   0
      Left            =   315
      MouseIcon       =   "BWMenuEditor.frx":23C4C
      MousePointer    =   99  'Custom
      Picture         =   "BWMenuEditor.frx":23D9E
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   900
      Width           =   1872
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWMenuEditor.frx":33038
      Height          =   1860
      Index           =   1
      Left            =   2535
      MouseIcon       =   "BWMenuEditor.frx":385FE
      MousePointer    =   99  'Custom
      Picture         =   "BWMenuEditor.frx":38750
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   900
      Width           =   1932
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWMenuEditor.frx":479EA
      Height          =   1812
      Index           =   2
      Left            =   4680
      MouseIcon       =   "BWMenuEditor.frx":4CFB0
      MousePointer    =   99  'Custom
      Picture         =   "BWMenuEditor.frx":4D102
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   900
      Width           =   2052
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Menu Editor"
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
      Height          =   570
      Left            =   75
      TabIndex        =   3
      Top             =   60
      Width           =   8940
   End
End
Attribute VB_Name = "frmMenuEditor"
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
    Call GUI.Place_Icons(Me, 4, 75, (1 / 3))
End Sub
Private Sub Icons_Click(Index As Integer)
    If Index = 0 Then
        Call Activate_Icons(False)
        'drinks
        Call GUI.Load_Form(frmEditMainMenu, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 1 Then
        Call Activate_Icons(False)
        'sides
        Call GUI.Load_Form(frmEditMainMenu, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 2 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(frmEditMainMenu, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 3 Then
        Unload Me
    End If
End Sub
Public Sub Activate_Icons(TrueFalse As Boolean)
    For Index = 0 To 3
        Icons(Index).Enabled = TrueFalse
    Next
    If title.BackColor = &HC0C0C0 Then
        title.BackColor = &HFFFFFF
    ElseIf title.BackColor = &HFFFFFF Then
        title.BackColor = &HC0C0C0
    End If
End Sub

