VERSION 5.00
Begin VB.Form FrmUsrManage 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "User Management"
   ClientHeight    =   3660
   ClientLeft      =   432
   ClientTop       =   3552
   ClientWidth     =   10956
   Icon            =   "FrmUsrManage2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   10956
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmUsrManage2.frx":030A
      Height          =   1848
      Index           =   0
      Left            =   180
      MouseIcon       =   "FrmUsrManage2.frx":58D0
      MousePointer    =   99  'Custom
      Picture         =   "FrmUsrManage2.frx":5A22
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1872
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmUsrManage2.frx":14CBC
      Height          =   1860
      Index           =   1
      Left            =   2160
      MouseIcon       =   "FrmUsrManage2.frx":1A282
      MousePointer    =   99  'Custom
      Picture         =   "FrmUsrManage2.frx":1A3D4
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1932
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmUsrManage2.frx":2966E
      Height          =   1812
      Index           =   2
      Left            =   4200
      MousePointer    =   99  'Custom
      Picture         =   "FrmUsrManage2.frx":2EC34
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1455
      Width           =   2052
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmUsrManage2.frx":3DECE
      Height          =   1812
      Index           =   4
      Left            =   8520
      MouseIcon       =   "FrmUsrManage2.frx":4D168
      MousePointer    =   99  'Custom
      Picture         =   "FrmUsrManage2.frx":4D2BA
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2052
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmUsrManage2.frx":5C554
      Height          =   1812
      Index           =   3
      Left            =   6360
      MouseIcon       =   "FrmUsrManage2.frx":61B1A
      MousePointer    =   99  'Custom
      Picture         =   "FrmUsrManage2.frx":61C6C
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2064
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Security Management"
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
      Left            =   285
      TabIndex        =   5
      Top             =   180
      Width           =   11250
   End
End
Attribute VB_Name = "FrmUsrManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Initialize()
    Title.Left = 100
    Title.Width = Me.Width - 200
    Title.Top = 200
    Title.BackColor = &HFFFFFF
End Sub
Private Sub Form_Load()
    Call GUI.Place_Icons(Me, 5, 20, (1 / 3))
End Sub
Private Sub Icons_Click(Index As Integer)
    If Index = 0 Then
        Call Activate_Icons(False)
        Stack.Push 1
        Call GUI.Load_Form(frmUsrDBINT, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 1 Then
        Call Activate_Icons(False)
        Stack.Push 2
        Call GUI.Load_Form(frmUsrDBINT, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 2 Then
        Call Activate_Icons(False)
        Stack.Push 3
        Call GUI.Load_Form(frmUsrDBINT, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 3 Then
        Call Activate_Icons(False)
        Stack.Push 4
        Call GUI.Load_Form(frmUsrDBINT, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 4 Then
        Unload Me
    End If
End Sub
Public Sub Activate_Icons(TrueFalse As Boolean)
    For Index = 0 To 4
        Icons(Index).Enabled = TrueFalse
    Next
    If Title.BackColor = &HC0C0C0 Then
        Title.BackColor = &HFFFFFF
    ElseIf Title.BackColor = &HFFFFFF Then
        Title.BackColor = &HC0C0C0
    End If
End Sub
