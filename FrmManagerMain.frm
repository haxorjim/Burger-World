VERSION 5.00
Begin VB.Form FrmManagerControlPanel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6336
   ClientLeft      =   1668
   ClientTop       =   1392
   ClientWidth     =   8472
   LinkTopic       =   "Form1"
   ScaleHeight     =   6336
   ScaleWidth      =   8472
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmManagerMain.frx":0000
      Height          =   1812
      Index           =   3
      Left            =   885
      MouseIcon       =   "FrmManagerMain.frx":55C6
      MousePointer    =   99  'Custom
      Picture         =   "FrmManagerMain.frx":5718
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3285
      Width           =   2052
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmManagerMain.frx":149B2
      Height          =   1812
      Index           =   4
      Left            =   3285
      MouseIcon       =   "FrmManagerMain.frx":17C34
      MousePointer    =   99  'Custom
      Picture         =   "FrmManagerMain.frx":17D86
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3270
      Width           =   2064
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmManagerMain.frx":27020
      Height          =   1812
      Index           =   5
      Left            =   5670
      MouseIcon       =   "FrmManagerMain.frx":362BA
      MousePointer    =   99  'Custom
      Picture         =   "FrmManagerMain.frx":3640C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   2052
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmManagerMain.frx":456A6
      Height          =   1812
      Index           =   2
      Left            =   5610
      MouseIcon       =   "FrmManagerMain.frx":4AC6C
      MousePointer    =   99  'Custom
      Picture         =   "FrmManagerMain.frx":4ADBE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2052
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmManagerMain.frx":59E24
      Height          =   1860
      Index           =   1
      Left            =   3360
      MouseIcon       =   "FrmManagerMain.frx":5BC7A
      MousePointer    =   99  'Custom
      Picture         =   "FrmManagerMain.frx":5BDCC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1170
      Width           =   1932
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "FrmManagerMain.frx":5FC7E
      Height          =   1848
      Index           =   0
      Left            =   990
      MouseIcon       =   "FrmManagerMain.frx":62178
      MousePointer    =   99  'Custom
      Picture         =   "FrmManagerMain.frx":622CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1176
      Width           =   1872
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Manager Control Panel"
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
      TabIndex        =   3
      Top             =   225
      Width           =   7905
   End
End
Attribute VB_Name = "FrmManagerControlPanel"
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
    Call GUI.Place_Icons(Me, 6, 125, (1 / 5))
End Sub
Private Sub Icons_Click(Index As Integer)
    If Index = 0 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(Payroll_menu, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 1 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(FrmUsrManage, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 2 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(frmMenuEditor, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 3 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(SalesJournal, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 4 Then
        Call Activate_Icons(False)
        Call GUI.Load_Form(FrmSchedule, FrmDesktop)
        Call Activate_Icons(True)
    ElseIf Index = 5 Then
        Unload Me
    End If
End Sub
Public Sub Activate_Icons(TrueFalse As Boolean)
    For Index = 0 To 5
        Icons(Index).Enabled = TrueFalse
    Next
    If title.BackColor = &HC0C0C0 Then
        title.BackColor = &HFFFFFF
    ElseIf title.BackColor = &HFFFFFF Then
        title.BackColor = &HC0C0C0
    End If
End Sub
