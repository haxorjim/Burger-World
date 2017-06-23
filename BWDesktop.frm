VERSION 5.00
Begin VB.Form FrmDesktop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Burger World Desktop"
   ClientHeight    =   5670
   ClientLeft      =   495
   ClientTop       =   945
   ClientWidth     =   7575
   Icon            =   "BWDesktop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWDesktop.frx":030A
      Enabled         =   0   'False
      Height          =   2115
      Index           =   2
      Left            =   3984
      MouseIcon       =   "BWDesktop.frx":58D0
      MousePointer    =   99  'Custom
      Picture         =   "BWDesktop.frx":5A22
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   672
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWDesktop.frx":14CBC
      Enabled         =   0   'False
      Height          =   2115
      Index           =   3
      Left            =   4740
      MouseIcon       =   "BWDesktop.frx":1A282
      MousePointer    =   99  'Custom
      Picture         =   "BWDesktop.frx":1A3D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWDesktop.frx":2966E
      Enabled         =   0   'False
      Height          =   2115
      Index           =   4
      Left            =   5664
      MaskColor       =   &H8000000F&
      MouseIcon       =   "BWDesktop.frx":2B6D0
      MousePointer    =   99  'Custom
      Picture         =   "BWDesktop.frx":2B822
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2532
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWDesktop.frx":2F62D
      Enabled         =   0   'False
      Height          =   2115
      Index           =   1
      Left            =   1968
      MouseIcon       =   "BWDesktop.frx":322A3
      MousePointer    =   99  'Custom
      Picture         =   "BWDesktop.frx":323F5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   756
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.CommandButton Icons 
      Appearance      =   0  'Flat
      DisabledPicture =   "BWDesktop.frx":4168F
      Enabled         =   0   'False
      Height          =   2115
      Index           =   0
      Left            =   204
      MouseIcon       =   "BWDesktop.frx":43AA5
      MousePointer    =   99  'Custom
      Picture         =   "BWDesktop.frx":43BF7
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   2190
   End
   Begin VB.Image Background 
      Height          =   5736
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7692
   End
   Begin VB.Image Default_Icon 
      Height          =   2115
      Left            =   1200
      Picture         =   "BWDesktop.frx":47CF2
      Top             =   3360
      Visible         =   0   'False
      Width           =   2190
   End
End
Attribute VB_Name = "FrmDesktop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call GUI.Full_Screen(Me)
    Call GUI.Fit_Background(Me)
    Call GUI.Window_Border(Me)
    Call GUI.Place_Icons(Me, 5, 125, (3 / 5))
    Call GUI.Visible_Icons(False)
    Me.Visible = True
    ''Loads jpg desktop image from resource file
    'LoadDataIntoFile 101, "tempjpegfile"
    'Background.Picture = LoadPicture("tempjpegfile")
    'Kill "tempjpegfile"
    '
    '*NEW* Load Desktop Image
    '
    Background.Picture = LoadPicture(Desktop_Image)
    Call GUI.Load_Form(FrmLogin, FrmDesktop)
End Sub
Private Sub Icons_Click(Index As Integer)
    If Index = 0 Then
        Call GUI.Activate_Icons(False)
        Call GUI.Load_Form(frmlunchmenu, FrmDesktop)
        Call GUI.Activate_Icons(True)
    ElseIf Index = 1 Then
        Call GUI.Activate_Icons(False)
        Call GUI.Load_Form(FrmKitchenMain, FrmDesktop)
        Call GUI.Activate_Icons(True)
    ElseIf Index = 2 Then
        Call GUI.Activate_Icons(False)
        Call GUI.Load_Form(FrmSchedule, FrmDesktop)
        Call GUI.Activate_Icons(True)
    ElseIf Index = 3 Then
        OnManagerPanel = True
        Call GUI.Activate_Icons(False)
        Call GUI.Load_Form(FrmManagerControlPanel, FrmDesktop)
        Call GUI.Activate_Icons(True)
        OnManagerPanel = False
    ElseIf Index = 4 Then
        Call GUI.Activate_Icons(False)
        Call GUI.Load_Form(FrmLogoff, FrmDesktop)
        Call GUI.Activate_Icons(True)
    End If
End Sub
