VERSION 5.00
Begin VB.Form FrmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Burger World Login"
   ClientHeight    =   5565
   ClientLeft      =   1155
   ClientTop       =   1965
   ClientWidth     =   6900
   Icon            =   "BWLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5565
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.ComboBox TxtUser 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      Left            =   4080
      TabIndex        =   52
      Text            =   "Customer"
      Top             =   360
      Width           =   2565
   End
   Begin VB.CommandButton cmdKeypad 
      Caption         =   "Show &Keypad"
      Height          =   735
      Left            =   2904
      TabIndex        =   51
      Top             =   1560
      Width           =   1296
   End
   Begin VB.CheckBox chkDidYouKnow 
      DownPicture     =   "BWLogin.frx":030A
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   180
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   1848
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   60
      TabIndex        =   46
      Top             =   2535
      Width           =   648
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   690
      TabIndex        =   45
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1290
      TabIndex        =   44
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1890
      TabIndex        =   43
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2490
      TabIndex        =   42
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3090
      TabIndex        =   41
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3690
      TabIndex        =   40
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4290
      TabIndex        =   39
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4905
      TabIndex        =   38
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   5490
      TabIndex        =   37
      Top             =   2535
      Width           =   612
   End
   Begin VB.CommandButton cmdBackSpace 
      Caption         =   " <--"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   6090
      TabIndex        =   36
      Top             =   2535
      Width           =   732
   End
   Begin VB.CommandButton Blank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   1
      Left            =   60
      TabIndex        =   35
      Top             =   3135
      Width           =   288
   End
   Begin VB.CommandButton cmdQ 
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   330
      TabIndex        =   34
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdW 
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   930
      TabIndex        =   33
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdE 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1530
      TabIndex        =   32
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdR 
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2130
      TabIndex        =   31
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdT 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2730
      TabIndex        =   30
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdY 
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3330
      TabIndex        =   29
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdU 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3930
      TabIndex        =   28
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdI 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4530
      TabIndex        =   27
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdO 
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   5130
      TabIndex        =   26
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton cmdP 
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   5730
      TabIndex        =   25
      Top             =   3135
      Width           =   612
   End
   Begin VB.CommandButton Blank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   0
      Left            =   6330
      TabIndex        =   24
      Top             =   3135
      Width           =   492
   End
   Begin VB.CommandButton Blank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   4
      Left            =   60
      TabIndex        =   23
      Top             =   3735
      Width           =   528
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   570
      TabIndex        =   22
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1170
      TabIndex        =   21
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdD 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1770
      TabIndex        =   20
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdF 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2370
      TabIndex        =   19
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdG 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2970
      TabIndex        =   18
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdH 
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3570
      TabIndex        =   17
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdJ 
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4170
      TabIndex        =   16
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdK 
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4770
      TabIndex        =   15
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton cmdL 
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   5370
      TabIndex        =   14
      Top             =   3735
      Width           =   612
   End
   Begin VB.CommandButton Blank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   4335
      Width           =   852
   End
   Begin VB.CommandButton cmdZ 
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   900
      TabIndex        =   12
      Top             =   4335
      Width           =   612
   End
   Begin VB.CommandButton cmdX 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   1500
      TabIndex        =   11
      Top             =   4335
      Width           =   612
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2100
      TabIndex        =   10
      Top             =   4335
      Width           =   612
   End
   Begin VB.CommandButton cmdV 
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   2700
      TabIndex        =   9
      Top             =   4335
      Width           =   612
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3300
      TabIndex        =   8
      Top             =   4335
      Width           =   612
   End
   Begin VB.CommandButton cmdN 
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3900
      TabIndex        =   7
      Top             =   4335
      Width           =   612
   End
   Begin VB.CommandButton cmdM 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   4500
      TabIndex        =   6
      Top             =   4335
      Width           =   612
   End
   Begin VB.CommandButton Blank 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Index           =   2
      Left            =   5115
      TabIndex        =   5
      Top             =   4335
      Width           =   870
   End
   Begin VB.CommandButton cmdEnter 
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   5970
      TabIndex        =   4
      Top             =   3735
      Width           =   852
   End
   Begin VB.CommandButton cmdSpace 
      Caption         =   "Space"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   60
      TabIndex        =   3
      Top             =   4935
      Width           =   6768
   End
   Begin VB.TextBox TxtPW 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   444
      IMEMode         =   3  'DISABLE
      Left            =   4080
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton cmdShutdown 
      Caption         =   "&Shutdown"
      Height          =   735
      Left            =   4320
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   5520
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Show Did You Know..."
      Height          =   255
      Left            =   720
      TabIndex        =   50
      Top             =   1965
      Width           =   2295
   End
   Begin VB.Label LblPW 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   330
      Left            =   2940
      TabIndex        =   48
      Top             =   1080
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label LblUser 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      Height          =   330
      Left            =   2925
      TabIndex        =   47
      Top             =   480
      Width           =   1155
   End
   Begin VB.Image Logo 
      Height          =   1620
      Left            =   360
      Top             =   225
      Width           =   2175
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Last_Had_Focus As String    'Display_Char()
Private Sub chkDidYouKnow_Click()
    If chkDidYouKnow.Value = 1 Then
        RanDidYouKnow = False
    Else
        RanDidYouKnow = True
    End If
End Sub
Private Sub Form_Activate()
    Logo.Picture = LoadPicture(Logo_Image)
    If RanDidYouKnow = False Then chkDidYouKnow.Value = 1
    If LastUser <> "" Then txtUser.Text = LastUser
    Call Add_Users
    Call Secure_Form
    If UCase(txtUser.Text) <> "CUSTOMER" Then txtPW.SetFocus
End Sub
Private Sub Add_Users()
    txtUser.AddItem ("Customer")
    DataLink.rsEmployee.Open
        Do While Not DataLink.rsEmployee.EOF And LoggedIn = False
            txtUser.AddItem (DataLink.rsEmployee.Fields("Username"))
            DataLink.rsEmployee.MoveNext
        Loop
    DataLink.rsEmployee.Close
End Sub
Private Sub Secure_Form()
    If UCase(txtUser.Text) = "CUSTOMER" Then
        cmdKeypad.Enabled = False
        cmdShutdown.Enabled = False
        LblPW.Visible = False
        txtPW.Visible = False
        'lblcustomer.Visible = True
    Else
        cmdKeypad.Enabled = True
        cmdShutdown.Enabled = True
        LblPW.Visible = True
        txtPW.Visible = True
        'lblcustomer.Visible = False
    End If
End Sub
Private Sub Form_Load()
    Height = 2448       'Keypad
End Sub
Private Sub User_Login()
    If UCase(txtUser.Text) = "CUSTOMER" Then
        LastUser = "Customer"
        Unload FrmLogin
        'This stuff runs after Login
        If RanDidYouKnow <> True Then
            RanDidYouKnow = True
            Call GUI.Load_Form(FrmDidYouKnow, FrmDesktop)
        End If
        'MsgBox "order entry loads modally here"
        Call GUI.Load_Form(frmlunchmenu, FrmDesktop)
        Call GUI.Load_Form(Me, FrmDesktop)
    Else
        DataLink.rsEmployee.Open
        'DataLink.rsEmployee.Requery
        Do While login$ <> txtUser.Text And Not DataLink.rsEmployee.EOF
            UID = DataLink.rsEmployee.Fields("UID")
            login$ = DataLink.rsEmployee.Fields("Username")
            pass$ = DataLink.rsEmployee.Fields("Password")
            passkey = DataLink.rsEmployee.Fields("Key")
            Admin = DataLink.rsEmployee.Fields("AdminFlag")
            Disabled = DataLink.rsEmployee.Fields("DisabledFlag")
            DataLink.rsEmployee.MoveNext
        Loop
        DataLink.rsEmployee.Close
        If Disabled = True Then
            'MsgBox "This Account has been disabled"
            Stack.Push "This Account has been disabled..."
            GUI.Load_Form frmMsgBox, FrmDesktop
            txtPW.Text = ""
            txtPW.SetFocus
        Else
            If txtUser.Text = login$ Then
                If txtPW.Text = Security.Decrypt(pass$, passkey) Then
                    If Admin = True Then
                        IsManager = True
                    Else
                        IsManager = False
                    End If
                    LastUser = login$
                    Unload FrmLogin
                    'This stuff runs after Login
                    Call GUI.Visible_Icons(True)
                    If RanDidYouKnow <> True Then
                        RanDidYouKnow = True
                        Call GUI.Load_Form(FrmDidYouKnow, FrmDesktop)
                    End If
                    Call GUI.Activate_Icons(True)
                Else
                    txtPW.Text = ""
                    txtPW.SetFocus
                End If
            End If
        End If
    End If
End Sub
Private Sub Display_Char(char)
    If Last_Had_Focus = "pw" Then
        txtPW.Text = txtPW.Text + char
    Else
        txtUser.Text = txtUser.Text + char
    End If
End Sub
Private Sub cmdOK_Click()
    If UCase(txtUser.Text) <> "CUSTOMER" And txtPW.Text = "" Then
        'MsgBox "Please Enter a Password"
        Stack.Push "Please enter a password..."
        GUI.Load_Form frmMsgBox, FrmDesktop
        
        txtPW.SetFocus
    Else
        Call User_Login
    End If
End Sub
Private Sub cmdKeypad_Click()
    Call Keypad
End Sub
Private Sub Keypad()
    Call GUI.Remove_Window_Border(Me)
    Call GUI.Remove_Fake_Transparency(Me)
    diff = (5600 - 2448)
    If Height = 2448 Then
        Height = 5600
        cmdKeypad.Caption = "Hide &Keypad"
        Top = Top - 0.5 * diff
    Else
        Height = 2448
        cmdKeypad.Caption = "Show &Keypad"
        Top = Top + 0.5 * diff
    End If
    Call GUI.Window_Border(Me)
    Call GUI.Fake_Transparency(Me)
End Sub
Private Sub cmdShutdown_Click()
    If txtPW.Text <> "" Then
        DataLink.rsEmployee.Open
        'DataLink.rsEmployee.Requery
        Do While login$ <> txtUser.Text And Not DataLink.rsEmployee.EOF
            UID = DataLink.rsEmployee.Fields("UID")
            login$ = DataLink.rsEmployee.Fields("Username")
            pass$ = DataLink.rsEmployee.Fields("Password")
            passkey = DataLink.rsEmployee.Fields("Key")
            Admin = DataLink.rsEmployee.Fields("AdminFlag")
            Disabled = DataLink.rsEmployee.Fields("DisabledFlag")
            DataLink.rsEmployee.MoveNext
        Loop
        DataLink.rsEmployee.Close
        If Admin = True Then
            If Disabled = True Then
                'MsgBox "This Account has been disabled"
                Stack.Push "This Account has been disabled"
                GUI.Load_Form frmMsgBox, FrmDesktop
                txtPW.Text = ""
                txtPW.SetFocus
            Else
                If txtUser.Text = login$ Then
                    If txtPW.Text = Security.Decrypt(pass$, passkey) Then
                        End
                    Else
                        'MsgBox "Invalid Password"
                        Stack.Push "Invalid password..."
                        GUI.Load_Form frmMsgBox, FrmDesktop
                        txtPW.Text = ""
                        txtPW.SetFocus
                    End If
                End If
            End If
        Else
            'MsgBox "Only Managers can Shutdown Burger World"
            Stack.Push "Only a manager can shutdown Burger World!"
            GUI.Load_Form frmMsgBox, FrmDesktop
        End If
    Else
        'MsgBox "A Password is Required to Shutdown Burger World."
        Stack.Push "A Password is Required to Shutdown Burger World."
        GUI.Load_Form frmMsgBox, FrmDesktop
        txtPW.SetFocus
    End If
End Sub
Private Sub TxtPW_LostFocus()
    Last_Had_Focus = "pw"
End Sub
Private Sub TxtUser_Change()
    Call Secure_Form
End Sub
Private Sub TxtUser_Click()
    Call Secure_Form
    txtPW.Text = ""
    If txtUser.Text = "Customer" Then
        cmdOK.SetFocus
    Else
        txtPW.SetFocus
    End If
End Sub
Private Sub TxtUser_LostFocus()
    Last_Had_Focus = "user"
    If cmdKeypad.Caption = "Hide &Keypad" And txtUser.Text = "Customer" Then
        Call Keypad
    End If
End Sub
Private Sub cmd1_Click()
    Call Display_Char("1")
End Sub
Private Sub cmd2_Click()
    Call Display_Char("2")
End Sub
Private Sub cmd3_Click()
    Call Display_Char("3")
End Sub
Private Sub cmd4_Click()
    Call Display_Char("4")
End Sub
Private Sub cmd5_Click()
    Call Display_Char("5")
End Sub
Private Sub cmd6_Click()
    Call Display_Char("6")
End Sub
Private Sub cmd7_Click()
    Call Display_Char("7")
End Sub
Private Sub cmd8_Click()
    Call Display_Char("8")
End Sub
Private Sub cmd9_Click()
    Call Display_Char("9")
End Sub
Private Sub cmd0_Click()
    Call Display_Char("0")
End Sub
Private Sub cmdA_Click()
    Call Display_Char("a")
End Sub
Private Sub cmdB_Click()
    Call Display_Char("b")
End Sub
Private Sub cmdC_Click()
    Call Display_Char("c")
End Sub
Private Sub cmdD_Click()
    Call Display_Char("d")
End Sub
Private Sub cmdE_Click()
    Call Display_Char("e")
End Sub
Private Sub cmdF_Click()
    Call Display_Char("f")
End Sub
Private Sub cmdG_Click()
    Call Display_Char("g")
End Sub
Private Sub cmdH_Click()
    Call Display_Char("h")
End Sub
Private Sub cmdI_Click()
    Call Display_Char("i")
End Sub
Private Sub cmdJ_Click()
    Call Display_Char("j")
End Sub
Private Sub cmdK_Click()
    Call Display_Char("k")
End Sub
Private Sub cmdL_Click()
    Call Display_Char("l")
End Sub
Private Sub cmdM_Click()
    Call Display_Char("m")
End Sub
Private Sub cmdN_Click()
    Call Display_Char("n")
End Sub
Private Sub cmdO_Click()
    Call Display_Char("o")
End Sub
Private Sub cmdP_Click()
    Call Display_Char("p")
End Sub
Private Sub cmdQ_Click()
    Call Display_Char("q")
End Sub
Private Sub cmdR_Click()
    Call Display_Char("r")
End Sub
Private Sub cmdS_Click()
    Call Display_Char("s")
End Sub
Private Sub cmdT_Click()
    Call Display_Char("t")
End Sub
Private Sub cmdU_Click()
    Call Display_Char("u")
End Sub
Private Sub cmdV_Click()
    Call Display_Char("v")
End Sub
Private Sub cmdW_Click()
    Call Display_Char("w")
End Sub
Private Sub cmdX_Click()
    Call Display_Char("x")
End Sub
Private Sub cmdY_Click()
    Call Display_Char("y")
End Sub
Private Sub cmdZ_Click()
    Call Display_Char("z")
End Sub
Private Sub cmdSpace_Click()
    Call Display_Char(" ")
End Sub
Private Sub cmdEnter_Click()
    If Last_Had_Focus = "pw" Then
        Call User_Login
    Else
        txtPW.SetFocus
    End If
End Sub
Private Sub cmdBackSpace_Click()
    If Last_Had_Focus = "pw" Then
        If Len(txtPW.Text) <> 0 Then txtPW.Text = Left(txtPW.Text, Len(txtPW.Text) - 1)
    Else
        If Len(txtUser.Text) <> 0 Then txtUser.Text = Left(txtUser.Text, Len(txtUser.Text) - 1)
    End If
End Sub
