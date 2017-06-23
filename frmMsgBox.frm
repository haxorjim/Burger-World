VERSION 5.00
Begin VB.Form frmMsgBox 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2070
   ClientLeft      =   3645
   ClientTop       =   3435
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOkay 
      Caption         =   "OK"
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Something went wrong, just click ok..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   345
      TabIndex        =   0
      Top             =   210
      Width           =   4830
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOkay_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Message.Caption = Stack.Pop
End Sub
