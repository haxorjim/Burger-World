VERSION 5.00
Begin VB.Form frmAreYouSure 
   BorderStyle     =   0  'None
   ClientHeight    =   2052
   ClientLeft      =   1848
   ClientTop       =   1896
   ClientWidth     =   5268
   LinkTopic       =   "Form1"
   ScaleHeight     =   2052
   ScaleWidth      =   5268
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdno 
      Caption         =   "No"
      Height          =   855
      Left            =   2640
      TabIndex        =   2
      Top             =   930
      Width           =   1305
   End
   Begin VB.CommandButton cmdyes 
      Caption         =   "Yes"
      Height          =   855
      Left            =   1215
      TabIndex        =   0
      Top             =   945
      Width           =   1305
   End
   Begin VB.Label Message 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Something went wrong, better click no..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   120
      TabIndex        =   1
      Top             =   195
      Width           =   4830
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAreYouSure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdno_Click()
    Stack.Push "NO"
    Unload Me
End Sub
Private Sub cmdyes_Click()
    Stack.Push "YES"
    Unload Me
End Sub
Private Sub Form_Load()
    Message.Caption = Stack.Pop
End Sub
