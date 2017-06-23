VERSION 5.00
Begin VB.Form FrmEditSides 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   1320
   ClientTop       =   1755
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheEnabled 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Desktop\Newest Shell\Menu.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Drinks"
      Top             =   5040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox TxtName 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   3240
      Width           =   1815
   End
   Begin VB.TextBox TxtLarge 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox TxtMedium 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   4320
      Width           =   1815
   End
   Begin VB.TextBox TxtSmall 
      DataSource      =   "Data1"
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Done"
      Height          =   372
      Left            =   2640
      TabIndex        =   9
      Top             =   4920
      Width           =   1812
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hash Brown"
      Height          =   492
      Index           =   2
      Left            =   4905
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   450
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Onion Rings"
      Height          =   492
      Index           =   1
      Left            =   2685
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   450
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fries"
      Height          =   492
      Index           =   0
      Left            =   450
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   450
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apple Pie"
      Height          =   492
      Index           =   3
      Left            =   465
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1050
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salad"
      Height          =   492
      Index           =   4
      Left            =   2700
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1065
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chili Cheese Fries"
      Height          =   492
      Index           =   5
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1065
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Soup"
      Height          =   492
      Index           =   6
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1680
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cookies"
      Height          =   492
      Index           =   7
      Left            =   2715
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1695
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muffins"
      Height          =   492
      Index           =   8
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1695
      Width           =   1692
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Side Attributes "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   360
      TabIndex        =   10
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Side Orders "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Available?"
      Height          =   255
      Left            =   4800
      TabIndex        =   21
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Large"
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   20
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Medium"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   19
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Small"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sizes and Prices"
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drink Name"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Shape Shape2 
      Height          =   2265
      Left            =   240
      Top             =   210
      Width           =   6690
   End
   Begin VB.Shape Shape1 
      Height          =   2175
      Left            =   240
      Top             =   2595
      Width           =   6645
   End
End
Attribute VB_Name = "FrmEditSides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
