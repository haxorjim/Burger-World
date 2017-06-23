VERSION 5.00
Begin VB.Form frmEditMainMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Edit Order Menu"
   ClientHeight    =   7020
   ClientLeft      =   -456
   ClientTop       =   1776
   ClientWidth     =   9648
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   9648
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtTax 
      DataField       =   "Tax"
      DataSource      =   "Data1"
      Height          =   288
      Left            =   7440
      TabIndex        =   35
      Top             =   4800
      Width           =   1092
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Menu.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Food"
      Top             =   8280
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   17
      Left            =   7452
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   16
      Left            =   7440
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   15
      Left            =   7452
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1944
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   14
      Left            =   6132
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   13
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   12
      Left            =   6132
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1944
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   11
      Left            =   4812
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   10
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   9
      Left            =   4812
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1944
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   8
      Left            =   3492
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   7
      Left            =   3492
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   6
      Left            =   3492
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1944
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   4
      Left            =   2172
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   3
      Left            =   2172
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1944
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   2
      Left            =   852
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   1
      Left            =   852
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton ComEditFood 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1944
      Width           =   1095
   End
   Begin VB.CommandButton ComExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5880
      Width           =   5055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "New Tax Rate in Decimal form:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4320
      TabIndex        =   34
      Top             =   4800
      Width           =   2892
   End
   Begin VB.Label LblEnabled 
      DataField       =   "Enabled"
      DataSource      =   "Data1"
      Height          =   492
      Left            =   7560
      TabIndex        =   33
      Top             =   6120
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   5
      Left            =   7452
      TabIndex        =   32
      Top             =   1224
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   4
      Left            =   6132
      TabIndex        =   31
      Top             =   1224
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   3
      Left            =   4692
      TabIndex        =   30
      Top             =   1224
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column 6"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   7404
      TabIndex        =   29
      Top             =   864
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column 4"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   4644
      TabIndex        =   28
      Top             =   864
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column 3"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   3324
      TabIndex        =   27
      Top             =   864
      Width           =   1212
   End
   Begin VB.Label lblItem 
      Caption         =   "Label Item"
      DataField       =   "Item Name"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2520
      TabIndex        =   26
      Top             =   10680
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column 2"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   2004
      TabIndex        =   4
      Top             =   864
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   1
      Left            =   2172
      TabIndex        =   2
      Top             =   1224
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   852
      TabIndex        =   1
      Top             =   1224
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column 1"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   684
      TabIndex        =   0
      Top             =   864
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Column 5"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   6084
      TabIndex        =   5
      Top             =   864
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   2
      Left            =   3372
      TabIndex        =   3
      Top             =   1224
      Width           =   1212
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CHOOSE A MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   3120
      TabIndex        =   6
      Top             =   240
      Width           =   3372
   End
End
Attribute VB_Name = "frmEditMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Change As Boolean
Dim InpTaxRate As Variant

Private Sub Refreshmenu()
    
    Dim Index As Integer
    Index = 0
    ' Set Command button captions to the database values

    Data1.Refresh
    Data1.Recordset.MoveFirst
    For Index = 0 To 17
        ComEditFood(Index).Caption = lblItem.Caption
        If lblEnabled.Caption = "2" Then
        ComEditFood(Index).BackColor = vbRed
        Else
        ComEditFood(Index).BackColor = vbWhite
        End If
        Data1.Recordset.MoveNext
    Next
    

End Sub

Private Sub Check1_Click(Index As Integer)

End Sub


Private Sub ComEditDrinks_Click()
    Form1.Show 1
End Sub

Private Sub ComEditFood_Click(Index As Integer)
Stack.Push (ComEditFood(Index).Caption)
'frmEditItem.Show 1
Call GUI.Load_Form(frmEditItem, Me)
Call Refreshmenu
End Sub

Private Sub ComEditSides_Click()
    FrmEditSides.Show 1
End Sub

Private Sub ComExit_Click()

If IsNumeric(TxtTax.Text) = False Then
    TxtTax.Text = 0.0575
End If
Unload Me
End Sub

Private Sub Command3_Click()
frmEditDrinks.Show 1
Call Refreshmenu
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
    frmSave.Show 1
    Unload Me
End Sub
Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    If File1.Path = Drive1.Drive Then
        SelectedBreakFile = File1.Path & File1.FileName
        SelectedLunchFile = File1.Path & File1.FileName
    Else
        SelectedBreakFile = File1.Path & "\" & File1.FileName
        SelectedLunchFile = File1.Path & "\" & File1.FileName
    End If
End Sub


Private Sub TxtEditFood_Change(Index As Integer)
    Changed% = 1
End Sub


Private Sub Form_Load()
Call GUI.Full_Screen(Me)
Dim Index As Integer
Call Refreshmenu
End Sub

Private Sub Option1_Click()
Call Refreshmenu
End Sub

Private Sub Option2_Click()
Call Refreshmenu
End Sub

Private Sub txtEditPrice_Change(Index As Integer)

End Sub


Private Sub TxtTax_Change()

TaxRate = Val(TxtTax.Text)
Open "Taxrate.txt" For Output As #2
Write #2, TaxRate
Close #2

End Sub
