VERSION 5.00
Begin VB.Form frmSides 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   1125
   ClientTop       =   2190
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LstTax 
      Height          =   1035
      ItemData        =   "frmSides.frx":0000
      Left            =   240
      List            =   "frmSides.frx":0002
      TabIndex        =   27
      Top             =   6120
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ListBox LstSizes 
      Height          =   1035
      ItemData        =   "frmSides.frx":0004
      Left            =   3360
      List            =   "frmSides.frx":0006
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox TxtSizes 
      DataField       =   "Special"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5280
      TabIndex        =   25
      Top             =   6600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access 2000;"
      DatabaseName    =   "menu.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sides"
      Top             =   6480
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox TxtName 
      DataField       =   "Side Name"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3960
      TabIndex        =   24
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtSmall 
      DataField       =   "Small"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3960
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtMedium 
      DataField       =   "Medium"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3960
      TabIndex        =   22
      Top             =   6120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtLarge 
      DataField       =   "Large"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   3960
      TabIndex        =   21
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtTax 
      DataField       =   "Tax"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   5640
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtEnabled 
      DataField       =   "Enabled"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.ListBox LstSmall 
      Height          =   1035
      ItemData        =   "frmSides.frx":0008
      Left            =   600
      List            =   "frmSides.frx":001B
      TabIndex        =   18
      Top             =   5760
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ListBox LstMedium 
      Height          =   1035
      ItemData        =   "frmSides.frx":0038
      Left            =   1080
      List            =   "frmSides.frx":004B
      TabIndex        =   17
      Top             =   6360
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ListBox LstLarge 
      Height          =   1035
      ItemData        =   "frmSides.frx":0068
      Left            =   2280
      List            =   "frmSides.frx":007B
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muffins"
      Height          =   492
      Index           =   8
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cookies"
      Height          =   492
      Index           =   7
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Soup"
      Height          =   492
      Index           =   6
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Chili Cheese Fries"
      Height          =   492
      Index           =   5
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cookies"
      Height          =   492
      Index           =   4
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1200
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apple Pie"
      Height          =   492
      Index           =   3
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CheckBox optsize 
      Caption         =   "Small"
      Height          =   420
      Index           =   0
      Left            =   564
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3096
      Value           =   1  'Checked
      Width           =   1584
   End
   Begin VB.CheckBox optsize 
      Caption         =   "Medium"
      Height          =   420
      Index           =   1
      Left            =   2688
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3108
      Width           =   1584
   End
   Begin VB.CheckBox optsize 
      Caption         =   "Large"
      Height          =   420
      Index           =   2
      Left            =   4908
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3108
      Width           =   1584
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "French Fries"
      CausesValidation=   0   'False
      Height          =   492
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   600
      Value           =   -1  'True
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Onion Rings"
      Height          =   492
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   1692
   End
   Begin VB.OptionButton Optitem 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muffins"
      Height          =   492
      Index           =   2
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   600
      Width           =   1692
   End
   Begin VB.CommandButton ComAdd 
      Caption         =   "Add to Order"
      Height          =   492
      Left            =   345
      TabIndex        =   1
      Top             =   4320
      Width           =   1932
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   492
      Left            =   4920
      TabIndex        =   0
      Top             =   4320
      Width           =   1812
   End
   Begin VB.Image logo 
      Height          =   1620
      Left            =   2580
      Top             =   3840
      Width           =   2175
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sizes "
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   324
      TabIndex        =   9
      Top             =   2700
      Width           =   492
   End
   Begin VB.Shape Shape1 
      Height          =   972
      Left            =   204
      Top             =   2820
      Width           =   6648
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Side Orders "
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   300
      TabIndex        =   5
      Top             =   228
      Width           =   960
   End
   Begin VB.Shape Shape2 
      Height          =   2175
      Left            =   120
      Top             =   360
      Width           =   6735
   End
End
Attribute VB_Name = "frmSides"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComAdd_Click()
If frmSides.optsize(0).Value = 1 Then
Size = "S"
Listnum = 1
End If
If frmSides.optsize(1).Value = 1 Then
Size = "M"
Listnum = 2
End If
If frmSides.optsize(2).Value = 1 Then
Size = "L"
Listnum = 3
End If
For SideIndex = 0 To 8
    If frmSides.Optitem(SideIndex).Value = True Then Exit For
Next
If LstSizes.List(SideIndex) = "False" Then Size = ""
If Listnum = 1 Then Cost = LstSmall.List(SideIndex)
If Listnum = 2 Then Cost = LstMedium.List(SideIndex)
If Listnum = 3 Then Cost = LstLarge.List(SideIndex)
Item = Size & Optitem(SideIndex).Caption
For X = 0 To frmlunchmenu.List1.ListCount
    If Item = frmlunchmenu.List1.List(X) Then
        frmlunchmenu.LstQuantity.List(X) = frmlunchmenu.LstQuantity.List(X) + 1
        Exit Sub
    End If
Next
If LstTax.List(Index) = "True" Then
Taxable = "True"
Else
Taxable = "False"
End If
frmlunchmenu.List1.AddItem (Item)
frmlunchmenu.LstQuantity.AddItem (1)
frmlunchmenu.LstPrice.AddItem (FormatCurrency(Cost))
frmlunchmenu.LstUnformPrice.AddItem (Cost)
frmlunchmenu.LstIngredient.AddItem ("Nothing")
frmlunchmenu.LstRegName.AddItem ("Nothing")
frmlunchmenu.LstTax.AddItem (Taxable)
End Sub

Private Sub Command3_Click()
Unload Me
End Sub


Private Sub Form_Load()
    logo.Picture = LoadPicture(Logo_Image)
    Data2.Refresh
    For Index = 0 To 8
        Optitem(Index).Caption = txtName.Text
        If TxtEnabled.Text = "True" Then Optitem(Index).Visible = True
        If TxtEnabled.Text = "False" Then Optitem(Index).Visible = False
        LstSmall.List(Index) = TxtSmall.Text
        LstMedium.List(Index) = TxtMedium.Text
        LstLarge.List(Index) = TxtLarge.Text
        LstSizes.List(Index) = TxtSizes.Text
        LstTax.List(Index) = TxtTax.Text
        Data2.Recordset.MoveNext
        If LstSizes.List(0) = "False" Then
        For X = 0 To 2
        optsize(X).Visible = False
        Shape1.Visible = False
        Label1.Visible = False
        Next
        Else
        For X = 0 To 2
        optsize(X).Visible = True
        Shape1.Visible = True
        Label1.Visible = True
    Next
    End If
    Next
End Sub

Private Sub Optitem_Click(Index As Integer)
    If LstSizes.List(Index) = "False" Then
    For X = 0 To 2
        optsize(X).Visible = False
        Shape1.Visible = False
        Label1.Visible = False
    Next
    Else
    For X = 0 To 2
        optsize(X).Visible = True
        Shape1.Visible = True
        Label1.Visible = True
    Next
    End If
End Sub

Private Sub optsize_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    For X = 0 To 2
        optsize(X).Value = 0
    Next
    optsize(Index).Value = 1
End Sub


