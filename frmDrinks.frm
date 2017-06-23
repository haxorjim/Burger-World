VERSION 5.00
Begin VB.Form frmDrinks 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Drinks"
   ClientHeight    =   4875
   ClientLeft      =   2115
   ClientTop       =   2790
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox LstTax 
      Height          =   1035
      ItemData        =   "frmDrinks.frx":0000
      Left            =   3600
      List            =   "frmDrinks.frx":0002
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ListBox LstSizes 
      Height          =   840
      ItemData        =   "frmDrinks.frx":0004
      Left            =   4560
      List            =   "frmDrinks.frx":0006
      TabIndex        =   25
      Top             =   5760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox TxtSizes 
      DataField       =   "Special"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   6720
      TabIndex        =   24
      Top             =   6360
      Width           =   1455
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Soda"
      Height          =   492
      Index           =   7
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1080
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Soda"
      Height          =   492
      Index           =   6
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1080
      Width           =   1692
   End
   Begin VB.TextBox TxtEnabled 
      DataField       =   "Enabled"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   6720
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtTax 
      DataField       =   "Tax"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   6720
      TabIndex        =   20
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtLarge 
      DataField       =   "Large"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5040
      TabIndex        =   19
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtMedium 
      DataField       =   "Medium"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   6360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtSmall 
      DataField       =   "Small"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5040
      TabIndex        =   17
      Top             =   5880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TxtName 
      DataField       =   "Drink Name"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox LstLarge 
      Height          =   1035
      ItemData        =   "frmDrinks.frx":0008
      Left            =   2520
      List            =   "frmDrinks.frx":000A
      TabIndex        =   15
      Top             =   5760
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ListBox LstMedium 
      Height          =   1035
      ItemData        =   "frmDrinks.frx":000C
      Left            =   1440
      List            =   "frmDrinks.frx":000E
      TabIndex        =   14
      Top             =   5760
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.ListBox LstSmall 
      Height          =   1035
      ItemData        =   "frmDrinks.frx":0010
      Left            =   360
      List            =   "frmDrinks.frx":0012
      TabIndex        =   13
      Top             =   5760
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CheckBox optsize 
      Caption         =   "Large"
      Height          =   492
      Index           =   2
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2280
      Width           =   1692
   End
   Begin VB.CheckBox optsize 
      Caption         =   "Medium"
      Height          =   492
      Index           =   1
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2280
      Width           =   1692
   End
   Begin VB.CheckBox optsize 
      Caption         =   "Small"
      Height          =   492
      Index           =   0
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2280
      Value           =   1  'Checked
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Milk Shakes"
      Height          =   492
      Index           =   5
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1080
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Orange Juice"
      Height          =   492
      Index           =   4
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bottled Water"
      Height          =   492
      Index           =   3
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   480
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hot Tea/Coffee"
      Height          =   492
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   480
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lemonade"
      Height          =   492
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Soda"
      Height          =   492
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   1692
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "menu.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   372
      Left            =   312
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Drinks"
      Top             =   4950
      Visible         =   0   'False
      Width           =   7332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   852
      Left            =   5400
      TabIndex        =   1
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton ComAdd 
      Caption         =   "Add this to the Order"
      Height          =   852
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Image logo 
      Height          =   1620
      Left            =   2970
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Beverages "
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   348
      TabIndex        =   9
      Top             =   132
      Width           =   900
   End
   Begin VB.Shape Shape2 
      Height          =   1488
      Left            =   192
      Top             =   228
      Width           =   7392
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sizes "
      ForeColor       =   &H80000008&
      Height          =   216
      Left            =   360
      TabIndex        =   8
      Top             =   1908
      Width           =   492
   End
   Begin VB.Shape Shape1 
      Height          =   972
      Left            =   240
      Top             =   2028
      Width           =   7356
   End
End
Attribute VB_Name = "frmDrinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cost As Single
Dim Size As String
Dim Drink As String
Private Sub Command1_Click()

End Sub

Private Sub ComAdd_Click()
Data2.Refresh
Data2.Recordset.MoveFirst
For Index = 0 To 7
    If OptDrink(Index).Value = True Then
    SelectedDrink = OptDrink(Index).Caption
    Exit For
    End If
Next
If frmDrinks.optsize(0).Value = 1 Then
Size = "S"
Listnum = 1
End If
If frmDrinks.optsize(1).Value = 1 Then
Size = "M"
Listnum = 2
End If
If frmDrinks.optsize(2).Value = 1 Then
Size = "L"
Listnum = 3
End If
If LstSizes.List(SideIndex) = "False" Then Size = ""
If Listnum = 1 Then Cost = LstSmall.List(Index)
If Listnum = 2 Then Cost = LstMedium.List(Index)
If Listnum = 3 Then Cost = LstLarge.List(Index)
Drink = Size & OptDrink(Index).Caption
For X = 0 To frmlunchmenu.List1.ListCount
    If Drink = frmlunchmenu.List1.List(X) Then
        frmlunchmenu.LstQuantity.List(X) = frmlunchmenu.LstQuantity.List(X) + 1
        Exit Sub
    End If
Next

If LstTax.List(Index) = "True" Then
Taxable = "True"
Else
Taxable = "False"
End If
frmlunchmenu.List1.AddItem (Drink)
frmlunchmenu.LstQuantity.AddItem (1)
frmlunchmenu.LstPrice.AddItem (FormatCurrency(Cost))
frmlunchmenu.LstUnformPrice.AddItem (Cost)
frmlunchmenu.LstIngredient.AddItem ("Nothing")
frmlunchmenu.LstRegName.AddItem ("Nothing")
frmlunchmenu.LstTax.AddItem (Taxable)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
    logo.Picture = LoadPicture(Logo_Image)
    Data2.Refresh
    For Index = 0 To 7
        OptDrink(Index).Caption = txtName.Text
        If TxtEnabled.Text = "True" Then OptDrink(Index).Visible = True
        If TxtEnabled.Text = "False" Then OptDrink(Index).Visible = False
        LstSmall.List(Index) = TxtSmall.Text
        LstMedium.List(Index) = TxtMedium.Text
        LstLarge.List(Index) = TxtLarge.Text
        LstSizes.List(Index) = TxtSizes.Text
        LstTax.List(Index) = TxtTax.Text
        Data2.Recordset.MoveNext
    Next
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
End Sub

Private Sub OptDrink_Click(Index As Integer)
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

Private Sub Text1_Change()

End Sub
