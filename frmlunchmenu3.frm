VERSION 5.00
Begin VB.Form frmlunchmenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Menu"
   ClientHeight    =   7092
   ClientLeft      =   60
   ClientTop       =   516
   ClientWidth     =   11448
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7092
   ScaleWidth      =   11448
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtUnformsubtotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   8100
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   8685
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ListBox LstTax 
      Height          =   2352
      ItemData        =   "frmlunchmenu3.frx":0000
      Left            =   7680
      List            =   "frmlunchmenu3.frx":0002
      TabIndex        =   68
      Top             =   6960
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8010
      Top             =   4620
   End
   Begin VB.TextBox TxtUnformPrice 
      Height          =   288
      Left            =   7620
      TabIndex        =   61
      Top             =   9060
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton ComTempExit 
      Caption         =   "Exit"
      Height          =   612
      Left            =   8955
      TabIndex        =   60
      Top             =   7440
      Width           =   2412
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   17
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   3255
      Width           =   852
   End
   Begin VB.ListBox LstOrgPrice 
      Height          =   240
      ItemData        =   "frmlunchmenu3.frx":0004
      Left            =   8580
      List            =   "frmlunchmenu3.frx":0006
      TabIndex        =   58
      Top             =   8820
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Data Data2 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "Menu.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   8640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Food"
      Top             =   9840
      Visible         =   0   'False
      Width           =   912
   End
   Begin VB.ListBox LstRegName 
      Height          =   240
      ItemData        =   "frmlunchmenu3.frx":0008
      Left            =   8880
      List            =   "frmlunchmenu3.frx":000A
      TabIndex        =   57
      Top             =   9240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.ListBox LstUnformPrice 
      Height          =   240
      ItemData        =   "frmlunchmenu3.frx":000C
      Left            =   8580
      List            =   "frmlunchmenu3.frx":000E
      TabIndex        =   56
      Top             =   9060
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.ListBox LstIngredient 
      Height          =   624
      ItemData        =   "frmlunchmenu3.frx":0010
      Left            =   8640
      List            =   "frmlunchmenu3.frx":0012
      TabIndex        =   55
      Top             =   9120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.TextBox TxtFinTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   10170
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox TxtTax 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   10170
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox TxtSub 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   10170
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   4185
      Width           =   1215
   End
   Begin VB.CommandButton ComClear 
      BackColor       =   &H000000FF&
      Caption         =   "Clear the Order List"
      Height          =   375
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6600
      Width           =   2412
   End
   Begin VB.CommandButton ComCheckout 
      BackColor       =   &H00FFFF00&
      Caption         =   "Checkout"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   612
      Left            =   5745
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6360
      Width           =   1455
   End
   Begin VB.ListBox LstPrice 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2352
      ItemData        =   "frmlunchmenu3.frx":0014
      Left            =   10620
      List            =   "frmlunchmenu3.frx":0016
      TabIndex        =   46
      Top             =   1305
      Width           =   732
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Cheese"
      DataField       =   "Ingredient 1"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   0
      Left            =   705
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5160
      Width           =   852
   End
   Begin VB.CommandButton Command5 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit to Main Menu"
      Height          =   624
      Left            =   11520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10800
      UseMaskColor    =   -1  'True
      Width           =   2616
   End
   Begin VB.CommandButton ComSides 
      Caption         =   "Add Fries and Other Sides"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton ComRemove 
      BackColor       =   &H000000FF&
      Caption         =   "Remove Selected Item(s)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6120
      Width           =   2412
   End
   Begin VB.CommandButton ComAddto 
      BackColor       =   &H0000FF00&
      Caption         =   "Add to Order"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5160
      Width           =   1455
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   16
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2175
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   15
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1095
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   14
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3255
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   13
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2175
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   12
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1095
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   6
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1095
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   11
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3255
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   10
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2175
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   9
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1095
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   8
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3255
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   7
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2175
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   5
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   3255
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   4
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2175
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   3
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1095
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   2
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3255
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2175
      Width           =   852
   End
   Begin VB.CommandButton ComDrinks 
      Caption         =   "Add a Drink"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7665
      Width           =   1455
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Ketchup"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   7
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   852
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Tomatoes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   6
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      Width           =   852
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Mushrooms"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   5
      Left            =   1785
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   852
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Mustard"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   3
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5160
      Width           =   852
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Mayonaise"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   2
      Left            =   2865
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   852
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Lettuce"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   1
      Left            =   1785
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   852
   End
   Begin VB.OptionButton OptFood 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   852
      Index           =   0
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1095
      Width           =   852
   End
   Begin VB.ListBox LstQuantity 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2352
      ItemData        =   "frmlunchmenu3.frx":0018
      Left            =   10260
      List            =   "frmlunchmenu3.frx":001A
      MousePointer    =   2  'Cross
      TabIndex        =   45
      Top             =   1305
      Width           =   372
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2352
      ItemData        =   "frmlunchmenu3.frx":001C
      Left            =   7620
      List            =   "frmlunchmenu3.frx":001E
      MousePointer    =   2  'Cross
      TabIndex        =   8
      Top             =   1305
      Width           =   2655
   End
   Begin VB.CheckBox CheIngredient 
      BackColor       =   &H008080FF&
      Caption         =   "Onions"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Index           =   4
      Left            =   705
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      Width           =   852
   End
   Begin VB.Label Lbltax 
      Caption         =   "Label11"
      DataField       =   "Tax"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   7080
      TabIndex        =   69
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label LblEnabled 
      Caption         =   "Label11"
      DataField       =   "Enabled"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   7560
      TabIndex        =   67
      Top             =   8520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10770
      TabIndex        =   66
      Top             =   840
      Width           =   450
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10290
      TabIndex        =   65
      Top             =   840
      Width           =   330
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7635
      TabIndex        =   64
      Top             =   840
      Width           =   1410
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Lunch Menu Selections   "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   585
      TabIndex        =   63
      Top             =   285
      Width           =   5310
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "   Order Summary   "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7635
      TabIndex        =   62
      Top             =   270
      Width           =   3690
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Total:"
      Height          =   255
      Left            =   9210
      TabIndex        =   54
      Top             =   5235
      Width           =   855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Tax:"
      Height          =   255
      Left            =   9195
      TabIndex        =   52
      Top             =   4770
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      Height          =   255
      Left            =   9210
      TabIndex        =   51
      Top             =   4275
      Width           =   735
   End
   Begin VB.Label lblSelectedPrice 
      Caption         =   "Selected Price"
      DataMember      =   "Menu"
      DataSource      =   "Datalink"
      Height          =   372
      Left            =   10080
      TabIndex        =   44
      Top             =   10800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblSelectedItem 
      Caption         =   "selected Item"
      DataMember      =   "Menu"
      DataSource      =   "Datalink"
      Height          =   372
      Left            =   8880
      TabIndex        =   43
      Top             =   10800
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      Caption         =   "Ingredient 8"
      DataField       =   "Ingredient 8"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   7
      Left            =   8760
      TabIndex        =   42
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      Caption         =   "Ingredient 7"
      DataField       =   "Ingredient 7"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   6
      Left            =   7560
      TabIndex        =   41
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      Caption         =   "Ingredient 6"
      DataField       =   "Ingredient 6"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   5
      Left            =   6480
      TabIndex        =   40
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      Caption         =   "Ingredient 5"
      DataField       =   "Ingredient 5"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   4
      Left            =   5280
      TabIndex        =   39
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      Caption         =   "Ingredient 4"
      DataField       =   "Ingredient 4"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   3
      Left            =   4200
      TabIndex        =   38
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      Caption         =   "Ingredient 3"
      DataField       =   "Ingredient 3"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   2
      Left            =   3120
      TabIndex        =   37
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      Caption         =   "Ingredient 2"
      DataField       =   "Ingredient 2"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   1
      Left            =   1920
      TabIndex        =   36
      Top             =   12000
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIngredient1 
      DataField       =   "Ingredient 1"
      DataSource      =   "Data2"
      Height          =   375
      Index           =   0
      Left            =   8520
      TabIndex        =   35
      Top             =   9120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPrice 
      Caption         =   "Price"
      DataField       =   "Price"
      DataSource      =   "Data2"
      Height          =   372
      Left            =   10200
      TabIndex        =   34
      Top             =   11640
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Label Label5 
      Caption         =   "Ingredient"
      DataField       =   "Item Name"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   252
      Left            =   10080
      TabIndex        =   33
      Top             =   11760
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.Label lblItem 
      DataField       =   "Item Name"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8580
      TabIndex        =   32
      Top             =   8820
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Make it with or without...."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1305
      TabIndex        =   7
      Top             =   4800
      Width           =   2655
   End
End
Attribute VB_Name = "frmlunchmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' This portion of the program was written by Jim Graf
' AKA Jimmy Chan
Dim AvailItem(17, 2) As Variant
Dim ItemCount As Long
Dim ActiveIng(7) As Boolean
Dim OptFoodActive As Boolean
Dim OrderedItem(127, 3) As Variant
Dim DefaultIng As Variant
Dim TaxRate As Variant
Public Sub PriceTotals()
    For X = 0 To LstPrice.ListCount
        subtotal = subtotal + (Val(LstUnformPrice.List(X)) * Val(LstQuantity.List(X)))
        If LstTax.List(X) = "False" Then
            Tax = 0
        Else
            Tax = Val(LstUnformPrice.List(X)) * Val(LstQuantity.List(X)) * TaxRate
        End If
        TotalTax = TotalTax + Tax
    Next
    FinalTotal = subtotal + TotalTax
    TxtUnformPrice.Text = FinalTotal
    TxtUnformsubtotal.Text = subtotal
        TxtSub.Text = FormatCurrency(subtotal)
        TxtTax.Text = FormatCurrency(TotalTax)
        TxtFinTotal = FormatCurrency(FinalTotal)
        
End Sub

Private Sub ComAddto_Click()
    
    On Error Resume Next
    Dim SelectedItem As String
    Dim SelectedPrice As Double
    Dim Specialorder As Boolean
    Dim SelectedIng As Double
    
For Index = 0 To 17
    If OptFood(Index).Value = True Then Exit For
Next
    ' Database stuff
    Data2.Recordset.MoveFirst
    Data2.Recordset.Move (Index)
    SelectedItem = AvailItem(Index, 1)
    SelectedPrice = AvailItem(Index, 2)
    
' If the item is to be made WITHOUT...
For IngCheck = 0 To 7
    If CheIngredient(IngCheck).Value = 1 And ActiveIng(IngCheck) = True Then
            SelectedItem = SelectedItem & " -"
            Exit For
    End If
Next
    ' Ingredient Check

For IngCheck = 0 To 7
    If CheIngredient(IngCheck).Value = 1 And ActiveIng(IngCheck) = True Then
        Select Case IngCheck
            Case Is = 0
                SelectedItem = SelectedItem & "CH"
            Case Is = 1
                SelectedItem = SelectedItem & "LE"
            Case Is = 2
                SelectedItem = SelectedItem & "MA"
            Case Is = 3
                SelectedItem = SelectedItem & "MU"
            Case Is = 4
                SelectedItem = SelectedItem & "ON"
            Case Is = 5
                SelectedItem = SelectedItem & "MS"
            Case Is = 6
                SelectedItem = SelectedItem & "TO"
            Case Is = 7
                SelectedItem = SelectedItem & "KE"
        End Select
    End If
Next

' If the item is to be made WITH...
For IngCheck = 0 To 7
    If CheIngredient(IngCheck).Value = 0 And ActiveIng(IngCheck) = False Then
            SelectedItem = SelectedItem & " +"
            Exit For
    End If
Next
    ' Ingredient Check

For IngCheck = 0 To 7
    If CheIngredient(IngCheck).Value = 0 And ActiveIng(IngCheck) = False Then
        Select Case IngCheck
            Case Is = 0
                SelectedItem = SelectedItem & "CH"
            Case Is = 1
                SelectedItem = SelectedItem & "LE"
            Case Is = 2
                SelectedItem = SelectedItem & "MA"
            Case Is = 3
                SelectedItem = SelectedItem & "MU"
            Case Is = 4
                SelectedItem = SelectedItem & "ON"
            Case Is = 5
                SelectedItem = SelectedItem & "MS"
            Case Is = 6
                SelectedItem = SelectedItem & "TO"
            Case Is = 7
                SelectedItem = SelectedItem & "KE"
        End Select
    End If
Next
    
' Check for repeat order

    For X = 0 To List1.ListCount
            If SelectedItem = List1.List(X) Then
                LstQuantity.List(X) = Val(LstQuantity.List(X)) + 1
                LstPrice.List(X) = FormatCurrency(Val(LstQuantity.List(X)) * Val(LstUnformPrice.List(X)))
                Call PriceTotals
                Exit Sub
            End If
    Next
        
List1.AddItem (SelectedItem)
LstQuantity.AddItem (1)
LstPrice.AddItem (FormatCurrency(SelectedPrice))
LstUnformPrice.AddItem (SelectedPrice)
LstIngredient.AddItem (DefaultIng)
LstSpecialNum.AddItem ("Norm")
LstRegName.AddItem ("Nothing")
LstTax.AddItem (lbltax.Caption)
    Call PriceTotals

    
End Sub
Private Sub Command2_Click()
    frmCombo.Show 1
End Sub
Private Sub Command3_Click()
    frmDrinks.Show 1
End Sub
Private Sub Command4_Click()

End Sub

Private Sub ComCombo_Click()
    frmCombo.Show 1
End Sub

Private Sub ComCheckout_Click()
    'frmCheckOut.Show 1
    Call GUI.Load_Form(frmCheckOut, Me)
    
End Sub

Private Sub ComDrinks_Click()
    Call GUI.Load_Form(frmDrinks, Me)
    'frmDrinks.Show 1
    Call PriceTotals
End Sub
Private Sub ComClear_Click()

Call ClearMenu
End Sub
Public Sub ClearMenu()
Do While List1.ListCount > 0
    List1.RemoveItem (0)
    LstQuantity.RemoveItem (0)
    LstPrice.RemoveItem (0)
    LstIngredient.RemoveItem (0)
    LstRegName.RemoveItem (0)
    LstUnformPrice.RemoveItem (0)
Loop
Call PriceTotals
End Sub
Private Sub ComEditMenu_Click()
    frmEditMainMenu.Show
End Sub

Private Sub ComTempExit_Click()
    'MsgBox "not yet password protected"
    Unload Me
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub ComRemove_Click()
On Error Resume Next
If Val(LstQuantity.List(List1.ListIndex)) < 2 Then
    List1.RemoveItem (List1.ListIndex)
    LstQuantity.RemoveItem (LstQuantity.ListIndex)
    LstPrice.RemoveItem (LstPrice.ListIndex)
    LstIngredient.RemoveItem (LstIngredient.ListIndex)
    LstRegName.RemoveItem (LstRegName.ListIndex)
    LstSpecialNum.RemoveItem (LstSpecialNum.ListIndex)
    LstUnformPrice.RemoveItem (LstUnformPrice.ListIndex)
    ItemCount = ItemCount - 1
Else
LstQuantity.List(List1.ListIndex) = Val(LstQuantity.List(List1.ListIndex)) - 1
LstPrice.List(List1.ListIndex) = FormatCurrency(Val(LstQuantity.List(List1.ListIndex)) * Val(LstUnformPrice.List(List1.ListIndex)))
End If

Call PriceTotals
End Sub

Private Sub ComSides_Click()
    Call GUI.Load_Form(frmSides, Me)
    'frmSides.Show 1
    Call PriceTotals
End Sub
Private Sub Option1_Click(Index As Integer)

End Sub

Private Sub Form_Load()
       
    Open "Taxrate.Txt" For Input As #1
    Input #1, InpTaxRate
    Close #1
    TaxRate = InpTaxRate
    ' Window Size adjustments
    Call GUI.Full_Screen(Me)

    ' DO NOT TOUCH! It Works! :-)
    Call PriceTotals
    Data2.Refresh
    Data2.Recordset.MoveFirst
    ' Set Option button captions to the database values
    For Index = 0 To 17
        OptFood(Index).Caption = lblItem.Caption
        If lblEnabled.Caption = "1" Then OptFood(Index).Visible = True
        If lblEnabled.Caption = "2" Then OptFood(Index).Visible = False
        Data2.Recordset.MoveNext
    Next
    ' Assign Database values to a 3D Array
    Data2.Recordset.MoveFirst
    For Index = 0 To 17
        AvailItem(Index, 1) = lblItem.Caption
        AvailItem(Index, 2) = lblPrice.Caption
        Data2.Recordset.MoveNext
    Next
    
End Sub

Private Sub LblIngredient4_Click()

End Sub

Private Sub LblIngredient8_Click()

End Sub


Private Sub List1_Click()
    Index = List1.ListIndex
    'MsgBox Index
    LstQuantity.Selected(Index) = True
    LstPrice.Selected(Index) = True
    LstIngredient.Selected(Index) = True
    LstRegName.Selected(Index) = True
    LstUnformPrice.Selected(Index) = True
    ItemCount = List1.ListIndex

End Sub

Private Sub LstPrice_Click()
    Index = LstPrice.ListIndex
    'MsgBox Index
    LstQuantity.Selected(Index) = True
    List1.Selected(Index) = True
    LstIngredient.Selected(Index) = True
    LstRegName.Selected(Index) = True
    LstUnformPrice.Selected(Index) = True
    ItemCount = List1.ListIndex
End Sub

Private Sub LstQuantity_Click()
    Index = LstQuantity.ListIndex
    'MsgBox Index
    List1.Selected(Index) = True
    LstPrice.Selected(Index) = True
    LstIngredient.Selected(Index) = True
    LstRegName.Selected(Index) = True
    LstUnformPrice.Selected(Index) = True
    ItemCount = LstQuantity.ListIndex
End Sub

Private Sub OptFood_Click(Index As Integer)
    'clear
    OptFoodActive = True
    ComAddto.Enabled = True
    ' Set the Buttons to deactive
    For SwitchBack = 0 To 7
        CheIngredient(SwitchBack).Value = 0
    Next
    
    ' Activate the Buttons
    For X = 0 To 7
        CheIngredient(X).Enabled = True
    Next
    Data2.Refresh
    Data2.Recordset.Move (Index)
    ' Set Default Ingredients
    For X = 0 To 7
        If LblIngredient1(X).Caption = "1" Then
            CheIngredient(X).Value = 0
            ActiveIng(X) = 1
        End If
        If LblIngredient1(X).Caption = "2" Then
            CheIngredient(X).Value = 1
            ActiveIng(X) = 0
        End If
    Next
        
End Sub

Private Sub Timer1_Timer()
    If List1.ListCount <> 0 Then
        ComCheckout.Enabled = True
    Else
        ComCheckout.Enabled = False
    End If
End Sub
