VERSION 5.00
Begin VB.Form frmCheckOut 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   Caption         =   "Final Order Summary"
   ClientHeight    =   7185
   ClientLeft      =   765
   ClientTop       =   900
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton optToGo 
      Caption         =   "To Go"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5400
      Value           =   -1  'True
      Width           =   1284
   End
   Begin VB.OptionButton OptHere 
      Caption         =   "For Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4380
      Width           =   1284
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Order Finished"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   6360
      TabIndex        =   16
      Top             =   5400
      Width           =   2292
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   852
      Left            =   6360
      TabIndex        =   15
      Top             =   4200
      Width           =   2292
   End
   Begin VB.TextBox TxtUnformPrice 
      Height          =   288
      Left            =   4800
      TabIndex        =   14
      Top             =   6720
      Visible         =   0   'False
      Width           =   1932
   End
   Begin VB.ListBox LstUnformPrice 
      Height          =   255
      ItemData        =   "frmCheckOut.frx":0000
      Left            =   7200
      List            =   "frmCheckOut.frx":0002
      TabIndex        =   13
      Top             =   6720
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.ListBox LstQuantity 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmCheckOut.frx":0004
      Left            =   4920
      List            =   "frmCheckOut.frx":0006
      TabIndex        =   11
      Top             =   1416
      Width           =   1092
   End
   Begin VB.ListBox LstPrice 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmCheckOut.frx":0008
      Left            =   6000
      List            =   "frmCheckOut.frx":000A
      TabIndex        =   2
      Top             =   1416
      Width           =   1332
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      ItemData        =   "frmCheckOut.frx":000C
      Left            =   1800
      List            =   "frmCheckOut.frx":000E
      TabIndex        =   1
      Top             =   1416
      Width           =   3132
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Dine In Tax:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3000
      TabIndex        =   21
      Top             =   5112
      Width           =   1332
   End
   Begin VB.Label lblEatIn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "$0.00"
      Height          =   372
      Left            =   4560
      TabIndex        =   20
      Top             =   5088
      Width           =   1332
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Here or To Go?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   600
      TabIndex        =   19
      Top             =   3936
      Width           =   1740
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4920
      TabIndex        =   12
      Top             =   1056
      Width           =   1092
   End
   Begin VB.Label lblFTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   4560
      TabIndex        =   10
      Top             =   5556
      Width           =   1332
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3000
      TabIndex        =   9
      Top             =   5556
      Width           =   1332
   End
   Begin VB.Label lblTaxPercentage 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "(5.75%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3240
      TabIndex        =   8
      Top             =   4656
      Width           =   972
   End
   Begin VB.Label lbltax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   4560
      TabIndex        =   7
      Top             =   4656
      Width           =   1332
   End
   Begin VB.Label lblSubtotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   4560
      TabIndex        =   6
      Top             =   4176
      Width           =   1332
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Subtotal:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3240
      TabIndex        =   5
      Top             =   4176
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   6120
      TabIndex        =   4
      Top             =   1056
      Width           =   1212
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Items ordered"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1680
      TabIndex        =   3
      Top             =   1056
      Width           =   3252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Burger World Check Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   240
      TabIndex        =   0
      Top             =   432
      Width           =   8532
   End
End
Attribute VB_Name = "frmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TaxRate As Variant
Dim Totalcost As Variant
Dim where As String * 1
Private Sub cmd7_Click()
    txtPayAmount.Text = txtPayAmount.Text + "7"
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command10_Click()
    txtPayAmount.Text = txtPayAmount.Text + "3"
End Sub
Private Sub Command11_Click()
    txtPayAmount.Text = txtPayAmount.Text + "0"
End Sub
Private Sub Command12_Click()
    txtPayAmount.Text = txtPayAmount.Text + "."
End Sub
Private Sub Command14_Click()
    If Len(txtPayAmount.Text) <> 0 Then txtPayAmount.Text = Left(txtPayAmount.Text, Len(txtPayAmount.Text) - 1)
End Sub
Private Sub Command2_Click()
    'create file if not there
    Open "Order.txt" For Append As #1
    Close #1
    'find next order number
    Open "Order.txt" For Input As #1
        Do While Not EOF(1)
            Input #1, OrderNumber, nuttin, nuttin, nuttin, nuttin
        Loop
    Close #1
    If OrderNumber = 0 Then
        OrderNumber = 1
    Else
        OrderNumber = OrderNumber + 1
    End If
    
    'open for write
    Open "Order.txt" For Append As #1
        For Index = 0 To List1.ListCount
            If Trim(List1.List(Index)) <> "" Then
                Write #1, OrderNumber, List1.List(Index), LstQuantity.List(Index), Totalcost, where
            End If
        Next
    Close #1
    'end check out
    'MsgBox "thank you goes here"

    'MsgBox "Thank you for Eating at BurgerWorld!"
    Stack.Push "Thank you for Eating at BurgerWorld!"
    GUI.Load_Form frmMsgBox, FrmDesktop
    'Open "OrderSummary.Txt" For Append As #4
    '    Write #4, Val(lblFTotal.Caption), Change
    'Close #4
    Unload frmCheckOut
    Call frmlunchmenu.ClearMenu
    For X = 0 To 17
        frmlunchmenu.OptFood(X).Value = False
    Next
    For X = 0 To 7
        frmlunchmenu.CheIngredient(X).Value = False
        frmlunchmenu.CheIngredient(X).Enabled = False
    Next
    frmlunchmenu.ComAddto.Enabled = False
    frmlunchmenu.ComCheckout.Enabled = False
End Sub
Private Sub Command3_Click()
    txtPayAmount.Text = txtPayAmount.Text + "8"
End Sub
Private Sub Command4_Click()
    txtPayAmount.Text = txtPayAmount.Text + "9"
End Sub
Private Sub Command5_Click()
    txtPayAmount.Text = txtPayAmount.Text + "4"
End Sub
Private Sub Command6_Click()
    txtPayAmount.Text = txtPayAmount.Text + "5"
End Sub
Private Sub Command7_Click()
    txtPayAmount.Text = txtPayAmount.Text + "6"
End Sub
Private Sub Command8_Click()
    txtPayAmount.Text = txtPayAmount.Text + "1"
End Sub
Private Sub Command9_Click()
    txtPayAmount.Text = txtPayAmount.Text + "2"
End Sub
Private Sub Form_Load()
    'Call Full_Screen(Me)
    Open "TaxRate.Txt" For Input As #1
    Input #1, InpTaxRate
    Close #1
    where = "G"
    DispTaxRate = Val(InpTaxRate) * 100
    TaxRate = InpTaxRate
    lblTaxPercentage.Caption = "(" & DispTaxRate & "%)"
   For X = 0 To frmlunchmenu.List1.ListCount
        List1.List(X) = frmlunchmenu.List1.List(X)
        LstQuantity.List(X) = frmlunchmenu.LstQuantity.List(X)
        LstPrice.List(X) = frmlunchmenu.LstPrice.List(X)
        LstUnformPrice.List(X) = frmlunchmenu.LstUnformPrice.List(X)
    Next
    lblSubtotal.Caption = frmlunchmenu.TxtSub.Text
    LblTax.Caption = frmlunchmenu.TxtTax.Text
    lblFTotal.Caption = frmlunchmenu.TxtFinTotal.Text
    Totalcost = frmlunchmenu.TxtUnformPrice.Text
    TxtUnformPrice.Text = frmlunchmenu.TxtUnformPrice.Text
End Sub
Private Sub lblDispSubtotal_Click()
End Sub
Private Sub OptHere_Click()
    'TaxRate = 0.0575
    where = "H"
    eatintax = Val(frmlunchmenu.TxtUnformPrice.Text) * TaxRate
    lblEatIn.Caption = eatintax
    lblFTotal.Caption = Val(frmlunchmenu.TxtUnformsubtotal.Text) + eatintax
    Totalcost = Val(frmlunchmenu.TxtUnformsubtotal.Text) + eatintax
    lblEatIn.Caption = FormatCurrency(Val(lblEatIn.Caption))
    lblFTotal.Caption = FormatCurrency(Val(lblFTotal.Caption))
End Sub
Private Sub optToGo_Click()
    where = "G"
    eatintax = 0
    lblEatIn.Caption = eatintax
    lblFTotal.Caption = Val(frmlunchmenu.TxtUnformPrice.Text)
    Totalcost = Val(frmlunchmenu.TxtUnformPrice.Text)
    lblEatIn.Caption = FormatCurrency(Val(lblEatIn.Caption))
    lblFTotal.Caption = FormatCurrency(Val(lblFTotal.Caption))
End Sub
