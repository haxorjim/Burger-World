VERSION 5.00
Begin VB.Form frmEditItem 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Edit this Item"
   ClientHeight    =   7635
   ClientLeft      =   1155
   ClientTop       =   1185
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox CheTaxable 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5640
      Width           =   1932
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   7
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Width           =   1692
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   6
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4800
      Width           =   1692
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   5
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4800
      Width           =   1692
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   4
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4800
      Width           =   1692
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   3
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3480
      Width           =   1692
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   2
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   1692
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   1
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   1692
   End
   Begin VB.CheckBox CheIng 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enabled"
      Height          =   372
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3480
      Width           =   1692
   End
   Begin VB.CheckBox Chedisabled 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5640
      Width           =   1932
   End
   Begin VB.Data Data2 
      Caption         =   "Click the Buttons to Cycle through the Available Menu"
      Connect         =   "Access"
      DatabaseName    =   "Menu.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Food"
      Top             =   6840
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CommandButton ComExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   7320
      TabIndex        =   14
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtEditPrice 
      DataField       =   "Price"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox txtEditName 
      DataField       =   "Item Name"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label LblTax 
      DataField       =   "Tax"
      DataSource      =   "Data2"
      Height          =   495
      Left            =   4080
      TabIndex        =   33
      Top             =   7440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 8"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   7
      Left            =   4440
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 7"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   6
      Left            =   3240
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 6"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   5
      Left            =   1920
      TabIndex        =   30
      Top             =   7320
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 5"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   4
      Left            =   600
      TabIndex        =   29
      Top             =   7440
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 4"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   3
      Left            =   4200
      TabIndex        =   28
      Top             =   6840
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 3"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   2
      Left            =   3000
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 2"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   1
      Left            =   1800
      TabIndex        =   26
      Top             =   6840
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label LblIng 
      DataField       =   "Ingredient 1"
      DataSource      =   "Data2"
      Height          =   372
      Index           =   0
      Left            =   600
      TabIndex        =   25
      Top             =   6840
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Label lblEnabled 
      DataField       =   "Enabled"
      DataSource      =   "Data2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   16
      Top             =   6600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Optional Ingredients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   13
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ketchup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   7200
      TabIndex        =   12
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tomato"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   11
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mushrooms"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2760
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Onion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Onion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Mustard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mustard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7200
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Mayo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Lettuce 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lettuce"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Cheese 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cheese"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Price"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Edit this Item's Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmEditItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RecordName As Variant
Private Sub CheEnable_Click()

End Sub

Private Sub CheIng_Click(Index As Integer)
If CheIng(Index).Value = 0 Then LblIng(Index).Caption = 1
If CheIng(Index).Value = 1 Then LblIng(Index).Caption = 2
Data2.Recordset.Edit
If CheIng(Index).Value = 0 Then CheIng(Index).Caption = "Enabled"
If CheIng(Index).Value = 1 Then CheIng(Index).Caption = "Disabled"

End Sub

Private Sub Chedisabled_Click()
If Chedisabled.Value = 0 Then lblEnabled.Caption = "1"
If Chedisabled.Value = 1 Then lblEnabled.Caption = "2"
Data2.Recordset.Edit
If Chedisabled.Value = 0 Then Chedisabled.Caption = "This item is ENABLED"
If Chedisabled.Value = 1 Then Chedisabled.Caption = "This item is DISABLED"


End Sub

Private Sub CheTaxable_Click()
If CheTaxable.Value = 0 Then LblTax.Caption = "True"
If CheTaxable.Value = 1 Then LblTax.Caption = "False"
Data2.Recordset.Edit
If CheTaxable.Value = 0 Then CheTaxable.Caption = "This item IS taxable"
If CheTaxable.Value = 1 Then CheTaxable.Caption = "This item is NOT taxable"
    
End Sub

Private Sub ComExit_Click()
Unload Me
End Sub

Private Sub ComFirst_Click()

End Sub


Private Sub ComLast_Click()

End Sub

Private Sub ComNext_Click()

End Sub

Private Sub Form_Load()
    'On Error Resume Next
    Data2.Refresh

    ' Set The Item Properties
    RecordName = Stack.Pop()
    Data2.Recordset.MoveFirst
    For X = 0 To 17
        If RecordName = txtEditName.Text Then
            For Index = 0 To 7
            If LblIng(Index).Caption = "1" Then
                CheIng(Index).Value = 0
            CheIng(Index).Caption = "Enabled"
            End If
           If LblIng(Index).Caption = "2" Then
                CheIng(Index).Value = 1
                CheIng(Index).Caption = "Disabled"
            End If
            Next
            If lblEnabled.Caption = "1" Then
                Chedisabled.Value = 0
                Chedisabled.Caption = "This item is ENABLED"
            ElseIf lblEnabled.Caption = "2" Then
                Chedisabled.Value = 1
                Chedisabled.Caption = "This item is DISABLED"
            End If
            If LblTax.Caption = "True" Then
                CheTaxable.Value = 0
            ElseIf LblTax.Caption = "False" Then
                CheTaxable.Value = 1
            End If
        Exit Sub
        End If
        Data2.Recordset.MoveNext
    Next
    
    
    
End Sub

Private Sub txtEditName_Change()
Data2.Recordset.Edit
End Sub

Private Sub txtEditPrice_Change()
If IsNumeric(txtEditPrice.Text) = False Then
txtEditPrice.Text = ""
Else
Data2.Recordset.Edit
End If
End Sub

Private Sub txtIngredients_Change(Index As Integer)

End Sub
