VERSION 5.00
Begin VB.Form FrmKitchenMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   1005
   ClientTop       =   1065
   ClientWidth     =   9735
   Icon            =   "KitchenMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.CommandButton Exit 
      Caption         =   "Leave the Kitchen"
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
      Left            =   7080
      TabIndex        =   6
      Top             =   6000
      Width           =   2412
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel Order"
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
      Left            =   4680
      TabIndex        =   20
      Top             =   6000
      Width           =   2172
   End
   Begin VB.CommandButton Command1 
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
      Left            =   120
      TabIndex        =   19
      Top             =   6000
      Width           =   4332
   End
   Begin VB.TextBox TxtPrice 
      Height          =   372
      Index           =   6
      Left            =   7200
      TabIndex        =   18
      Top             =   6840
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox TxtPrice 
      Height          =   372
      Index           =   5
      Left            =   7080
      TabIndex        =   17
      Top             =   6840
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox TxtPrice 
      Height          =   372
      Index           =   4
      Left            =   7080
      TabIndex        =   16
      Top             =   6960
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox TxtPrice 
      Height          =   372
      Index           =   3
      Left            =   7080
      TabIndex        =   15
      Top             =   6960
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox TxtPrice 
      Height          =   372
      Index           =   2
      Left            =   7200
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox TxtPrice 
      Height          =   372
      Index           =   1
      Left            =   7440
      TabIndex        =   13
      Top             =   6840
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.TextBox TxtPrice 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   3240
      TabIndex        =   11
      Top             =   5448
      Width           =   1212
   End
   Begin VB.ListBox Order 
      ForeColor       =   &H00808080&
      Height          =   1425
      Index           =   4
      Left            =   7068
      TabIndex        =   9
      Top             =   2388
      Width           =   2412
   End
   Begin VB.ListBox Order 
      ForeColor       =   &H00808080&
      Height          =   1425
      Index           =   5
      Left            =   4560
      TabIndex        =   8
      Top             =   4080
      Width           =   2412
   End
   Begin VB.ListBox Order 
      ForeColor       =   &H00808080&
      Height          =   1425
      Index           =   6
      Left            =   7068
      TabIndex        =   7
      Top             =   4068
      Width           =   2415
   End
   Begin VB.ListBox Order 
      Height          =   3960
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   4320
   End
   Begin VB.ListBox Order 
      ForeColor       =   &H00808080&
      Height          =   1425
      Index           =   3
      Left            =   4560
      TabIndex        =   3
      Top             =   2400
      Width           =   2412
   End
   Begin VB.ListBox Order 
      ForeColor       =   &H00808080&
      Height          =   1425
      Index           =   2
      Left            =   7068
      TabIndex        =   2
      Top             =   720
      Width           =   2412
   End
   Begin VB.ListBox Order 
      ForeColor       =   &H00808080&
      Height          =   1425
      Index           =   1
      Left            =   4560
      TabIndex        =   1
      Top             =   720
      Width           =   2412
   End
   Begin VB.Timer MainTimer 
      Interval        =   10
      Left            =   2712
      Top             =   1956
   End
   Begin VB.Label lblEatWhere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   468
      Left            =   360
      TabIndex        =   21
      Top             =   5400
      Width           =   1332
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   12
      Top             =   5472
      Width           =   1296
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity - Item - Ingredients"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   792
      Width           =   4344
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order Queue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   4560
      TabIndex        =   5
      Top             =   120
      Width           =   4932
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Order Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   492
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4332
   End
End
Attribute VB_Name = "FrmKitchenMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filedateold As String
Public Sub deleteorder()
    Open "order.txt" For Input As #1           'open order file
    Open "temp order.txt" For Output As #2     'open temporary order file
        Do While Not EOF(1)                    'remove order one and write the rest to a temp file
            Input #1, OrderNumber, ItemDesc, Quantity, Totalcost, where                    'input the order number, item name, price
            If a <> 1 Then Write #2, OrderNumber, ItemDesc, Quantity, Totalcost, where     'write to a temp file
        Loop
    Close #1, #2
    Open "order.txt" For Output As #1         'open order file
    Open "temp order.txt" For Input As #2     'open temporary order file
    Do While Not EOF(2)
        Input #2, OrderNumber, ItemDesc, Quantity, Totalcost, where                      'input the order number, item name, price
        Write #1, (OrderNumber - 1), ItemDesc, Quantity, Totalcost, where                 'write the input to the new order.txt
    Loop
    Close #1, #2
    'test
End Sub
Private Sub Command1_Click()
    'Call GUI.Load_Form(FrmAreYouSure, Me)
    Command1.Enabled = False
    Call deleteorder
        Open "saleslog" For Append As #1
            For Index = 0 To Order(0).ListCount
                If Trim(Order(0).List(Index)) <> "" Then
                    Write #1, Order(0).List(Index), Date
                End If
            Next
        Close #1
    Command1.Enabled = True
End Sub
Private Sub Command2_Click()
    'cancel order
    Command2.Enabled = False
    Stack.Push "Are you sure you wish to delete this order?"
    GUI.Load_Form frmAreYouSure, FrmDesktop
    If Stack.Pop = "YES" Then
        Call deleteorder
    End If
    Command2.Enabled = True
End Sub
Private Sub Exit_Click()
    filedateold = vbNull
    Unload Me
End Sub

Private Sub Form_Load()
    Call GUI.Full_Screen(Me)
End Sub

Private Sub MainTimer_Timer()
    If filedateold <> FileDateTime("order.txt") Then
        filedateold = FileDateTime("order.txt")
        'reload order windows
        For Index = 0 To 6
            Call UpdateOrderWindows(Index)
        Next
        'hide order queue windows if empty
        For Index = 1 To 6
            If Order(Index).ListCount = 0 Then Order(Index).Visible = False
        Next
    End If
    If Order(0).ListCount = 0 Then
        Command1.Enabled = False
        Command2.Enabled = False
    Else
        Command1.Enabled = True
        Command2.Enabled = True
    End If
End Sub
Private Sub UpdateOrderWindows(Index)
    Dim Quantity As String * 4
    Order(Index).Clear
    Order(Index).Visible = True
    If Index = 0 Then TxtPrice(0).Text = ""
    If Index = 0 Then lblEatWhere.Caption = ""
    Open "order.txt" For Input As #1            'open order file
        Do While Not EOF(1)
            Input #1, OrderNumber, OrderedItem, Quantity, Totalcost, where
            If OrderNumber - 1 = Index Then
                If Index = 0 Then
                    TxtPrice(0).Text = FormatCurrency(Val(Totalcost))
                    If where = "G" Then
                        lblEatWhere.Caption = "To Go"
                    ElseIf where = "H" Then
                        lblEatWhere.Caption = "For Here"
                    End If
                End If
                OrderedItem = Quantity & " " & OrderedItem
                Order(Index).AddItem OrderedItem  'add the food item to the order list
            End If
        Loop
    Close #1
End Sub
