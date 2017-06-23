VERSION 5.00
Begin VB.Form SalesJournal 
   BorderStyle     =   0  'None
   Caption         =   "Sales Journal"
   ClientHeight    =   7290
   ClientLeft      =   975
   ClientTop       =   915
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   756
      Top             =   288
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "Clear Sales Journal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   240
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   6
      Top             =   5760
      Width           =   3372
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1212
      Left            =   6120
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   1
      Top             =   5760
      Width           =   3372
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   9252
   End
   Begin VB.Image Logo 
      Height          =   1620
      Left            =   3840
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sale Date"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   2052
      WordWrap        =   -1  'True
   End
   Begin VB.Label title 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sales Journal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   576
      Left            =   2160
      TabIndex        =   4
      Top             =   240
      Width           =   5016
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Quanity Sold"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   1452
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item Name"
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
      Left            =   3720
      TabIndex        =   2
      Top             =   1080
      Width           =   5772
   End
End
Attribute VB_Name = "SalesJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim filedateold As Date
Private Sub Command1_Click()
    Unload Me
End Sub
Private Sub Command2_Click()
    Stack.Push "Are you sure you wish to clear the Sales Journal?"
    GUI.Load_Form frmAreYouSure, FrmDesktop
    If Stack.Pop = "YES" Then
        Open "saleslog" For Output As #1
        Close #1
        List1.Clear
    End If
End Sub

Private Sub Form_Load()
    logo.Picture = LoadPicture(Logo_Image)
    title.Left = 100
    title.Width = Me.Width - 200
    title.Top = 200
    title.BackColor = &HFFFFFF
    Open "saleslog" For Input As #1
        Do While Not EOF(1)
            Input #1, qtyitem, saledate
            qty = Left(qtyitem, 4)
            Item = Mid(qtyitem, 6)
            saledate = Mid(saledate, 1)
            lineitem = saledate + String(30, " ") + qty + String(13, " ") + Item
            List1.AddItem lineitem
        Loop
    Close #1
End Sub

Private Sub Timer1_Timer()
    If List1.ListCount = 0 Then
        Command2.Enabled = False
    Else
        Command2.Enabled = True
    End If
    
    If filedateold <> FileDateTime("saleslog") Then
        filedateold = FileDateTime("saleslog")
        List1.Clear
        Open "saleslog" For Input As #1
        Do While Not EOF(1)
            Input #1, qtyitem, saledate
            qty = Left(qtyitem, 4)
            Item = Mid(qtyitem, 6)
            saledate = Mid(saledate, 1)
            lineitem = saledate + String(30, " ") + qty + String(13, " ") + Item
            List1.AddItem lineitem
        Loop
    Close #1
    End If
End Sub
