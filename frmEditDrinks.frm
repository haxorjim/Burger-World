VERSION 5.00
Begin VB.Form frmEditDrinks 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Drinks"
   ClientHeight    =   7725
   ClientLeft      =   780
   ClientTop       =   1215
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7725
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ComExit 
      Caption         =   "Exit"
      Height          =   855
      Left            =   7200
      TabIndex        =   26
      Top             =   4215
      Width           =   1575
   End
   Begin VB.Data Data2 
      Caption         =   "Click the Buttons to Cycle through the Available Menu"
      Connect         =   "Access"
      DatabaseName    =   "Menu.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   885
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Drinks"
      Top             =   5385
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.CheckBox CheSizes 
      BackColor       =   &H80000009&
      Caption         =   "Check1"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CheckBox CheTax 
      BackColor       =   &H80000009&
      Caption         =   "Check1"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CheckBox CheAvailable 
      BackColor       =   &H80000009&
      Caption         =   "Check1"
      DataField       =   "Special"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox TxtSizes 
      DataField       =   "Special"
      DataSource      =   "Data3"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   6360
      Width           =   1455
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   7
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   6
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   1692
   End
   Begin VB.TextBox TxtEnabled 
      DataField       =   "Enabled"
      DataSource      =   "Data3"
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtTax 
      DataField       =   "Tax"
      DataSource      =   "Data3"
      Height          =   375
      Left            =   6720
      TabIndex        =   11
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox TxtLarge 
      DataField       =   "Large"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox TxtMedium 
      DataField       =   "Medium"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   2640
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox TxtSmall 
      DataField       =   "Small"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox TxtName 
      DataField       =   "Drink Name"
      DataSource      =   "Data2"
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   5
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   4
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1080
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   3
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   480
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   2
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   1
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   480
      Width           =   1692
   End
   Begin VB.OptionButton OptDrink 
      BackColor       =   &H00FFFFFF&
      Height          =   492
      Index           =   0
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   1692
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Multiple Sizes?"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2880
      TabIndex        =   24
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Taxable?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Available?"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   20
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Large"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5400
      TabIndex        =   19
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Medium"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   3120
      TabIndex        =   18
      Top             =   3600
      Width           =   585
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Small"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   840
      TabIndex        =   17
      Top             =   3600
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drink Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   1800
      Width           =   975
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
      TabIndex        =   6
      Top             =   132
      Width           =   900
   End
   Begin VB.Shape Shape2 
      Height          =   1488
      Left            =   192
      Top             =   228
      Width           =   7392
   End
End
Attribute VB_Name = "frmEditDrinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim Cost As Single
'Dim Size As String
'Dim Drink As String
'Dim SelectedDrink As Integer
Private Sub ComExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
   Data2.Refresh
   Call Refreshmenu
End Sub
Public Sub Refreshmenu()
    Data2.Refresh
    Data2.Recordset.MoveFirst
    For Index = 0 To 7
        'MsgBox "Stop" & TxtName.Text
        OptDrink(Index).Caption = TxtName.Text
        If TxtEnabled.Text = "True" Then OptDrink(Index).BackColor = vbWhite
        If TxtEnabled.Text = "False" Then OptDrink(Index).BackColor = vbRed
        Data2.Recordset.MoveNext
    Next
End Sub
'Private Sub CheAvailable_Click()
'    If CheAvailable.Value = 0 Then
'        TxtEnabled.Text = "False"
'        CheAvailable.Caption = "No"
'    Else
'        TxtEnabled.Text = "True"
'        CheAvailable.Caption = "Yes"
'    End If
'    Call Refreshmenu
'End Sub
'Private Sub CheSizes_Click()
'    If CheSizes.Value = 0 Then
'        TxtSizes.Text = "False"
'        CheSizes.Caption = "No"
'        TxtSmall.Visible = False
'        TxtLarge.Visible = False
'        Label4.Visible = False
'        Label5.Caption = "Price"
'        Label6.Visible = False
'    Else
'        TxtSizes.Text = "True"
'        CheSizes.Caption = "Yes"
'        TxtSmall.Visible = True
'        TxtLarge.Visible = True
'        Label4.Visible = True
'        Label5.Caption = "Medium"
'        Label6.Visible = True
'    End If
'    Call Refreshmenu
'End Sub
'Private Sub CheTax_Click()
'    If CheTax.Value = 0 Then
'        TxtTax.Text = "False"
'        CheTax.Caption = "No"
'    Else
'        TxtTax.Text = "True"
'        CheTax.Caption = "Yes"
'    End If
'    Call Refreshmenu
'End Sub
'Private Sub OptDrink_Click(Index As Integer)
'    Data2.Refresh
'    Data2.Recordset.Move (Index)
'    SelectedDrink = Index
'    If TxtSizes.Text = "False" Then
'        For X = 0 To 2
'            TxtSmall.Visible = False
'            TxtLarge.Visible = False
'            Label4.Visible = False
'            Label5.Caption = "Price"
'            Label6.Visible = False
'        Next
'    Else
'        For X = 0 To 2
'            TxtSmall.Visible = True
'            TxtLarge.Visible = True
'            Label4.Visible = True
'            Label5.Caption = "Medium"
'            Label6.Visible = True
'        Next
'    End If
'End Sub
'Private Sub optdrink_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'    For X = 0 To 2
'        OptDrink(X).Value = 0
'    Next
'    OptDrink(Index).Value = 1
'End Sub
'Private Sub TxtEnabled_Change()
'    Call Refreshmenu
'    Data2.Recordset.Edit
'End Sub
'Private Sub TxtLarge_Change()
'    If IsNumeric(TxtLarge.Text) = True Then
'        Data2.Recordset.Edit
'    Else
'    TxtLarge.Text = ""
'    End If
'    Call Refreshmenu
'End Sub
'Private Sub TxtMedium_Change()
'    If IsNumeric(TxtMedium.Text) = True Then
'        Data2.Recordset.Edit
'    Else
'        TxtMedium.Text = ""
'    End If
'    Call Refreshmenu
'End Sub
'Private Sub TxtName_Change()
'    Call Refreshmenu
'    Data2.Recordset.Edit
'End Sub
'Private Sub TxtSizes_Change()
'    Data2.Recordset.Edit
'End Sub
'Private Sub TxtSmall_Change()
'    If IsNumeric(TxtSmall.Text) = True Then
'        Data2.Recordset.Edit
'    Else
'        TxtSmall.Text = ""
'    End If
'End Sub
'Private Sub TxtTax_Change()
'    Data2.Recordset.Edit
'End Sub
