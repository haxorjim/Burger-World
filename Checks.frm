VERSION 5.00
Begin VB.Form Payroll_Checks 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Pay Check"
   ClientHeight    =   6204
   ClientLeft      =   1788
   ClientTop       =   1812
   ClientWidth     =   9252
   Icon            =   "Checks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6204
   ScaleWidth      =   9252
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Return 
      Caption         =   "Return To Payroll Calculator"
      Height          =   1050
      Left            =   5520
      TabIndex        =   0
      Top             =   4440
      Width           =   2205
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Dave Varhol"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7425
      TabIndex        =   20
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year to Date Pay"
      Height          =   192
      Left            =   888
      TabIndex        =   19
      Top             =   4680
      Width           =   1236
   End
   Begin VB.Label Year_to_date 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   960
      TabIndex        =   18
      Top             =   5040
      Width           =   1092
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dependents"
      Height          =   192
      Left            =   2400
      TabIndex        =   17
      Top             =   4680
      Width           =   888
   End
   Begin VB.Label Depents 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   252
      Left            =   2400
      TabIndex        =   16
      Top             =   5040
      Width           =   1092
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Withholding Tax"
      Height          =   195
      Left            =   4080
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Witholding 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tax Information "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   540
      TabIndex        =   13
      Top             =   4125
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      Height          =   1536
      Left            =   456
      Top             =   4212
      Width           =   3516
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pay Check "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   540
      TabIndex        =   12
      Top             =   270
      Width           =   900
   End
   Begin VB.Shape Shape1 
      Height          =   3495
      Left            =   435
      Top             =   375
      Width           =   8430
   End
   Begin VB.Image Image1 
      Height          =   384
      Left            =   732
      Picture         =   "Checks.frx":030A
      Top             =   768
      Width           =   384
   End
   Begin VB.Label Pay_date 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   5895
      TabIndex        =   11
      Top             =   1005
      Width           =   390
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "1782 Crew Lane Mentor Ohio 44060"
      Height          =   435
      Left            =   735
      TabIndex        =   10
      Top             =   1485
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
      Height          =   195
      Left            =   7575
      TabIndex        =   9
      Top             =   765
      Width           =   1065
   End
   Begin VB.Label Check_number 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   7935
      TabIndex        =   8
      Top             =   1005
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Burger World"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1290
      TabIndex        =   7
      Top             =   735
      Width           =   2820
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Pay  To The Order Of"
      Height          =   435
      Left            =   735
      TabIndex        =   6
      Top             =   1965
      Width           =   930
   End
   Begin VB.Label Payment 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   7575
      TabIndex        =   5
      Top             =   2445
      Width           =   975
   End
   Begin VB.Label LblEmployee 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   735
      TabIndex        =   4
      Top             =   2445
      Width           =   4815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "For"
      Height          =   375
      Left            =   735
      TabIndex        =   3
      Top             =   3165
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   735
      X2              =   2055
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bi-Weekly Payroll"
      Height          =   195
      Left            =   735
      TabIndex        =   2
      Top             =   3405
      Width           =   1290
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   7095
      X2              =   8655
      Y1              =   3645
      Y2              =   3645
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      Height          =   195
      Left            =   5895
      TabIndex        =   1
      Top             =   765
      Width           =   345
   End
End
Attribute VB_Name = "Payroll_Checks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Pay_date.Caption = Date$
Open "Check Number.dat" For Input As #5
    Input #5, Number
Close #5
Open "Check Number.dat" For Output As #5
    If Number = 0 Then Number = Number + 1
        Check_number.Caption = Number
        Number = Number + 1
        Write #5, Number
        Witholding.Caption = Format$((Rnd(100) * 50), "currency")
        Depents.Caption = Stack.Pop()
        Year_to_date.Caption = Format(Stack.Pop(), "currency")
        Payment.Caption = Stack.Pop()
        LblEmployee.Caption = Stack.Pop()

'Open "D:\Payroll\Married.xls" For Input As #6
'Do While Not EOF(6)
'    Input #6, At_least, No_more, dep1, dep2, dep3, dep4, dep5, dep6, dep7, dep8, dep9, dep10
'Loop
End Sub

Private Sub Return_Click()
'Unload Checks
Unload Me
Close #4
Close #5
Close #6
'Payroll_Calculator.Show

End Sub

