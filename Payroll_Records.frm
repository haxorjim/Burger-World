VERSION 5.00
Begin VB.Form Payroll_Records 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Payroll Record"
   ClientHeight    =   7368
   ClientLeft      =   2856
   ClientTop       =   432
   ClientWidth     =   6792
   Icon            =   "Payroll_Records.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7368
   ScaleWidth      =   6792
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Return 
      Caption         =   "Return"
      Height          =   612
      Left            =   2640
      TabIndex        =   5
      Top             =   1200
      Width           =   3972
   End
   Begin VB.CommandButton Display 
      Caption         =   "Display"
      Height          =   612
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1812
   End
   Begin VB.ComboBox Lstemployee 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      TabIndex        =   2
      Top             =   480
      Width           =   3972
   End
   Begin VB.ComboBox LstPeriod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "Payroll_Records.frx":030A
      Left            =   240
      List            =   "Payroll_Records.frx":030C
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Tax Information This Period "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   270
      TabIndex        =   24
      Top             =   5580
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      Height          =   1410
      Left            =   180
      Top             =   5670
      Width           =   6375
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "City Tax"
      Height          =   195
      Left            =   375
      TabIndex        =   23
      Top             =   5985
      Width           =   570
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "State Tax"
      Height          =   195
      Left            =   2535
      TabIndex        =   22
      Top             =   5985
      Width           =   690
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Tax"
      Height          =   195
      Left            =   4695
      TabIndex        =   21
      Top             =   5985
      Width           =   840
   End
   Begin VB.Label LblCity_tax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   375
      TabIndex        =   20
      Top             =   6585
      Width           =   390
   End
   Begin VB.Label lblState_Tax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      Height          =   195
      Left            =   2535
      TabIndex        =   19
      Top             =   6585
      Width           =   405
   End
   Begin VB.Label lblFederal_tax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      Height          =   195
      Left            =   4695
      TabIndex        =   18
      Top             =   6585
      Width           =   405
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pay This Period "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   270
      TabIndex        =   17
      Top             =   3735
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      Height          =   1605
      Left            =   180
      Top             =   3825
      Width           =   6375
   End
   Begin VB.Label lblGross_pay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   4710
      Width           =   405
   End
   Begin VB.Label lblnet_pay 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      Height          =   195
      Left            =   2640
      TabIndex        =   15
      Top             =   4710
      Width           =   405
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Gross Pay"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   4230
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Net Pay"
      Height          =   195
      Left            =   2640
      TabIndex        =   13
      Top             =   4230
      Width           =   1305
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Employee Info "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   300
      TabIndex        =   6
      Top             =   2100
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      Height          =   195
      Left            =   420
      TabIndex        =   12
      Top             =   2565
      Width           =   1155
   End
   Begin VB.Label lblEmployee_number 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2700
      TabIndex        =   11
      Top             =   3045
      Width           =   405
   End
   Begin VB.Label lblEmployee_name 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   10
      Top             =   3045
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Height          =   195
      Left            =   2700
      TabIndex        =   9
      Top             =   2565
      Width           =   1290
   End
   Begin VB.Label lblEmployee_hours 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Blank"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4740
      TabIndex        =   8
      Top             =   3045
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Hours"
      Height          =   195
      Left            =   4740
      TabIndex        =   7
      Top             =   2565
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      Height          =   1425
      Left            =   180
      Top             =   2190
      Width           =   6375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Employee"
      Height          =   192
      Left            =   2604
      TabIndex        =   3
      Top             =   192
      Width           =   1608
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Your Pay Period"
      Height          =   192
      Left            =   204
      TabIndex        =   0
      Top             =   192
      Width           =   1668
   End
End
Attribute VB_Name = "Payroll_Records"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Employee_data As Variant
Dim DateIndex As Integer
Dim switch As Boolean
Dim EmpNameIndex As Integer

Private Sub Display_Click()
With DataLink.rscmdGetInfo
.Open
Do Until .Fields("Name") = Lstemployee.Text And .Fields("payperiod") = LstPeriod.Text
.MoveNext
If .EOF Then
    'MsgBox "No Existing Records for that person!", vbCritical, "Error "
    Stack.Push "No Existing Records for that person!"
    GUI.Load_Form frmMsgBox, FrmDesktop
    .Close
    Exit Sub
End If
Loop
lblEmployee_name.Caption = .Fields("Name")
lblEmployee_number.Caption = .Fields("UID")
lblEmployee_hours.Caption = .Fields("Hoursworked")
lblGross_pay.Caption = .Fields("Grosspay")
lblnet_pay.Caption = .Fields("Netpay")
LblCity_tax.Caption = .Fields("citytax")
lblState_Tax.Caption = .Fields("statetax")
lblFederal_tax.Caption = .Fields("fedtax")
.Close

End With

End Sub
Private Sub Form_Load()
switch = False
'Open "D:\payroll\Payroll Record.rpt" For Input As #1
DataLink.rscmdGetFields.Open
Do While Not DataLink.rscmdGetFields.EOF
'Input #1, Employee_name, Employee_number, Hours, Gross_pay, Net_pay, Pay_Period, City_tax, State_tax, Federal_tax
    employee_name = DataLink.rscmdGetFields.Fields(0)
    Pay_period = DataLink.rscmdGetFields.Fields(1)
    DataLink.rscmdGetFields.MoveNext
    
Repeat = False
Repeat2 = False
If switch = False Then
        switch = True
        Lstemployee.AddItem (employee_name)
        LstPeriod.AddItem (Pay_period)
End If
For X = 0 To Lstemployee.ListCount
    If employee_name = Lstemployee.List(X) Then
    Repeat = True
    Exit For
    End If
Next X

For Y = 0 To LstPeriod.ListCount
    If Pay_period = LstPeriod.List(Y) Then
        Repeat2 = True
    Exit For
    End If
Next Y
If Repeat <> True Then Lstemployee.AddItem (employee_name)
If Repeat2 <> True Then LstPeriod.AddItem (Pay_period)
Loop
DataLink.rscmdGetFields.Close

'Close #1
End Sub


Private Sub LstPeriod_Change()
DateIndex = LstPeriod.ListIndex
End Sub

Private Sub Return_Click()
'Close #1
Unload Me
'Payroll_menu.Show

End Sub



Private Sub LstEmployee_Change()
EmpNameIndex = Lstemployee.ListIndex
End Sub
