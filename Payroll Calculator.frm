VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Payroll_Calculator 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Payroll Calculator"
   ClientHeight    =   7320
   ClientLeft      =   915
   ClientTop       =   810
   ClientWidth     =   7575
   Icon            =   "Payroll Calculator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Hours 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3015
      TabIndex        =   24
      Top             =   984
      Width           =   1095
   End
   Begin VB.CommandButton Clear 
      Caption         =   "&Clear"
      Height          =   495
      Left            =   3015
      TabIndex        =   23
      Top             =   2184
      Width           =   1095
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Calculate"
      Default         =   -1  'True
      Height          =   495
      Left            =   3015
      TabIndex        =   22
      Top             =   1464
      Width           =   1095
   End
   Begin VB.CommandButton cmdLast 
      Height          =   735
      Left            =   5430
      Picture         =   "Payroll Calculator.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdNext 
      Height          =   735
      Left            =   4470
      Picture         =   "Payroll Calculator.frx":034E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   735
      Left            =   3750
      Picture         =   "Payroll Calculator.frx":0690
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   735
      Left            =   2790
      Picture         =   "Payroll Calculator.frx":09D2
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Write_checks 
      Caption         =   "Write this Check"
      Height          =   615
      Left            =   4740
      TabIndex        =   1
      Top             =   2250
      Width           =   2295
   End
   Begin VB.CommandButton Return 
      Caption         =   "&Exit"
      Height          =   735
      Left            =   6150
      TabIndex        =   0
      Top             =   6240
      Width           =   1095
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   255
      Left            =   10215
      TabIndex        =   2
      Top             =   2535
      Visible         =   0   'False
      Width           =   495
      _Version        =   524288
      _ExtentX        =   873
      _ExtentY        =   450
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2000
      Month           =   2
      Day             =   12
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.51
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.01
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image logo 
      Height          =   1620
      Left            =   4785
      Top             =   330
      Width           =   2175
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Taxes "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2925
      TabIndex        =   40
      Top             =   4665
      Width           =   555
   End
   Begin VB.Shape Shape4 
      Height          =   1365
      Left            =   2805
      Top             =   4770
      Width           =   4425
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Federal Tax"
      Height          =   195
      Left            =   5910
      TabIndex        =   39
      Top             =   5175
      Width           =   870
   End
   Begin VB.Label Federal_tax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5910
      TabIndex        =   38
      Top             =   5535
      Width           =   1095
   End
   Begin VB.Label State_tax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4470
      TabIndex        =   37
      Top             =   5535
      Width           =   1095
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "State Tax"
      Height          =   195
      Left            =   4470
      TabIndex        =   36
      Top             =   5175
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City Tax"
      Height          =   195
      Left            =   3030
      TabIndex        =   35
      Top             =   5175
      Width           =   570
   End
   Begin VB.Label City_tax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3030
      TabIndex        =   34
      Top             =   5535
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Earnings "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2895
      TabIndex        =   27
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Pay"
      Height          =   195
      Left            =   3000
      TabIndex        =   33
      Top             =   3585
      Width           =   750
   End
   Begin VB.Label Gross_Pay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3000
      TabIndex        =   32
      Top             =   3945
      Width           =   1095
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Net Pay"
      Height          =   195
      Left            =   4440
      TabIndex        =   31
      Top             =   3585
      Width           =   570
   End
   Begin VB.Label Net_Pay 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4440
      TabIndex        =   30
      Top             =   3945
      Width           =   1095
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period"
      Height          =   195
      Left            =   5880
      TabIndex        =   29
      Top             =   3585
      Width           =   810
   End
   Begin VB.Label Pay_period 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   3945
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      Height          =   1365
      Left            =   2775
      Top             =   3210
      Width           =   4425
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Time Sheet "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   2865
      TabIndex        =   26
      Top             =   240
      Width           =   930
   End
   Begin VB.Shape Shape1 
      Height          =   2670
      Left            =   2775
      Top             =   330
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hours Worked"
      Height          =   195
      Left            =   3015
      TabIndex        =   25
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   21
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rate Of Pay"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   20
      Top             =   2280
      Width           =   1290
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   19
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   18
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   750
      TabIndex        =   17
      Top             =   3480
      Width           =   1005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblWage 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   195
      Left            =   750
      TabIndex        =   16
      Top             =   2640
      Width           =   390
   End
   Begin VB.Label lblUID 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   750
      TabIndex        =   15
      Top             =   1800
      Width           =   390
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   750
      TabIndex        =   14
      Top             =   960
      Width           =   390
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   750
      TabIndex        =   13
      Top             =   4320
      Width           =   1005
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dependents"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   12
      Top             =   3960
      Width           =   1320
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   750
      TabIndex        =   11
      Top             =   5160
      Width           =   1005
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Earnings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   10
      Top             =   4800
      Width           =   945
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "blank"
      Height          =   195
      Left            =   750
      TabIndex        =   9
      Top             =   6000
      Width           =   1005
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maritial Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   510
      TabIndex        =   8
      Top             =   5640
      Width           =   1470
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Employee Record "
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   375
      TabIndex        =   7
      Top             =   210
      Width           =   1410
   End
   Begin VB.Shape Shape2 
      Height          =   6660
      Left            =   270
      Top             =   330
      Width           =   2295
   End
End
Attribute VB_Name = "Payroll_Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calculate_Click()
If Val(Hours.Text) > 336 Then
    'MsgBox "Too Many Hours! Please Re-Enter", vbOKOnly, "Too Many Hours"
    Stack.Push "Too Many Hours! Please Re-Enter"
    GUI.Load_Form frmMsgBox, FrmDesktop
    Hours.Text = ""
    Gross_Pay.Caption = ""
    Hours.SetFocus
ElseIf Val(Hours.Text) > 0 Then
    If Val(Hours.Text) > 80 Then
        Regular = lblWage.Caption * 80
        Overtime = (Hours.Text - 80) * (lblWage.Caption * 1.5)
        Gross_Pay.Caption = Format$(Regular + Overtime, "currency")
    Else
        Gross_Pay.Caption = Format$((Val(Hours.Text) * lblWage.Caption), "currency")
    End If
        Federal_tax.Caption = Format$(Gross_Pay.Caption * 0.08, "currency")
        State_tax.Caption = Format$(Gross_Pay.Caption * 0.02, "currency")
        City_tax.Caption = Format$(Gross_Pay.Caption * 0.01, "Currency")
        Net_Pay.Caption = Format$(Gross_Pay.Caption - (Federal_tax.Caption) - (State_tax.Caption) - (City_tax.Caption), "currency")
        'Write #1, lblName.Caption; lblUID.Caption; Hours.Text, Gross_Pay.Caption, Net_Pay.Caption, Pay_period.Caption, City_tax.Caption, State_tax.Caption, Federal_tax.Caption
SQL = "INSERT INTO Payroll (UID, HoursWorked, GrossPay, Netpay, PayPeriod, CityTax, StateTax, FedTax) VALUES ('" _
            & lblUID.Caption & "', '" _
            & Hours.Text & "', '" _
            & Gross_Pay.Caption & "', '" _
            & Net_Pay.Caption & "', '" _
            & Pay_period.Caption & "', '" _
            & City_tax.Caption & "', '" _
            & State_tax.Caption & "', '" _
            & Federal_tax.Caption & "')"
    DataLink.EmpDB.Execute SQL
  ' Year_to_date = Year_to_date + Gross_Pay.Caption
End If
End Sub
Private Sub Clear_Click()
    Hours.Text = ""
    Gross_Pay.Caption = ""
    Net_Pay.Caption = ""
    City_tax.Caption = ""
    State_tax.Caption = ""
    Federal_tax.Caption = ""
    Hours.SetFocus
End Sub
Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError
  DataLink.rsEmployee.MoveFirst
  Call Display_Data
  Call Clear_Click
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub cmdLast_Click()
On Error GoTo GoLastError
  DataLink.rsEmployee.MoveLast
  Call Display_Data
  Call Clear_Click
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
'On Error GoTo GoNextError
With DataLink.rsEmployee
  If Not .EOF Then .MoveNext
  If .EOF And .RecordCount > 0 Then
    Beep
    .MoveLast
  End If
End With
Call Display_Data
Call Clear_Click
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub cmdPrevious_Click()
On Error GoTo GoPrevError
With DataLink.rsEmployee
  If Not .BOF Then .MovePrevious
  If .BOF And .RecordCount > 0 Then
    Beep
    .MoveFirst
  End If
End With
Call Display_Data
Call Clear_Click
Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub
Private Sub Display_Data()
  lblUID.Caption = DataLink.rsEmployee.Fields("UID")
  lblName.Caption = DataLink.rsEmployee.Fields("Name")
  lblWage.Caption = DataLink.rsEmployee.Fields("wage")
  lblTitle.Caption = DataLink.rsEmployee.Fields("Title")
  Label17.Caption = DataLink.rsEmployee.Fields("Dependents")
  Label19.Caption = Format$(DataLink.rsEmployee.Fields("YTDEarnings"), "currency")
  If UCase$(DataLink.rsEmployee.Fields("MaritalStatus")) = "M" Then Label21.Caption = "Married" Else Label21.Caption = "Single"
End Sub
Private Sub Form_Activate()
    Hours.SetFocus
    logo.Picture = LoadPicture(Logo_Image)
End Sub
Private Sub Form_Load()
    DataLink.rsEmployee.Open
    Call Display_Data
    Pay_period.Caption = Date$
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DataLink.rsEmployee.Close
End Sub
Private Sub Hours_KeyPress(KeyAscii As Integer)
    'LIMIT CHARACTERS TO 0-9, BACKSPACE AND "."
    If KeyAscii = 27 Then
        Gross_Pay.Caption = ""
        Net_Pay.Caption = ""
        City_tax.Caption = ""
        State_tax.Caption = ""
        Federal_tax.Caption = ""
        Hours.Text = ""
        Hours.SetFocus
     Else
        'keyascii 13 is enter
        If KeyAscii = 13 Then Call Calculate_Click
        'keyascii 45 is "."
        If KeyAscii > 45 Then
            'keyascii 47 we don't want
            If KeyAscii = 47 Then KeyAscii = 0
            'keyascii 48-57 is 0-9
            If KeyAscii > 57 Then KeyAscii = 0
        Else
            'keyascii 8 is backspace
            If KeyAscii <> 8 Then KeyAscii = 0
        End If
    End If
End Sub
Private Sub Rate_of_pay_Change()
    Rate_of_pay.Caption = Format$(Rate_of_pay.Caption, "Currency")
End Sub
Private Sub lblWage_Change()
lblWage.Caption = Format$(lblWage.Caption, "currency")
End Sub
Private Sub Return_Click()
    Unload Payroll_Calculator
End Sub
Private Sub Write_checks_Click()
If Hours.Text = "" Then Exit Sub
With DataLink.rsEmployee
ytd = .Fields("YTDEarnings")
ytd = ytd + Gross_Pay.Caption
SQL = "UPDATE Employee SET YTDEarnings = " & ytd & " WHERE (UID = " & lblUID.Caption & " )"
DataLink.EmpDB.Execute SQL
End With
Stack.Push lblName.Caption
Stack.Push Net_Pay.Caption
Stack.Push ytd
Stack.Push Label17.Caption
Call GUI.Load_Form(Payroll_Checks, Me)
End Sub
