VERSION 5.00
Begin VB.Form frmUsrDBINT 
   BorderStyle     =   0  'None
   ClientHeight    =   6225
   ClientLeft      =   2340
   ClientTop       =   1440
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkMarital 
      DownPicture     =   "frmUsrDBINT.frx":0000
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3216
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1728
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.TextBox txtDep 
      Height          =   288
      Left            =   6240
      TabIndex        =   34
      Tag             =   "T"
      Top             =   1812
      Width           =   375
   End
   Begin VB.CheckBox chkDisabled 
      DownPicture     =   "frmUsrDBINT.frx":07AE
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   132
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2268
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CheckBox chkAdmin 
      DownPicture     =   "frmUsrDBINT.frx":0F5C
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   132
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   1692
      UseMaskColor    =   -1  'True
      Width           =   450
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   375
      Left            =   3540
      TabIndex        =   29
      Top             =   5664
      Width           =   1215
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update . . ."
      Height          =   375
      Left            =   2100
      TabIndex        =   28
      Top             =   5664
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add User. . ."
      Height          =   375
      Left            =   2100
      TabIndex        =   16
      Top             =   5664
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdFirst 
      Height          =   300
      Left            =   132
      Picture         =   "frmUsrDBINT.frx":170A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdPrevious 
      Height          =   300
      Left            =   480
      Picture         =   "frmUsrDBINT.frx":1A4C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdNext 
      Height          =   300
      Left            =   5952
      Picture         =   "frmUsrDBINT.frx":1D8E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdLast 
      Height          =   300
      Left            =   6300
      Picture         =   "frmUsrDBINT.frx":20D0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete User. . ."
      Height          =   375
      Left            =   2100
      TabIndex        =   15
      Top             =   5664
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtWage 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   285
      Left            =   5280
      TabIndex        =   9
      Tag             =   "T"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   132
      MultiLine       =   -1  'True
      TabIndex        =   10
      Tag             =   "T"
      Top             =   3144
      Width           =   6615
   End
   Begin VB.TextBox txtPW 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   5640
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   8
      Tag             =   "T"
      Top             =   810
      Width           =   972
   End
   Begin VB.TextBox txtUser 
      Height          =   288
      Left            =   5640
      MaxLength       =   10
      TabIndex        =   7
      Tag             =   "T"
      Top             =   480
      Width           =   972
   End
   Begin VB.TextBox txtTitle 
      Height          =   288
      Left            =   600
      TabIndex        =   6
      Tag             =   "T"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.TextBox txtZIP 
      Height          =   288
      Left            =   3600
      TabIndex        =   5
      Tag             =   "T"
      Top             =   840
      Width           =   972
   End
   Begin VB.TextBox txtState 
      Height          =   288
      Left            =   2760
      TabIndex        =   4
      Tag             =   "T"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtCity 
      Height          =   288
      Left            =   480
      TabIndex        =   3
      Tag             =   "T"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtAddress 
      Height          =   288
      Left            =   960
      TabIndex        =   2
      Tag             =   "T"
      Top             =   480
      Width           =   3615
   End
   Begin VB.TextBox txtName 
      Height          =   288
      Left            =   720
      TabIndex        =   1
      Tag             =   "T"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Year to Date Earnings"
      Height          =   195
      Left            =   3600
      TabIndex        =   39
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label txtYTD 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5370
      TabIndex        =   38
      Top             =   2370
      Width           =   1260
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Marital Status"
      Height          =   192
      Left            =   3816
      TabIndex        =   37
      Top             =   1848
      Width           =   960
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dependents"
      Height          =   192
      Left            =   5208
      TabIndex        =   36
      Top             =   1848
      Width           =   888
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Account Disabled"
      Height          =   192
      Left            =   696
      TabIndex        =   33
      Top             =   2400
      Width           =   1272
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Administrator Privledges"
      Height          =   192
      Left            =   696
      TabIndex        =   31
      Top             =   1824
      Width           =   1752
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Hourly Wage"
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label lblUID 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5640
      TabIndex        =   26
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "UID"
      Height          =   255
      Left            =   5160
      TabIndex        =   25
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      Height          =   252
      Left            =   168
      TabIndex        =   24
      Top             =   2892
      Width           =   852
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   6720
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1200
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "ZIP"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      Height          =   255
      Left            =   2280
      TabIndex        =   21
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   255
      Left            =   4800
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      Height          =   255
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmUsrDBINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo AddProb
    PW = Encrypt(TxtPW.Text)
    Key = Stack.Pop()
    'debug
    If txtComment.Text = "" Then txtComment.Text = " "
    If chkMarital.Value = 1 Then Mstat = "M" Else Mstat = "S"
    SQL = "INSERT INTO Employee (Address, AdminFlag, City, Comments, DisabledFlag, `Key`, Name, `Password`, State, Title, Username, Wage, ZIP,MaritalStatus,Dependents) VALUES ('" _
            & txtAddress.Text & "', '" _
            & chkAdmin.Value & "', '" _
            & txtCity.Text & "', '" _
            & txtComment.Text & "', '" _
            & chkDisabled.Value & "', '" _
            & Key & "', '" _
            & txtName.Text & "', '" _
            & PW & "', '" _
            & txtState.Text & "', '" _
            & txtTitle.Text & "', '" _
            & TxtUser.Text & "', '" _
            & txtWage.Text & "', '" _
            & txtZIP.Text & "', '" _
            & Mstat & "', '" _
            & txtDep.Text & "')"
    DataLink.EmpDB.Execute SQL
    'MsgBox "Added.", , "Completed"
    Stack.Push "Employee Added..."
    GUI.Load_Form frmMsgBox, FrmDesktop
Exit Sub
AddProb:
MsgBox Err.Description
End Sub
Private Sub cmdDel_Click()
'On Error GoTo DelProb
Stack.Push "Are you sure you wish to delete this record?"
GUI.Load_Form frmAreYouSure, FrmDesktop
If Stack.Pop = "YES" Then
'If MsgBox("Are you sure you wish to delete this record?", vbYesNo + vbExclamation, "Confirmation") = vbYes Then
    UIDToDel = lblUID.Caption
    SQL = "DELETE FROM Employee WHERE UID = " & UIDToDel
    If DataLink.rsEmployee.Bookmark = DataLink.rsEmployee.RecordCount Then DataLink.rsEmployee.MovePrevious
    DataLink.EmpDB.Execute SQL
    DataLink.rsEmployee.Requery
    'MsgBox "Deleted.", , "Completed"
    Stack.Push "Employee Deleted..."
    GUI.Load_Form frmMsgBox, FrmDesktop
    DataLink.rsEmployee.Requery
    Call cmdNext_Click
End If
Exit Sub
DelProb:
MsgBox Err.Description
End Sub
Private Sub cmdDone_Click()
    Unload Me
End Sub
Private Sub cmdFirst_Click()
  'On Error GoTo GoFirstError
  DataLink.rsEmployee.MoveFirst
  Call Display_Data
  Exit Sub
GoFirstError:
  MsgBox Err.Description
End Sub
Private Sub Display_Data()
  lblUID.Caption = DataLink.rsEmployee.Fields("UID")
  txtName.Text = DataLink.rsEmployee.Fields("Name")
  TxtUser.Text = DataLink.rsEmployee.Fields("Username")
  TxtPW.Text = Security.Decrypt(DataLink.rsEmployee.Fields("Password"), DataLink.rsEmployee.Fields("Key"))
  txtAddress.Text = DataLink.rsEmployee.Fields("Address")
  txtCity.Text = DataLink.rsEmployee.Fields("City")
  txtState.Text = DataLink.rsEmployee.Fields("State")
  txtZIP.Text = DataLink.rsEmployee.Fields("ZIP")
  txtTitle.Text = DataLink.rsEmployee.Fields("Title")
  txtWage.Text = FormatCurrency(DataLink.rsEmployee.Fields("wage"))
  If DataLink.rsEmployee.Fields("AdminFlag") = True Then chkAdmin.Value = 1 Else chkAdmin.Value = 0
  If DataLink.rsEmployee.Fields("DisabledFlag") = True Then chkDisabled.Value = 1 Else chkDisabled.Value = 0
  txtComment.Text = DataLink.rsEmployee.Fields("Comments")
  If UCase(DataLink.rsEmployee.Fields("MaritalStatus")) = "M" Then chkMarital.Value = 1 Else chkMarital.Value = 0
  txtDep.Text = DataLink.rsEmployee.Fields("Dependents")
  txtYTD.Caption = FormatCurrency(DataLink.rsEmployee.Fields("YTDEarnings"))
End Sub
Private Sub cmdLast_Click()
On Error GoTo GoLastError
  DataLink.rsEmployee.MoveLast
  Call Display_Data
  Exit Sub
GoLastError:
  MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
On Error GoTo GoNextError
With DataLink.rsEmployee
  If Not .EOF Then .MoveNext
  If .EOF And .RecordCount > 0 Then
    Beep
    .MoveLast
  End If
End With
Call Display_Data
  Exit Sub
GoNextError:
  MsgBox Err.Description
End Sub
Private Sub cmdPrevious_Click()
'On Error GoTo GoPrevError
With DataLink.rsEmployee
  If Not .BOF Then .MovePrevious
  If .BOF And .RecordCount > 0 Then
    Beep
    .MoveFirst
  End If
End With
Call Display_Data
Exit Sub
GoPrevError:
  MsgBox Err.Description
End Sub
Private Sub cmdUpdate_Click()
If chkMarital.Value = 1 Then Mstat = "M" Else Mstat = "S"
SQL = "UPDATE Employee SET Address = " & "'" & txtAddress.Text & "'" & ",AdminFlag = " & "'" & chkAdmin.Value & "'" & ",City = " & "'" & txtCity.Text & "'" & ",Comments = " & "'" & txtComment.Text & "'" & ",DisabledFlag = " & "'" & chkDisabled.Value & "'" & ",`Password` = " & "'" & Security.Encrypt(TxtPW.Text) & "'" & ",`Key` =  " & "'" & Stack.Pop() & "'" & ",Name = " & "'" & txtName.Text & "'" & ",State = " & "'" & txtState.Text & "'" & ",Title = " & "'" & txtTitle.Text & "'" & ",Username = " & "'" & TxtUser.Text & "'" & ",Wage = " & "'" & txtWage.Text & "'" & ",ZIP = " & "'" & txtZIP.Text & "'" & ",MaritalStatus = " & "'" & Mstat & "'" & ",Dependents = " & "'" & txtDep.Text & "'" & "WHERE (UID = " & lblUID.Caption & ")"
DataLink.EmpDB.Execute SQL
DataLink.rsEmployee.Requery
'MsgBox "Updated.", , "Completed."
Stack.Push "Employee Record Updated..."
GUI.Load_Form frmMsgBox, FrmDesktop
End Sub
Private Sub Form_Load()
'Load Database
DataLink.rsEmployee.Open
'on Error Resume Next
'Global Variable - UsrDBCtrl
'PHASED OUT & Replaced by myt nifty stack 2/4/01
'numbers remain the same
' 1 - New Employee
' 2 - Change Employee
' 3 - Delete Employee
' 4 - List Employees
' 5 - Search for Employee
usrdbctrl = Stack.Pop()
If usrdbctrl = 1 Then
    Me.Caption = "Add a New User"
    cmdAdd.Visible = True
ElseIf usrdbctrl = 2 Then
    Me.Caption = "Modify User Entry"
    cmdFirst.Visible = True
    cmdLast.Visible = True
    cmdNext.Visible = True
    cmdPrevious.Visible = True
    cmdUpdate.Visible = True
    Call cmdFirst_Click
ElseIf usrdbctrl = 3 Then
    Me.Caption = "Delete User"
    cmdDel.Visible = True
    cmdFirst.Visible = True
    cmdLast.Visible = True
    cmdNext.Visible = True
    cmdPrevious.Visible = True
    Call cmdFirst_Click
    For Each ctrl In Controls
        If ctrl.Tag = "T" Then
            ctrl.Locked = True
        End If
    
        If ctrl.Tag = "C" Then
            ctrl.Enabled = False
        End If
    Next ctrl
    'disabled check buttons
    chkDisabled.Enabled = False
    chkAdmin.Enabled = False
    chkMarital.Enabled = False
ElseIf usrdbctrl = 4 Then
    Me.Caption = "Employee List -- Read Only"
    cmdFirst.Visible = True
    cmdLast.Visible = True
    cmdNext.Visible = True
    cmdPrevious.Visible = True
    'disabled check buttons
    chkDisabled.Enabled = False
    chkAdmin.Enabled = False
    chkMarital.Enabled = False
    Call cmdFirst_Click
    For Each ctrl In Controls
        If ctrl.Tag = "T" Then
            ctrl.BorderStyle = 1
            ctrl.BackColor = vbMenuBar
            ctrl.Locked = True
        End If
        If ctrl.Tag = "C" Then
            ctrl.Enabled = False
        End If
    Next ctrl
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
DataLink.rsEmployee.Close
End Sub
Private Sub txtComment_Change()
    If Right(txtComment.Text, 1) = "'" Then
        txtComment.Text = Left(txtComment.Text, Len(txtComment.Text) - 1)
        txtComment.SelStart = Len(txtComment.Text)
    End If
End Sub
