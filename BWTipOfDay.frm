VERSION 5.00
Begin VB.Form FrmDidYouKnow 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Tip of the Day"
   ClientHeight    =   4080
   ClientLeft      =   2565
   ClientTop       =   2055
   ClientWidth     =   6795
   Icon            =   "BWTipOfDay.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   735
      Left            =   4845
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdNextFact 
      Caption         =   "&Next Tip"
      Height          =   735
      Left            =   4815
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3555
      Left            =   225
      ScaleHeight     =   3495
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   255
      Width           =   3975
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Did you know?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   855
         TabIndex        =   4
         Top             =   210
         Width           =   3285
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "BWTipOfDay.frx":030A
         Top             =   150
         Width           =   480
      End
      Begin VB.Label lblFactText 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Height          =   2610
         Left            =   105
         TabIndex        =   1
         Top             =   960
         Width           =   3720
      End
   End
   Begin VB.Image Logo 
      Height          =   2250
      Left            =   4425
      Top             =   240
      Width           =   2250
   End
End
Attribute VB_Name = "FrmDidYouKnow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Facts As New Collection
Dim CurrentFact As Long
Private Sub Form_Load()
    Randomize
    ' Load a random startup fact from fact file
    If LoadFacts(Fact_File) = False Then
        lblFactText.Caption = "That the " & Fact_File & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & Fact_File & " using NotePad with 1 Fact per line. " & _
           "Then place it in the same directory as the application. "
    End If
    Logo.Picture = LoadPicture(Logo_Image)
End Sub
Private Sub DoNextFact()
    'Increment fact counter and display new fact,
    'if out of facts loop and display first fact again
    CurrentFact = CurrentFact + 1
    If Facts.Count < CurrentFact Then
        CurrentFact = 1
    End If
    FrmDidYouKnow.DisplayCurrentFact
End Sub
Function LoadFacts(sFile As String) As Boolean
    Dim NextFact As String
    Dim InFile As Integer
    InFile = FreeFile
    If sFile = "" Then
        LoadFacts = False
        Exit Function
    End If
    If Dir(sFile) = "" Then
        LoadFacts = False
        Exit Function
    End If
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextFact
        Facts.Add NextFact
    Wend
    Close InFile
    DoNextFact
    LoadFacts = True
End Function
Public Sub DisplayCurrentFact()
    If Facts.Count > 0 Then
        lblFactText.Caption = Facts.Item(CurrentFact)
    End If
End Sub
Private Sub cmdNextFact_Click()
    'Load the next Fact
    DoNextFact
End Sub
Private Sub cmdOK_Click()
    'Close the Window
    Unload FrmDidYouKnow
End Sub
