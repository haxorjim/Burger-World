Attribute VB_Name = "GUI"
Public RanDidYouKnow As Boolean       'Fact Of Day Variable
Public LastUser As String           'Saves last user logged in
Public IsManager As Boolean
Public OnManagerPanel As Boolean
Public Sub Activate_Icons(TrueFalse As Boolean)
    'Controls whether or not desktop icons are available to a employee/manager
    For Index = 0 To 2
        FrmDesktop.Icons(Index).Enabled = TrueFalse
    Next
    If IsManager = True And Index = 3 Then
        FrmDesktop.Icons(3).Enabled = TrueFalse
    End If
    Index = 4
    FrmDesktop.Icons(Index).Enabled = TrueFalse
End Sub
Public Sub Visible_Icons(TrueFalse As Boolean)
    'Controls whether a employee/manager can see the desktop icons
    For Index = 0 To 4
        FrmDesktop.Icons(Index).Visible = TrueFalse
    Next Index
End Sub
Public Sub Center_Form(Form)
    'centers any window on the screen
    Form.Left = (Screen.Width - Form.Width) / 2
    Form.Top = (Screen.Height - Form.Height) / 2
End Sub
Public Sub Fake_Transparency(Form)
    'creates the illusion of a see through window in any window
    Set FT_Background = Form.Controls.Add("VB.Image", "FT_Background")
    FT_Background.Stretch = True
    FT_Background.Picture = LoadPicture(Transparency_Image)
    FT_Background.Width = Screen.Width
    FT_Background.Height = Screen.Height
    FT_Background.Left = 0 - Form.Left
    FT_Background.Top = 0 - Form.Top
    FT_Background.Visible = True
End Sub
Public Sub Remove_Fake_Transparency(Form)
    'removes the fake transparency (used while refreshing a window)
    Form.Controls.Remove ("FT_Background")
End Sub
Public Sub Fit_Background(Form)
    'stretches and positions an image to the full size of a form
    Form.Background.Left = 0
    Form.Background.Top = 0
    Form.Background.Height = Form.Height
    Form.Background.Width = Form.Width
    Form.Background.ZOrder 1
End Sub
Public Sub Full_Screen(Form)
    'makes a form full screen
    Form.Left = 0
    Form.Top = 0
    Form.Width = Screen.Width
    Form.Height = Screen.Height
End Sub
Public Sub Load_Form(Form, Owner)
    'Adds GUI Enhancments to a form before actually showing it
    Load Form
    Call GUI.Center_Form(Form)
    Call GUI.Fake_Transparency(Form)
    Call GUI.Window_Border(Form)
    Form.Show 1, Owner
End Sub
Public Sub Window_Border(Form)
    'Dynamic Window Borders, pulled from Resource File
    Set WB_Horizontal_Bar_Top = Form.Controls.Add("VB.Image", "WB_Horizontal_Bar_Top")
    Set WB_Vertical_Bar_Left = Form.Controls.Add("VB.Image", "WB_Vertical_Bar_Left")
    Set WB_Horizontal_Bar_Bottom = Form.Controls.Add("VB.Image", "WB_Horizontal_Bar_Bottom")
    Set WB_Vertical_Bar_Right = Form.Controls.Add("VB.Image", "WB_Vertical_Bar_Right")
    Set WB_Bottom_Corner = Form.Controls.Add("VB.Image", "WB_Bottom_Corner")
    WB_Horizontal_Bar_Top.Picture = LoadResPicture("HBT", vbResBitmap)
    WB_Horizontal_Bar_Top.Left = Form.Width - WB_Horizontal_Bar_Top.Width
    WB_Horizontal_Bar_Top.Top = 0
    WB_Horizontal_Bar_Top.ZOrder 0
    WB_Horizontal_Bar_Top.Visible = True
    WB_Vertical_Bar_Left.Picture = LoadResPicture("VBL", vbResBitmap)
    WB_Vertical_Bar_Left.Left = 0
    WB_Vertical_Bar_Left.Top = 0
    WB_Vertical_Bar_Left.ZOrder 0
    WB_Vertical_Bar_Left.Visible = True
    WB_Horizontal_Bar_Bottom.Picture = LoadResPicture("HBB", vbResBitmap)
    WB_Horizontal_Bar_Bottom.Left = 0
    WB_Horizontal_Bar_Bottom.Top = Form.Height - WB_Horizontal_Bar_Bottom.Height
    WB_Horizontal_Bar_Bottom.ZOrder 0
    WB_Horizontal_Bar_Bottom.Visible = True
    WB_Vertical_Bar_Right.Picture = LoadResPicture("VBR", vbResBitmap)
    WB_Vertical_Bar_Right.Left = Form.Width - WB_Vertical_Bar_Right.Width
    WB_Vertical_Bar_Right.Top = WB_Horizontal_Bar_Top.Height
    WB_Vertical_Bar_Right.ZOrder 0
    WB_Vertical_Bar_Right.Visible = True
    WB_Bottom_Corner.Picture = LoadResPicture("BC", vbResBitmap)
    WB_Bottom_Corner.Left = Form.Width - WB_Bottom_Corner.Width
    WB_Bottom_Corner.Top = Form.Height - WB_Bottom_Corner.Height
    WB_Bottom_Corner.ZOrder 0
    WB_Bottom_Corner.Visible = True
End Sub
Public Sub Remove_Window_Border(Form)
    'removes the window borders (used while refreshing a window)
    Form.Controls.Remove ("WB_Horizontal_Bar_Top")
    Form.Controls.Remove ("WB_Vertical_Bar_Left")
    Form.Controls.Remove ("WB_Horizontal_Bar_Bottom")
    Form.Controls.Remove ("WB_Vertical_Bar_Right")
    Form.Controls.Remove ("WB_Bottom_Corner")
End Sub
'Public Sub LoadDataIntoFile(DataName As Integer, FileName As String)
'    Dim myArray() As Byte
'    Dim myFile As Long
'    If Dir(FileName) = "" Then
'        myArray = LoadResData(DataName, "CUSTOM")
'        myFile = FreeFile
'        Open FileName For Binary Access Write As #myFile
'        Put #myFile, , myArray
'        Close #myFile
'    End If
'End Sub
Public Sub Place_Icons(Form, Icon_Count As Integer, Icon_Spacing As Integer, Icon_Position)
    'Very Sloppy routine for placing icons on submenus
    If Icon_Count <> 6 Then
        For Index = 0 To Icon_Count - 1
            Form.Icons(Index).Width = FrmDesktop.Default_Icon.Width
            Form.Icons(Index).Height = FrmDesktop.Default_Icon.Height
            Form.Icons(Index).Top = (Form.Height * Icon_Position)
        Next
        Length = (Icon_Count * Form.Icons(0).Width) + ((Icon_Count - 1) * Icon_Spacing)
        Form.Icons(0).Left = (Form.Width - Length) / 2
        For Index = 1 To Icon_Count - 1
            Form.Icons(Index).Left = Form.Icons(Index - 1).Left + Form.Icons(Index - 1).Width + Icon_Spacing
        Next
    Else
        For Index = 0 To 2
            Form.Icons(Index).Width = FrmDesktop.Default_Icon.Width
            Form.Icons(Index).Height = FrmDesktop.Default_Icon.Height
            Form.Icons(Index).Top = (Form.Height * Icon_Position)
        Next
        For Index = 3 To 5
            Form.Icons(Index).Width = FrmDesktop.Default_Icon.Width
            Form.Icons(Index).Height = FrmDesktop.Default_Icon.Height
            'Form.Icons(Index).Top = (Form.Height * (Icon_Position * 2.25))
            Form.Icons(Index).Top = Form.Icons(0).Top + Form.Icons(0).Height + Icon_Spacing
        Next
        Length = (3 * Form.Icons(0).Width) + (4 * Icon_Spacing)
        Form.Icons(0).Left = (Form.Width - Length) / 2
        Form.Icons(3).Left = (Form.Width - Length) / 2
        For Index = 1 To 2
            Form.Icons(Index).Left = Form.Icons(Index - 1).Left + Form.Icons(Index - 1).Width + Icon_Spacing
        Next
        For Index = 4 To 5
            Form.Icons(Index).Left = Form.Icons(Index - 1).Left + Form.Icons(Index - 1).Width + Icon_Spacing
        Next
    End If
End Sub
