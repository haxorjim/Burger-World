Attribute VB_Name = "Startup"
Public Fact_File As String
Public Schedule_File As String
Public Logo_Image As String
Public Desktop_Image As String
Public Transparency_Image As String
Public Splash_Image As String
Private Sub main()
    Call Initialize_Paths
    FrmSplash.Show
End Sub
Private Sub Initialize_Paths()
    ChDir App.Path
    Fact_File = "Config\DidYouKnow"
    Schedule_File = "Config\Schedule"
    'BurgerWorld
    'Logo_Image = "Graphics\BurgerWorld\BWLogo.gif"
    'Desktop_Image = "Graphics\BurgerWorld\BWDesktop.jpg"
    'Transparency_Image = "Graphics\BurgerWorld\BWTransparency.jpg"
    'Splash_Image = "Graphics\BurgerWorld\BWSplash.jpg"
    
    'BurgerKing
    'Logo_Image = "Graphics\BurgerKing\BKLogo.gif"
    'Desktop_Image = "Graphics\BurgerKing\BKDesktop.jpg"
    'Transparency_Image = "Graphics\BurgerKing\BKTransparency.jpg"
    'Splash_Image = "Graphics\BurgerKing\BKSplash.jpg"
    
    'McDonalds
    Logo_Image = "Graphics\McDonalds\McdsLogo.gif"
    Desktop_Image = "Graphics\McDonalds\McdsDesktop.jpg"
    Transparency_Image = "Graphics\McDonalds\McdsTransparency.jpg"
    Splash_Image = "Graphics\McDonalds\McdsSplash.jpg"
End Sub
