Attribute VB_Name = "Module1"
Option Explicit

Public Function ReturnFontName(pFont As eFonts) As String
    Select Case pFont
        Case eFonts.NormalPixel_5x7
            ReturnFontName = "Normal Pixel 5x7"
        Case eFonts.OutLand9x7
            ReturnFontName = "OutLand 9x7"
        Case eFonts.DPComic10x14
            ReturnFontName = "DPComic 10x14"
    End Select
End Function
Public Function ReturnColorName(pColor As eMatrixColor) As String
    Select Case pColor
        Case eMatrixColor.Black
            ReturnColorName = "Black"
        Case eMatrixColor.Blue
            ReturnColorName = "Blue"
        Case eMatrixColor.GreenDark
            ReturnColorName = "GreenDark"
        Case eMatrixColor.GreenLight
            ReturnColorName = "GreenLight"
        Case eMatrixColor.NavyDark
            ReturnColorName = "NavyDark"
        Case eMatrixColor.NavyLight
            ReturnColorName = "NavyLight"
        Case eMatrixColor.OliveDark
            ReturnColorName = "OliveDark"
        Case eMatrixColor.OliveLight
            ReturnColorName = "OliveLight"
        Case eMatrixColor.OliveSuperDark
            ReturnColorName = "OliveSuperDark"
        Case eMatrixColor.OrangeDark
            ReturnColorName = "OrangeDark"
        Case eMatrixColor.OrangeLight
            ReturnColorName = "OrangeLight"
        Case eMatrixColor.Red
            ReturnColorName = "Red"
        Case eMatrixColor.White
            ReturnColorName = "White"
        Case eMatrixColor.Yellow
            ReturnColorName = "Yellow"
    End Select
End Function

Public Function ReturnColor(pColor As eMatrixColor) As Long
    Select Case pColor
        Case eMatrixColor.Black
            ReturnColor = vbBlack
        Case eMatrixColor.OliveDark
            ReturnColor = RGB(148, 157, 64)
        Case eMatrixColor.OliveSuperDark
            ReturnColor = RGB(84, 91, 34)
        Case eMatrixColor.OliveLight
            ReturnColor = RGB(164, 176, 70)
        Case eMatrixColor.NavyDark
            ReturnColor = RGB(87, 146, 210)
        Case eMatrixColor.NavyLight
            ReturnColor = RGB(101, 167, 216)
        Case eMatrixColor.OrangeDark
            ReturnColor = RGB(213, 128, 0)
        Case eMatrixColor.OrangeLight
            ReturnColor = RGB(255, 157, 38)
        Case eMatrixColor.GreenDark
            ReturnColor = RGB(0, 83, 0)
        Case eMatrixColor.GreenLight
            ReturnColor = vbGreen
        Case eMatrixColor.White
            ReturnColor = vbWhite
        Case eMatrixColor.Yellow
            ReturnColor = vbYellow
        Case eMatrixColor.Red
            ReturnColor = vbRed
        Case eMatrixColor.Blue
            ReturnColor = vbBlue
    End Select
End Function
