VERSION 5.00
Object = "{6F2FE922-F57D-4D42-8468-B5B824A2D50E}#2.0#0"; "DisplayMatOcx.ocx"
Begin VB.Form frmTest 
   Caption         =   "Display Matrix Test Form"
   ClientHeight    =   3990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix7 
      Height          =   315
      Left            =   3780
      TabIndex        =   23
      Top             =   1620
      Width           =   8055
      _extentx        =   14208
      _extenty        =   556
      mtxforecolor    =   0
      mtxbackcolor    =   4
      mtxbackforecolor=   3
      mtxfont         =   0
      mtxdotheight    =   2
      mtxdotwidth     =   2
      mtxcaption      =   "...and the best: it's FREE!!! :-)"
      mtxautoajustheight=   -1  'True
   End
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix5 
      Height          =   630
      Left            =   3780
      TabIndex        =   22
      Top             =   960
      Width           =   8055
      _extentx        =   14208
      _extenty        =   1111
      mtxforecolor    =   2
      mtxbackcolor    =   10
      mtxbackforecolor=   5
      mtxfont         =   2
      mtxdotheight    =   2
      mtxdotwidth     =   1
      mtxcaption      =   "More Than 500 Combinations"
      mtxautoajustheight=   -1  'True
   End
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix4 
      Height          =   315
      Left            =   3780
      TabIndex        =   21
      Top             =   600
      Width           =   8055
      _extentx        =   14208
      _extenty        =   556
      mtxforecolor    =   6
      mtxbackcolor    =   11
      mtxbackforecolor=   7
      mtxfont         =   1
      mtxdotheight    =   2
      mtxdotwidth     =   2
      mtxcaption      =   "Three Display Fonts"
      mtxautoajustheight=   -1  'True
   End
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix3 
      Height          =   210
      Left            =   3780
      TabIndex        =   20
      Top             =   360
      Width           =   6135
      _extentx        =   10821
      _extenty        =   370
      mtxforecolor    =   13
      mtxbackcolor    =   5
      mtxbackforecolor=   0
      mtxfont         =   0
      mtxdotheight    =   1
      mtxdotwidth     =   1
      mtxcaption      =   "- Ideal for Instrumentation applications"
      mtxautoajustheight=   -1  'True
   End
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix2 
      Height          =   315
      Left            =   3780
      TabIndex        =   19
      Top             =   0
      Width           =   5715
      _extentx        =   10081
      _extenty        =   556
      mtxforecolor    =   9
      mtxbackcolor    =   8
      mtxbackforecolor=   0
      mtxfont         =   0
      mtxdotheight    =   2
      mtxdotwidth     =   2
      mtxcaption      =   "- Full Customizable OCX"
   End
   Begin VB.Frame Frame1 
      Caption         =   "DisplayMatrix Properties"
      Height          =   3435
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3615
      Begin VB.CheckBox chkScroll 
         Caption         =   "Scrollable Text (no OCX)"
         Height          =   195
         Left            =   1020
         TabIndex        =   18
         Top             =   3120
         Width           =   2235
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   1020
         TabIndex        =   17
         Top             =   2760
         Width           =   2415
      End
      Begin VB.ComboBox cboSize 
         Height          =   315
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox cboForeColor 
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1260
         Width           =   2475
      End
      Begin VB.ComboBox cboBackColor 
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   2475
      End
      Begin VB.ComboBox cboDisplayBackColor 
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   2475
      End
      Begin VB.ComboBox cboSizeY 
         Height          =   315
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chkAutoHeight 
         Caption         =   "AutoHeight"
         Height          =   195
         Left            =   2400
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox cboFont 
         Height          =   315
         Left            =   540
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "Caption:"
         Height          =   255
         Left            =   60
         TabIndex        =   16
         Top             =   2820
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "Font:"
         Height          =   255
         Left            =   60
         TabIndex        =   15
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Size:"
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   900
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "ForeColor:"
         Height          =   255
         Left            =   60
         TabIndex        =   13
         Top             =   1380
         Width           =   795
      End
      Begin VB.Label Label4 
         Caption         =   "BackColor:"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Display BackColor:"
         Height          =   435
         Left            =   60
         TabIndex        =   11
         Top             =   2220
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Width"
         Height          =   195
         Left            =   600
         TabIndex        =   10
         Top             =   660
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "Height"
         Height          =   195
         Left            =   1560
         TabIndex        =   9
         Top             =   660
         Width           =   675
      End
   End
   Begin DisplayMatOcx.DisplayMatrix DisplayMatrix1 
      Height          =   435
      Left            =   60
      TabIndex        =   0
      Top             =   3540
      Width           =   11775
      _extentx        =   20770
      _extenty        =   767
      mtxforecolor    =   0
      mtxbackcolor    =   1
      mtxbackforecolor=   3
      mtxfont         =   0
      mtxdotheight    =   3
      mtxdotwidth     =   3
      mtxcaption      =   "MatrixDisplay 1.0"
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Left            =   3660
      Top             =   3000
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ScrollText As String

Private Sub cboBackColor_Click()
    cboBackColor.BackColor = ReturnColor(cboBackColor.ItemData(cboBackColor.ListIndex))
    DisplayMatrix1.MtxBackColor = cboBackColor.ItemData(cboBackColor.ListIndex)
End Sub

Private Sub cboDisplayBackColor_Click()
    cboDisplayBackColor.BackColor = ReturnColor(cboDisplayBackColor.ItemData(cboDisplayBackColor.ListIndex))
    DisplayMatrix1.MtxDisplayBackColor = cboDisplayBackColor.ItemData(cboDisplayBackColor.ListIndex)
End Sub

Private Sub cboFont_Click()
    DisplayMatrix1.MtxFont = cboFont.ItemData(cboFont.ListIndex)
    Me.Height = DisplayMatrix1.Top + DisplayMatrix1.Height + 512
End Sub

Private Sub cboForeColor_Click()
    cboForeColor.BackColor = ReturnColor(cboForeColor.ItemData(cboForeColor.ListIndex))
    DisplayMatrix1.MtxForeColor = cboForeColor.ItemData(cboForeColor.ListIndex)
End Sub


Private Sub cboSize_Click()
    DisplayMatrix1.MtxDotWidth = cboSize.ItemData(cboSize.ListIndex)
    Me.Height = DisplayMatrix1.Top + DisplayMatrix1.Height + 512
End Sub


Private Sub cboSizeY_Click()
    DisplayMatrix1.MtxDotHeight = cboSizeY.ItemData(cboSizeY.ListIndex)
    Me.Height = DisplayMatrix1.Top + DisplayMatrix1.Height + 512
End Sub

Private Sub chkAutoHeight_Click()
    If chkAutoHeight.Value = vbChecked Then
        DisplayMatrix1.MtxAutoAjustHeight = True
        Me.Height = DisplayMatrix1.Top + DisplayMatrix1.Height + 512
    Else
        DisplayMatrix1.MtxAutoAjustHeight = False
    End If
End Sub

Private Sub chkScroll_Click()
    If chkScroll.Value = vbChecked Then
        ScrollText = DisplayMatrix1.MtxCaption
        tmrScroll.Interval = 150
        tmrScroll.Enabled = True
    Else
        tmrScroll.Enabled = False
    End If
End Sub



Private Sub Form_Load()
Dim iFonts As eFonts
Dim iSize As eDotSize
Dim iColor As eMatrixColor
Dim color As Long
    
    cboFont.Clear
    cboSize.Clear
    cboSizeY.Clear
    cboForeColor.Clear
    cboBackColor.Clear
    cboDisplayBackColor.Clear
    For iFonts = 0 To eFonts.DPComic10x14
        cboFont.AddItem ReturnFontName(iFonts)
        cboFont.ItemData(cboFont.NewIndex) = iFonts
    Next iFonts
    
    cboFont.ListIndex = DisplayMatrix1.MtxFont
    
    For iSize = eDotSize.Size1Pix To eDotSize.Size10Pix
        cboSize.AddItem iSize & "Pix"
        cboSize.ItemData(cboSize.NewIndex) = iSize
        cboSizeY.AddItem iSize & "Pix"
        cboSizeY.ItemData(cboSizeY.NewIndex) = iSize
    Next iSize
    
    cboSize.ListIndex = DisplayMatrix1.MtxDotWidth - 1
    cboSizeY.ListIndex = DisplayMatrix1.MtxDotHeight - 1
    
    
    For iColor = 0 To eMatrixColor.Blue
        cboForeColor.AddItem ReturnColorName(iColor)
        cboForeColor.ItemData(cboForeColor.NewIndex) = iColor
        cboBackColor.AddItem ReturnColorName(iColor)
        cboBackColor.ItemData(cboBackColor.NewIndex) = iColor
        cboDisplayBackColor.AddItem ReturnColorName(iColor)
        cboDisplayBackColor.ItemData(cboDisplayBackColor.NewIndex) = iColor
    Next iColor
    
    cboForeColor.ListIndex = DisplayMatrix1.MtxForeColor
    cboBackColor.ListIndex = DisplayMatrix1.MtxBackColor
    cboDisplayBackColor.ListIndex = DisplayMatrix1.MtxDisplayBackColor
    
    If DisplayMatrix1.MtxAutoAjustHeight = True Then
        chkAutoHeight.Value = vbChecked
    Else
        chkAutoHeight.Value = vbUnchecked
    End If
        
    txtCaption.Text = DisplayMatrix1.MtxCaption
End Sub

Private Sub tmrScroll_Timer()
    ScrollText = Mid(ScrollText, 2) & Left(ScrollText, 1)
    DisplayMatrix1.MtxCaption = ScrollText
End Sub

Private Sub txtCaption_Change()
    DisplayMatrix1.MtxCaption = txtCaption.Text
End Sub

