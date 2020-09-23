VERSION 5.00
Begin VB.UserControl DisplayMatrix 
   AutoRedraw      =   -1  'True
   ClientHeight    =   660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3450
   PropertyPages   =   "DisplayMatrix.ctx":0000
   ScaleHeight     =   44
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   230
   ToolboxBitmap   =   "DisplayMatrix.ctx":0014
   Begin VB.PictureBox PicMain 
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   217
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox PicClean 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   217
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox PicWork 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   37
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   217
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "DisplayMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim IsPaint As Boolean
Dim ReadFromBag As Boolean
Public Enum eFonts
    NormalPixel_5x7 = 0
    OutLand9x7
    DPComic10x14
End Enum

Public Enum eDotSize
    Size1Pix = 1
    Size2Pix
    Size3Pix
    Size4Pix
    Size5Pix
    Size6Pix
    Size7Pix
    Size8Pix
    Size9Pix
    Size10Pix
End Enum

Public Enum eMatrixColor
    Black = 0
    OliveDark
    OliveSuperDark
    OliveLight
    NavyDark
    NavyLight
    OrangeDark
    OrangeLight
    GreenDark
    GreenLight
    White
    Yellow
    Red
    Blue
End Enum

Dim mForeColor As eMatrixColor
Dim mBackColor As eMatrixColor
Dim mBackForeColor As eMatrixColor
Dim mFont As eFonts
'Dim mDotSize As eDotSize
Dim mCaption As String
Dim mDotWidth As eDotSize
Dim mDotHeight As eDotSize
Dim mAutoAjustHeight As Boolean
Dim SDC As Long
Dim DDC As Long

Dim TableFont As String

Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Function ReturnColor(pColor As eMatrixColor) As Long
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

Private Sub InitBackColor(ObjPicture As PictureBox, pColor As eMatrixColor)
    ObjPicture.BackColor = ReturnColor(pColor)
    DrawDots
End Sub
Public Property Get MtxAutoAjustHeight() As Boolean
Attribute MtxAutoAjustHeight.VB_Description = "Adjust automatically the height of Display Area"
    MtxAutoAjustHeight = mAutoAjustHeight
End Property
Public Property Let MtxAutoAjustHeight(ByVal Value As Boolean)
     mAutoAjustHeight = Value
    DrawDots
'    If mAutoAjustHeight = True Then
'        Select Case mFont
'            Case eFonts.NormalPixel_5x7
'                UserControl.ScaleHeight = (7 * mDotHeight) + 2
'            Case eFonts.OutLand9x7
'                UserControl.ScaleHeight = (9 * mDotHeight) + 2
'            Case eFonts.DPComic10x14
'                UserControl.ScaleHeight = (10 * mDotHeight) + 2
'                UserControl.Height = ((10 * mDotHeight) + 2) * Screen.TwipsPerPixelX
'        End Select
''        UserControl.Refresh
''        DrawDots
'    End If
    
End Property


Public Property Get MtxDotWidth() As eDotSize
Attribute MtxDotWidth.VB_Description = "Dots width ( 1 - 10 Pixels)"
    MtxDotWidth = mDotWidth
End Property

Public Property Let MtxDotWidth(ByVal Value As eDotSize)
    If Value < Size1Pix Or Value > Size10Pix Then
            Err.Raise 340, , "Invalid Property Value", "DisplayMatrix.HLP", 3001
        Exit Property
    End If
    
    mDotWidth = Value
    
    InitBackColor frmGFX.PicDots(0), mForeColor
    InitBackColor frmGFX.PicDots(1), mBackColor

End Property

Public Property Get MtxDotHeight() As eDotSize
Attribute MtxDotHeight.VB_Description = "Dots height ( 1 - 10 Pixels)"
    MtxDotHeight = mDotHeight
End Property

Public Property Let MtxDotHeight(ByVal Value As eDotSize)
    If Value < Size1Pix Or Value > Size10Pix Then
            Err.Raise 340, , "Invalid Property Value", "DisplayMatrix.HLP", 3001
        Exit Property
    End If
    
    mDotHeight = Value
    
    InitBackColor frmGFX.PicDots(0), mForeColor
    InitBackColor frmGFX.PicDots(1), mBackColor

End Property


Public Property Get MtxCaption() As String
Attribute MtxCaption.VB_Description = "Caption of Display Matrix Text"
    MtxCaption = mCaption
End Property
'Public Property Let Refresh(ByVal Value As Boolean)
'    UserControl.Refresh
'End Property


Public Property Let MtxCaption(ByVal Value As String)
    mCaption = Value
    DrawDots
End Property

Public Property Get MtxFont() As eFonts
Attribute MtxFont.VB_Description = "Select font for Display Matrix"
    MtxFont = mFont
End Property

Public Property Let MtxFont(ByVal Value As eFonts)
    If Value < NormalPixel_5x7 Or Value > DPComic10x14 Then
            Err.Raise 340, , "Invalid Property Value", "DisplayMatrix.HLP", 3001
        Exit Property
    End If
    
    mFont = Value
    InizitializeFont mFont
End Property

Public Property Get MtxForeColor() As eMatrixColor
Attribute MtxForeColor.VB_Description = "Select the Fore Color "
    MtxForeColor = mForeColor
End Property

Public Property Let MtxForeColor(ByVal Value As eMatrixColor)
    If Value < Black Or Value > Blue Then
            Err.Raise 340, , "Invalid Property Value", "DisplayMatrix.HLP", 3001
        Exit Property
    End If
    
    mForeColor = Value
    
    InitBackColor frmGFX.PicDots(0), Value

End Property

Public Property Get MtxBackColor() As eMatrixColor
Attribute MtxBackColor.VB_Description = "Back Color for Matrix Display"
    MtxBackColor = mBackColor
End Property

Public Property Let MtxBackColor(ByVal Value As eMatrixColor)
    If Value < Black Or Value > Blue Then
            Err.Raise 340, , "Invalid Property Value", "DisplayMatrix.HLP", 3001
        Exit Property
    End If
    
    mBackColor = Value
    
    InitBackColor frmGFX.PicDots(1), Value

End Property

Public Property Get MtxDisplayBackColor() As eMatrixColor
Attribute MtxDisplayBackColor.VB_Description = "Back Ground Color for Display"
    MtxDisplayBackColor = mBackForeColor
End Property

Public Property Let MtxDisplayBackColor(ByVal Value As eMatrixColor)
    If Value < Black Or Value > Blue Then
            Err.Raise 340, , "Invalid Property Value", "DisplayMatrix.HLP", 3001
        Exit Property
    End If
    
    mBackForeColor = Value
    InitBackColor PicClean, Value
    
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub



Private Sub InizitializeFont(pFont As eFonts)
    Debug.Print "InizitializeFont"
    'Select Case pFont
        'Case eFonts.NormalPixel_5x7
            frmGFX.PicFont = frmGFX.PicSprFonts(pFont)
            TableFont = InitializeStringFont(pFont)
            
    'End Select
    GetObjectAPI frmGFX.PicFont.Picture, Len(bmpBuff), bmpBuff
    
    With saBuff
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = bmpBuff.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bmpBuff.bmWidthBytes
        .pvData = bmpBuff.bmBits
    End With
    
   CopyMemory ByVal VarPtrArray(pictBuff), VarPtr(saBuff), 4
'    If mAutoAjustHeight = True Then
'        Select Case mFont
'            Case eFonts.NormalPixel_5x7
'                UserControl.Height = (7 * mDotHeight) + 2
'            Case eFonts.OutLand9x7
'                UserControl.Height = (9 * mDotHeight) + 2
'            Case eFonts.DPComic10x14
'                UserControl.Height = (10 * mDotHeight) + 2
'        End Select
'
'    Else
        DrawDots
'    End If
End Sub

Private Sub UserControl_Initialize()


    Debug.Print "initialize"
'    InizitializeFont mFont
'End Sub
''
'Private Sub UserControl_InitProperties()
    Debug.Print "initproperties"
    'If ReadFromBag = False Then
        mForeColor = Black
        mBackColor = OliveDark
        mBackForeColor = OliveLight
        mFont = NormalPixel_5x7
    
        mDotWidth = Size3Pix
        mDotHeight = Size3Pix
        mAutoAjustHeight = False
    
        PicClean.BackColor = ReturnColor(OliveLight) ' RGB(164, 176, 70)
        'UserControl.BackColor = ReturnColor(OliveLight) 'RGB(164, 176, 70)
    
        frmGFX.PicDots(0).BackColor = ReturnColor(Black) '0
        frmGFX.PicDots(1).BackColor = ReturnColor(OliveDark) 'RGB(148, 157, 64)
    
        mCaption = "MatrixDisplay 1.0" ' !" & Chr$(34) & "#$%&'()*+,-./"
    'Else
    
    'End If
    InizitializeFont mFont

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'Private Sub UserControl_Paint()
'    If IsPaint = False Then
'        IsPaint = True
'        Debug.Print "Paint"
'        UserControl.Refresh
'    End If
'End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Debug.Print "readproperties"
'    If Ambient.UserMode Then
'        mCaption = ""
'    Else
'        mCaption = "---IDE Mode---"
'    End If
    ReadFromBag = True
    mForeColor = PropBag.ReadProperty("MtxForeColor", "") 'Black)
    mBackColor = PropBag.ReadProperty("MtxBackColor", "") 'OliveDark)
    mBackForeColor = PropBag.ReadProperty("MtxBackForeColor", "") 'OliveLight)
    mFont = PropBag.ReadProperty("MtxFont", "") 'NormalPixel_5x7)
    mDotHeight = PropBag.ReadProperty("MtxDotHeight", "") 'Size3Pix)
    mDotWidth = PropBag.ReadProperty("MtxDotWidth", "") 'Size3Pix)
    mCaption = PropBag.ReadProperty("MtxCaption", "")
    mAutoAjustHeight = PropBag.ReadProperty("MtxAutoAjustHeight", False)
        PicClean.BackColor = ReturnColor(mBackForeColor) ' RGB(164, 176, 70)
        'UserControl.BackColor = ReturnColor(OliveLight) 'RGB(164, 176, 70)
    
        frmGFX.PicDots(0).BackColor = ReturnColor(mForeColor) '0
        frmGFX.PicDots(1).BackColor = ReturnColor(mBackColor) 'RGB(148, 157, 64)
    
    
    InizitializeFont mFont
End Sub

Private Sub UserControl_Resize()
    Debug.Print "resize"
    UserControl.ScaleMode = 3
    PicMain.Height = UserControl.ScaleHeight
    PicClean.Height = UserControl.ScaleHeight
    PicWork.Height = UserControl.ScaleHeight
    PicMain.Width = UserControl.ScaleWidth
    PicClean.Width = UserControl.ScaleWidth
    PicWork.Width = UserControl.ScaleWidth
    PicClean.ScaleMode = 3
    PicWork.ScaleMode = 3
    'InizitializeFont mFont
    DrawDots
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Debug.Print "writeproperties"
    Call PropBag.WriteProperty("MtxForeColor", mForeColor, "") 'Black)
    Call PropBag.WriteProperty("MtxBackColor", mBackColor, "") 'OliveDark)
    Call PropBag.WriteProperty("MtxBackForeColor", mBackForeColor, "") 'OliveLight)
    Call PropBag.WriteProperty("MtxFont", mFont, "") 'NormalPixel_5x7)
    Call PropBag.WriteProperty("MtxDotHeight", mDotHeight, "") 'mDotHeight, Size3Pix)
    Call PropBag.WriteProperty("MtxDotWidth", mDotWidth, "") 'Size3Pix)
    Call PropBag.WriteProperty("MtxCaption", mCaption, "")
    Call PropBag.WriteProperty("MtxAutoAjustHeight", mAutoAjustHeight, False)
End Sub

Private Sub Clean()
    SDC = PicClean.hDC
    DDC = PicWork.hDC
    BitBlt DDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub

Private Sub Render()
    SDC = PicWork.hDC
    DDC = UserControl.hDC
    'DDC = PicMain.hDC
    BitBlt DDC, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight, SDC, 0, 0, SRCCOPY
End Sub

Private Sub DrawDots()
Dim DCDotFore As Long
Dim DCDotBack As Long
Dim Ancho As Integer
Dim TamX As Integer
Dim TamY As Integer
Dim SprH As Integer
Dim SprW As Integer
Dim TempX As Integer
Dim TempY As Integer
Dim ActualX As Integer
Dim ActualY As Integer
Dim NumChar As Integer
Dim i As Integer
Dim jX As Integer
Dim jY As Integer
Dim QueChar As String
Dim InternalCaption As String
Dim PositionOnTable As Integer
    
    Debug.Print "DrawDots"
    Clean
    'UserControl.BackColor = ReturnColor(mBackForeColor)
    
    DCDotFore = frmGFX.PicDots(0).hDC
    DCDotBack = frmGFX.PicDots(1).hDC
    SprH = mDotHeight
    SprW = mDotWidth
    
    Select Case mFont
        Case eFonts.NormalPixel_5x7
            TamX = 5
            TamY = 7
        Case eFonts.OutLand9x7
            TamX = 9
            TamY = 7
        Case eFonts.DPComic10x14
            TamX = 10
            TamY = 14
        
    End Select
    
    Ancho = (TamX * SprW) + (TamX - 1) + SprW
    
    NumChar = (UserControl.ScaleWidth \ Ancho) + 1
    InternalCaption = mCaption
    InternalCaption = InternalCaption & Space$(NumChar)
    
    TempX = 0
    TempY = 0
    ActualX = 0
    ActualY = 0
    
    DDC = PicWork.hDC
    
'    If TableFont = "" Then
'        InizitializeFont mFont
'    End If

    
    If mAutoAjustHeight = True Then
        Select Case mFont
            Case eFonts.NormalPixel_5x7, eFonts.OutLand9x7
                UserControl.ScaleHeight = (7 * mDotHeight) + 7
                UserControl.Height = ((7 * mDotHeight) + 7) * Screen.TwipsPerPixelX
'            Case eFonts.OutLand9x7
'                UserControl.ScaleHeight = (7 * mDotHeight) + 7
'                UserControl.Height = ((7 * mDotHeight) + 7) * Screen.TwipsPerPixelX
            Case eFonts.DPComic10x14
                UserControl.ScaleHeight = (14 * mDotHeight) + 14
                UserControl.Height = ((14 * mDotHeight) + 14) * Screen.TwipsPerPixelX
        End Select
    End If
    
    For i = 1 To NumChar
        
        QueChar = Mid(InternalCaption, i, 1)
        PositionOnTable = InStr(1, TableFont, QueChar)
        
        If PositionOnTable = 0 Then
            PositionOnTable = InStr(1, TableFont, "?")
        End If
        
        For jY = 1 To TamY
            For jX = 1 To TamX
                If ScanDot(jX, jY, PositionOnTable, TamX, TamY) = True Then
                    BitBlt DDC, ActualX, ActualY, SprW, SprH, DCDotFore, 0, 0, SRCCOPY
                Else
                    BitBlt DDC, ActualX, ActualY, SprW, SprH, DCDotBack, 0, 0, SRCCOPY
                End If
                ActualX = ActualX + SprW + 1
            Next jX
            
            ActualX = TempX
            ActualY = ActualY + SprH + 1
        Next jY
        
        ActualY = 0
        TempX = TempX + Ancho
        ActualX = TempX
    
    Next i
    
    Render
    'PicMain.Refresh
    UserControl.Refresh
    
End Sub

Private Function InitializeStringFont(pFont As eFonts) As String
    Select Case pFont
        Case eFonts.NormalPixel_5x7
            InitializeStringFont = " !" & Chr$(34) & "#$%&'()*+,-./0123456789;:<=>?@ABCDEFGHIJKLMNÑOPQRSTUVWXYZ[\]^_`abcdefghijklmnñopqrstuvwxyz{|}~"
        Case eFonts.DPComic10x14
            InitializeStringFont = " !" & Chr$(34) & "#$%'()*+,-./0123456789;:<=>?@ABCDEFGHIJKLMNÑOPQRSTUVWXYZ[\]^_`abcdefghijklmnñopqrstuvwxyz{|}"
        Case eFonts.OutLand9x7
            InitializeStringFont = " !" & Chr$(34) & "()+,-./0123456789;:?ABCDEFGHIJKLMNÑOPQRSTUVWXYZ\_abcdefghijklmnñopqrstuvwxyz{|}"
    End Select
End Function

Private Function ScanDot(PosX As Integer, PosY As Integer, Position As Integer, FontWidth As Integer, FontHeight As Integer) As Boolean
'On Error Resume Next
Dim RealY As Integer
Dim RealX As Integer
    RealY = FontHeight - PosY
    RealX = (Position - 1) * (FontWidth + 1) + PosX
    ScanDot = False
    If pictBuff(RealX, RealY) = 0 Then
        ScanDot = True
    End If
End Function
