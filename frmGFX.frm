VERSION 5.00
Begin VB.Form frmGFX 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicSprFonts 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   210
      Index           =   2
      Left            =   60
      Picture         =   "frmGFX.frx":0000
      ScaleHeight     =   14
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   1060
      TabIndex        =   5
      Top             =   420
      Width           =   15900
   End
   Begin VB.PictureBox PicSprFonts 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   1
      Left            =   60
      Picture         =   "frmGFX.frx":0A44
      ScaleHeight     =   7
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   831
      TabIndex        =   4
      Top             =   300
      Width           =   12465
   End
   Begin VB.PictureBox PicDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   1
      Left            =   720
      ScaleHeight     =   10
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   10
      TabIndex        =   3
      Top             =   900
      Width           =   150
   End
   Begin VB.PictureBox PicDots 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Index           =   0
      Left            =   540
      ScaleHeight     =   10
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   10
      TabIndex        =   2
      Top             =   900
      Width           =   150
   End
   Begin VB.PictureBox PicFont 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   60
      ScaleHeight     =   9
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   693
      TabIndex        =   1
      Top             =   0
      Width           =   10395
   End
   Begin VB.PictureBox PicSprFonts 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   105
      Index           =   0
      Left            =   60
      Picture         =   "frmGFX.frx":1142
      ScaleHeight     =   7
      ScaleMode       =   3  'Píxel
      ScaleWidth      =   592
      TabIndex        =   0
      Top             =   180
      Width           =   8880
   End
End
Attribute VB_Name = "frmGFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

