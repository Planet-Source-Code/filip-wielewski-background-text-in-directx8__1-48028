VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "TextPic"
   ClientHeight    =   3420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4905
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   228
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   327
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1800
      Top             =   1440
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================='
'=                                                            ='
'= ======AUTHOR======                                         ='
'= THIS IS A FREE CODE                                        ='
'= BY FILIP WIELEWSKI                                         ='
'= E-MAIL: WIELFILIST@WP.PL                                   ='
'=                                                            ='
'= ======SORRY FOR:======                                     ='
'= my bad english which I use in descriptions :]              ='
'=                                                            ='
'=============================================================='

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    'check KeyAscii of pressed key
    PclsSubs.KeyPressed KeyAscii
    
End Sub

Private Sub Form_Load()
    
    'init DirectX8 objects
    PclsInits.MainInit frmMain.hWnd

End Sub

Private Sub tmrMain_Timer()
    
    'do render (draw all graphic)
    PclsSubs.Render

End Sub
