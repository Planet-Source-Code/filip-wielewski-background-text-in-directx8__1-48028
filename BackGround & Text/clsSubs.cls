VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub KeyPressed(Key As Integer)
    
    Select Case Key
        Case vbKeyEscape
            End
    End Select
    
End Sub

Public Sub CleanUp()
    
    'free memory
    Set DX = Nothing
    Set D3DX = Nothing
    Set D3D = Nothing
    Set D3Ddevice = Nothing
    Set DX8BackGround = Nothing

End Sub

Public Sub Render()
    
    On Error GoTo Blad
    
    Dim ActualText As String
    
    LeftText = LeftText + 1
    If LeftText = (TextLenght + 5) Then LeftText = 1
    ActualText = Left(Text, LeftText)
    
    'Clear screen to a black color
    D3Ddevice.Clear 0, ByVal 0&, D3DCLEAR_TARGET, &HAAEEEE, 1, 0
     
    'Copy background image to back buffer
    D3Ddevice.CopyRects DX8BackGround, ByVal 0&, 0, D3Ddevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO), ByVal 0&
    
    'begin scene (now you can draw graphic)
    D3Ddevice.BeginScene
    
    'Render FPS-text
    D3DX.DrawText MainFont, &HFFFFFFFF, ActualText, TextRect, DT_CENTER
    
    'End the scene
    D3Ddevice.EndScene
        
    'Present the backbuffer to the screen
    D3Ddevice.Present ByVal 0&, ByVal 0&, 0, ByVal 0&
    
    Exit Sub
    
Blad:
    CleanUp
    MsgBox "Sub Render failed."
    End
    
End Sub

