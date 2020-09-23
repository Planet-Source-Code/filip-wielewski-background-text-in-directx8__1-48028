Attribute VB_Name = "modDeclare"
Option Explicit

'DirectX main variables
Public DX As New DirectX8
Public D3DX As New D3DX8
Public D3D As Direct3D8
Public D3Ddevice As Direct3DDevice8
Public DX8BackGround As Direct3DSurface8

'Classes
Public PclsInits As New clsInits
Public PclsSubs As New clsSubs

'Font variables
Public MainFont As D3DXFont         'This is the main font object
Public MainFontDesc As IFont        'We use this temporarily to setup the font
Public TextRect As RECT             'Text draw place
Public Fnt As New StdFont           'Used to describe and setup the font

'Miscellanous variables
Public TextLenght As Byte
Public LeftText As Byte

'Consts
Public Const Text = "Filip" & vbCrLf & "Wielewski"
