VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub MainInit(hWnd As Long)

    Dim result As Boolean
    
    'init main objects
    
    result = InitD3D(hWnd)
    If result = False Then
        MsgBox "InitD3D failed."
        PclsSubs.CleanUp
        End
    End If
    
    result = InitScreen
    If result = False Then
        MsgBox "InitScreen failed."
        PclsSubs.CleanUp
        End
    End If
    
    TextLenght = Len(Text)
    LeftText = 1
    
    'enable timer (tmrMain) - it will do render sub
    frmMain.tmrMain.Enabled = True
    
End Sub


Public Function InitD3D(hWnd As Long) As Boolean
On Error GoTo ErrHandler            'If something goes wrong go to errhandler

    Set D3D = DX.Direct3DCreate()   'creates D3D Object
    If D3D Is Nothing Then Exit Function 'if object isn't created exit sub

    Dim mode As D3DDISPLAYMODE      'Dim variable to find display mode
    'Retrieve the Display mode format
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode

    'Dims d3dpp as a D3DPRESENT_PARAMETERS object. Each field is filled in to set
    'up for a new Direct3D Device.
    Dim d3dpp As D3DPRESENT_PARAMETERS
    d3dpp.Windowed = False
    d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    d3dpp.BackBufferFormat = mode.Format
    d3dpp.BackBufferHeight = 600
    d3dpp.BackBufferWidth = 800
    d3dpp.hDeviceWindow = hWnd
    'Creates D3DDevice
    Set D3Ddevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If D3Ddevice Is Nothing Then Exit Function
        
    InitD3D = True                  'Everything went fine
    Exit Function
    
ErrHandler:                         'Something went wrong
    InitD3D = False                 'Set function to false so program stops running
    
End Function

Public Function InitScreen() As Boolean
    
    On Error GoTo Blad
    
    'Setup font for rendering text
    'font name:
    Fnt.Name = "Impact"
    'font size:
    Fnt.Size = 16
    'font bold:
    Fnt.Bold = False
    'font underline:
    Fnt.Underline = True
    
    'create font device
    Set MainFontDesc = Fnt
    Set MainFont = D3DX.CreateFont(D3Ddevice, MainFontDesc.hFont)
    
    'create region for text
    TextRect.Top = 5
    TextRect.Left = 5
    TextRect.bottom = 70
    TextRect.Right = 100

    'Load start screen image (We are using 800x600 image)
    Set DX8BackGround = D3Ddevice.CreateRenderTarget(800, 600, D3DFMT_X8R8G8B8, D3DMULTISAMPLE_NONE, True)
    D3DX.LoadSurfaceFromFile DX8BackGround, ByVal 0&, ByVal 0&, App.Path & "\Orion.jpg", ByVal 0&, D3DX_FILTER_NONE, 0, ByVal 0&
    
    InitScreen = True
    Exit Function
    
Blad:
    InitScreen = False
    
End Function

