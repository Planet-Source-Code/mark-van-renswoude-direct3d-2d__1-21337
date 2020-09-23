Attribute VB_Name = "modMain"
Option Explicit

'// Main DirectX Declarations
Public mDX As New DirectX7
Public mDDraw As DirectDraw7
Public mD3D As Direct3D7
Public mD3DDevice As Direct3DDevice7

'// Vertex Declarations
Public mtlSprite(3) As D3DTLVERTEX
Public mtlBackground(3) As D3DTLVERTEX

'// Surfaces / Textures Declarations
Public msFront As DirectDrawSurface7
Public msBack As DirectDrawSurface7
Public msFrame1 As DirectDrawSurface7
Public msFrame2 As DirectDrawSurface7

'// Screen Declarations
Public SCREEN_WIDTH As Long
Public SCREEN_HEIGHT As Long
Public SCREEN_DEPTH As Long
Public SCREEN_BACKCOLOR As Long

'// Other Declarations
Public mbRunning As Boolean
Public Sub InitDX(Width As Long, Height As Long, Depth As Long, DeviceGUID As String)
    Dim ddsd As DDSURFACEDESC2
    Dim caps As DDSCAPS2
    
    '// Create DirectDraw object
    Set mDDraw = mDX.DirectDrawCreate("")
    
    '// Set Cooperative Level (fullscreen, exclusive access)
    mDDraw.SetCooperativeLevel frmMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN Or DDSCL_ALLOWREBOOT
    mDDraw.SetDisplayMode Width, Height, Depth, 0, DDSDM_DEFAULT
    
    '// Create primary surface
    ddsd.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    ddsd.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_3DDEVICE Or DDSCAPS_PRIMARYSURFACE
    ddsd.lBackBufferCount = 1
    
    Set msFront = mDDraw.CreateSurface(ddsd)
    
    '// Create the backbuffer (used for 3D drawing)
    caps.lCaps = DDSCAPS_BACKBUFFER Or DDSCAPS_3DDEVICE
    Set msBack = msFront.GetAttachedSurface(caps)
    
    '// Create Direct3D
    Set mD3D = mDDraw.GetDirect3D
    
    '// Set the Device
    Set mD3DDevice = mD3D.CreateDevice(DeviceGUID, msBack)
End Sub

Public Sub Start(DeviceGUID As String)
    '// Show the Form
    frmMain.Show
    DoEvents
    
    '// Screen Settings
    SCREEN_WIDTH = 640
    SCREEN_HEIGHT = 480
    SCREEN_DEPTH = 16
    SCREEN_BACKCOLOR = RGB2DX(0, 96, 184)
    
    '// Initialize DirectX at 640x480x16
    Call InitDX(SCREEN_WIDTH, SCREEN_HEIGHT, SCREEN_DEPTH, DeviceGUID)
    
    '// Load Textures and Surfaces
    Set msFrame1 = CreateTexture(App.Path & "\Frame1.bmp", 64, 64, Magenta)
    Set msFrame2 = CreateTexture(App.Path & "\Frame2.bmp", 64, 64, Magenta)
    
    '// Start Running
    Call MainLoop
    
    '// End it all
    Call Terminate
End Sub


Public Sub ClearDevice()
    '// Clear the device for drawing operations
    Dim rClear(0) As D3DRECT
    
    rClear(0).X2 = SCREEN_WIDTH
    rClear(0).Y2 = SCREEN_HEIGHT
    mD3DDevice.Clear 1, rClear, D3DCLEAR_TARGET, SCREEN_BACKCOLOR, 0, 0
End Sub

Public Sub MainLoop()
    Dim iOffsetAngle As Integer
    Dim iAngle As Integer
    Dim iFrame As Integer
    Dim lXOffset As Long
    Dim lYOffset As Long
    Dim lColor As Long
    
    mbRunning = True
    iAngle = 180
    iOffsetAngle = -45
    lColor = RGB2DX(255, 255, 255)
    
    '// Draw until the program should stop running
    Do While mbRunning
        '// Rotate sprite
        iAngle = iAngle + 2
        If iAngle > 360 Then iAngle = iAngle - 360
        
        '// Move frame
        If iAngle / 5 = CInt(iAngle / 5) Then
            iFrame = iFrame + 1
            If iFrame = 2 Then iFrame = 0
        End If
        
        '// Calculate Offset
        iOffsetAngle = iOffsetAngle + 2
        If iOffsetAngle > 360 Then iOffsetAngle = iOffsetAngle - 360
        
        lXOffset = CalcCoordX(SCREEN_WIDTH / 2, 75, iOffsetAngle)
        lYOffset = CalcCoordY(SCREEN_HEIGHT / 2, 75, iOffsetAngle)
        
        '// Create the four vertices which make the sprite
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle), CalcCoordY(lYOffset, 32, iAngle), 0, 1, lColor, 0, 0, 0, mtlSprite(0))
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle - 90), CalcCoordY(lYOffset, 32, iAngle - 90), 0, 1, lColor, 0, 1, 0, mtlSprite(1))
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle + 90), CalcCoordY(lYOffset, 32, iAngle + 90), 0, 1, lColor, 0, 0, 1, mtlSprite(2))
        Call mDX.CreateD3DTLVertex(CalcCoordX(lXOffset, 32, iAngle + 180), CalcCoordY(lYOffset, 32, iAngle + 180), 0, 1, lColor, 0, 1, 1, mtlSprite(3))
        
        '// Clear the device
        Call ClearDevice
        
        '// Start the scene
        mD3DDevice.BeginScene
        
            '// Enable color key
            mD3DDevice.SetRenderState D3DRENDERSTATE_COLORKEYENABLE, True
            
            '// Set texture
            If iFrame = 0 Then
                mD3DDevice.SetTexture 0, msFrame1
            Else
                mD3DDevice.SetTexture 0, msFrame2
            End If
            
            '// Draw sprite
            Call mD3DDevice.DrawPrimitive(D3DPT_TRIANGLESTRIP, D3DFVF_TLVERTEX, mtlSprite(0), 4, D3DDP_DEFAULT)
            
        '// End the scene
        mD3DDevice.EndScene
        
        '// Draw Text
        msBack.SetForeColor RGB(255, 255, 255)
        msBack.DrawText 5, 5, "Press 'Any' key to exit...", False
        
        '// Draw Circles
        msBack.SetForeColor RGB(33, 148, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 140
        msBack.SetForeColor RGB(65, 170, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 130
        msBack.SetForeColor RGB(135, 197, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 120
        msBack.SetForeColor RGB(165, 225, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 110
        msBack.SetForeColor RGB(255, 255, 255)
        msBack.DrawCircle SCREEN_WIDTH / 2, SCREEN_HEIGHT / 2, 100
        
        '// Flip
        msFront.Flip Nothing, DDFLIP_WAIT
        DoEvents
    Loop
End Sub

Public Sub Terminate()
    '// Clean up DirectX
    Call mDDraw.RestoreDisplayMode
    Call mDDraw.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    
    Set mD3DDevice = Nothing
    Set mD3D = Nothing
    Set msBack = Nothing
    Set msFront = Nothing
    Set mDDraw = Nothing
    Set mDX = Nothing
    
    Unload frmMain
End Sub


