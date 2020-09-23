VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectDraw7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'class variables
Public DXObj As DirectX7
Public DDObj As DirectDraw7
Public Screen As clsDDSurface7
Dim DDObjClip As DirectDrawClipper
Dim rScreen As DxVBLib.RECT
Dim hSourceWindow As Long

Private Sub Class_Initialize()
    'create the DirectX7 object
    Set DXObj = New DirectX7
    'create the DirectDraw object
    Set DDObj = DXObj.DirectDrawCreate("")
End Sub

Private Sub Class_Terminate()
    Set DDObj = Nothing
    Set DXObj = Nothing
End Sub

Public Function Startup(ByVal hWnd As Long, ByVal lWidth As Long, _
    ByVal lHeight As Long, ByVal lBPP As Long, _
    ByVal bFullscreen As Boolean) As Boolean
    
    hSourceWindow = hWnd
    If bFullscreen Then
        'set the color depth
        DDObj.SetCooperativeLevel hWnd, DDSCL_FULLSCREEN Or _
            DDSCL_EXCLUSIVE
        If lBPP < 1 Or lBPP > 32 Then lBPP = 32
        
        'set the display mode
        DDObj.SetDisplayMode lWidth, lHeight, lBPP, 0, DDSDM_DEFAULT
    Else
        'windowed mode
        DDObj.SetCooperativeLevel hWnd, DDSCL_NORMAL
    End If
   
    'set up the primary display surface
    Set Screen = New clsDDSurface7
    
    'create the primary display surface
    Startup = Screen.Create(DDObj, lWidth, lHeight, True)
    If Startup Then
        'create the clipper object
        Set DDObjClip = DDObj.CreateClipper(0)
        
        'assoiciate the window handle with the clipper
        DDObjClip.SetHWnd hSourceWindow
        
        'clip blitting routines to the window
        Screen.Surface.SetClipper DDObjClip
    End If
End Function

Public Sub Shutdown()
    DDObj.RestoreDisplayMode
    DDObj.SetCooperativeLevel hSourceWindow, DDSCL_NORMAL
End Sub

Public Sub Blt(ByRef ddsSource As DirectDrawSurface7, _
    ByRef srcRect As DxVBLib.RECT)
    Screen.Blt ScreenRect, ddsSource, srcRect
End Sub

Public Sub Draw(ByRef cddsSource As clsDDSurface7, _
    ByRef srcRect As DxVBLib.RECT)
    Screen.Blt ScreenRect, cddsSource.Surface, srcRect
End Sub
Public Property Get ScreenRect() As DxVBLib.RECT
    DXObj.GetWindowRect hSourceWindow, rScreen
    ScreenRect = rScreen
End Property


