Attribute VB_Name = "mod_DX"
' ####################################################################################################################
' #
' #  MDL Model DirectX module - Copyright (c) TPD Software
' #
' #     This module interfaces with DirectX
' #
' #     Features:   - lighting
' #                 - rotation
' #                 - scaling
' #                 - backbuffering
' #                 - directinput keyboard/mouse
' #
' ####################################################################################################################

Public Const PI   As Single = 3.14159

Public DX         As DirectX8
Public D3D        As Direct3D8
Public D3DX       As D3DX8
Public D3DDevice  As Direct3DDevice8

Public DI         As DirectInput8
Public DIDevice1  As DirectInputDevice8
Public DIDevice2  As DirectInputDevice8

Public DIKState   As DIKEYBOARDSTATE
Public DIMState   As DIMOUSESTATE

Private TEXTFONT  As D3DXFont
Private TEXTRECT  As RECT
Private iFNT      As IFont
Private stdFNT    As StdFont

Private DISPMODE  As D3DDISPLAYMODE
Private D3DPP     As D3DPRESENT_PARAMETERS

Private matProj   As D3DMATRIX
Private matView   As D3DMATRIX
Private matWorld  As D3DMATRIX

Private mAngle    As D3DVECTOR
Private mScale    As D3DVECTOR
Private mPosition As D3DVECTOR
   
Private FPS       As Long
   
Public cMDL       As cls_MDL
Public cMD2       As cls_MD2

Public MDLType    As Long     ' 1=MDL, 2=MD2

Function DXLaunch(Container As Object) As Boolean

   On Error GoTo InitFailure
   
   ' ## create dx objects
   Set DX = New DirectX8
   Set D3DX = New D3DX8
   Set DI = DX.DirectInputCreate()
   Set D3D = DX.Direct3DCreate()
   
   Dim bFlags As CONST_D3DCREATEFLAGS
   Dim DIProp As DIPROPLONG
   
   With frmMenu
   
        ' ## get current display parameters
        D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DISPMODE
    
        ' ## initialize presentation
        D3DPP.Windowed = 1
        D3DPP.BackBufferCount = 1
        D3DPP.BackBufferFormat = DISPMODE.Format
        D3DPP.EnableAutoDepthStencil = 1
        D3DPP.AutoDepthStencilFormat = D3DFMT_D16
        D3DPP.SwapEffect = D3DSWAPEFFECT_COPY
               
        ' ## check if we can do some hardware accelaration
        If DeviceCapExists(D3DDEVCAPS_HWTRANSFORMANDLIGHT) Then
           bFlags = D3DCREATE_HARDWARE_VERTEXPROCESSING
           If DeviceCapExists(D3DDEVCAPS_PUREDEVICE) Then
              bFlags = bFlags Or D3DCREATE_PUREDEVICE
           End If
        Else
           bFlags = D3DCREATE_SOFTWARE_VERTEXPROCESSING
        End If
        
        ' ## create the device
        Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, _
                                         D3DDEVTYPE_HAL, _
                                         Container.hWnd, _
                                         bFlags, _
                                         D3DPP)
   
        ' ## set ambient light
        D3DDevice.SetRenderState D3DRS_AMBIENT, D3DColorXRGB(150, 150, 150)
        ' ## set cullmode to clockwise
        D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
        ' ## enable light
        D3DDevice.SetRenderState D3DRS_LIGHTING, 1
        ' ## enable z-buffer
        D3DDevice.SetRenderState D3DRS_ZENABLE, 1
        
        ' ## create an direct input keyboard device
        Set DIDevice1 = DI.CreateDevice("GUID_SysKeyboard")
       
        ' ## keyboard use
        DIDevice1.SetCommonDataFormat DIFORMAT_KEYBOARD
        ' ## link it to our form
        DIDevice1.SetCooperativeLevel Container.Parent.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
        
        ' ## create an direct input keyboard device
        Set DIDevice2 = DI.CreateDevice("guid_SysMouse")
       
        ' ## mouse use
        DIDevice2.SetCommonDataFormat DIFORMAT_MOUSE
        ' ## link it to our form
        DIDevice2.SetCooperativeLevel Container.Parent.hWnd, DISCL_FOREGROUND Or DISCL_NONEXCLUSIVE
        ' ## create a buffer
        DIProp.lHow = DIPH_DEVICE
        DIProp.lObj = 0
        DIProp.lData = 10
        DIDevice2.SetProperty "DIPROP_BUFFERSIZE", DIProp
        
        ' ## set perspective parameters
        D3DXMatrixPerspectiveFovLH matProj, PI / 4, 1, 0.1, 1000
        D3DDevice.SetTransform D3DTS_PROJECTION, matProj
        
   End With
   
   DXLaunch = True
      
   Exit Function
   
InitFailure:
      
End Function

Sub DXEnd()
   
   ' ## free up directx resources
   If Not DIDevice1 Is Nothing Then
      DIDevice1.Unacquire
      Set DIDevice1 = Nothing
   End If
   
   If Not DIDevice2 Is Nothing Then
      DIDevice2.Unacquire
      Set DIDevice2 = Nothing
   End If
   
   If Not TEXTFONT Is Nothing Then
      Set TEXTFONT = Nothing
   End If
   
   If Not D3DDevice Is Nothing Then
      Set D3DDevice = Nothing
   End If
   
   If Not DI Is Nothing Then
      Set DI = Nothing
   End If
   
   If Not D3D Is Nothing Then
      Set D3D = Nothing
   End If
   
   If Not DX Is Nothing Then
      Set DX = Nothing
   End If
   
End Sub

' ## make a light
Sub SetLight(Index As Long, LType As CONST_D3DLIGHTTYPE, pos As D3DVECTOR, Dir As D3DVECTOR, Range As Single, Color As D3DCOLORVALUE)

   Dim LIGHT As D3DLIGHT8
   
   With LIGHT
       .Type = LType
       .Position = pos
       .Direction = Dir
       .Range = Range
       .diffuse = Color
       .Attenuation0 = 0
       .Attenuation1 = 0.15
       .Attenuation2 = 0.015
   End With
   
   D3DDevice.SetLight Index, LIGHT
   D3DDevice.LightEnable 0, 1

End Sub

' ## test hardware capabilities
Function DeviceCapExists(D3DDEVCAPS As CONST_D3DDEVCAPSFLAGS) As Boolean
   Dim DevCaps As D3DCAPS8
   D3D.GetDeviceCaps D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, DevCaps
   DeviceCapExists = DevCaps.DevCaps And D3DDEVCAPS
End Function

' ## position the camera
Sub SetCamera(Eye As D3DVECTOR, At As D3DVECTOR, Up As D3DVECTOR)
   D3DXMatrixLookAtLH matView, Eye, At, Up
   D3DDevice.SetTransform D3DTS_VIEW, matView
End Sub

' ## simplify -> make a vector
Function MakeVector(x As Single, y As Single, z As Single) As D3DVECTOR
   With MakeVector
      .x = x
      .y = y
      .z = z
   End With
End Function

' ## simplify -> create a color value
Function CreateD3DColorVal(r As Byte, g As Byte, b As Byte, a As Byte) As D3DCOLORVALUE
   With CreateD3DColorVal
      .r = r
      .g = g
      .b = b
      .a = a
   End With
End Function

Sub InitializeMatrix()
   ' ## no rotation
   mAngle = MakeVector(0, 0, 0)
   ' ## scale factor 1
   mScale = MakeVector(1, 1, 1)
   ' ## origin position
   mPosition = MakeVector(0, 0, 0)
End Sub

' ## mainloop
Sub MainLoop()
   
   Dim tMat      As D3DMATRIX
   Dim oMat      As D3DMATRIX
   
   Dim AnimSleep As Long
   Dim AnimSpeed As Long
   Dim tFPS      As Long
   Dim cFPS      As Long
   
   AnimSpeed = 40
   AnimSleep = 0
     
   InitializeMatrix
   
   Do

      ' ## get keyboard and mouse inputs
      DIDevice1.Acquire
      DIDevice1.GetDeviceStateKeyboard DIKState
      On Error Resume Next
      DIDevice2.Acquire
      DIDevice2.GetDeviceStateMouse DIMState
      On Error GoTo 0

      ' ## set matrix identities
      D3DXMatrixIdentity tMat
      D3DXMatrixIdentity oMat
   
      ' ## rotate object with keyboard
      With mAngle
         
         If DIKState.Key(DIK_LEFT) <> 0 Then .x = .x + 0.01
         If DIKState.Key(DIK_RIGHT) <> 0 Then .x = .x - 0.01
         
         If .x < 0 Then .x = 2 * PI
         If .x > 2 * PI Then .x = 0
         
         If DIKState.Key(DIK_UP) <> 0 Then .y = .y + 0.01
         If DIKState.Key(DIK_DOWN) <> 0 Then .y = .y - 0.01
         
         If .y < 0 Then .y = 2 * PI
         If .y > 2 * PI Then .y = 0
                  
         If DIKState.Key(DIK_PGDN) <> 0 Then .z = .z + 0.01
         If DIKState.Key(DIK_DELETE) <> 0 Then .z = .z - 0.01
                  
         If .z < 0 Then .z = 2 * PI
         If .z > 2 * PI Then .z = 0
                
         D3DXMatrixRotationYawPitchRoll tMat, .x, .y, .z
         D3DXMatrixMultiply oMat, oMat, tMat
                
      End With
      
      ' ## set animation speed with left mouse button
      If DIMState.Buttons(0) Then
          AnimSpeed = AnimSpeed + CLng(DIMState.lY)
          If AnimSpeed < 5 Then AnimSpeed = 5
          If AnimSpeed > 200 Then AnimSpeed = 200
      End If
   
      ' ## set uniform scale of model for all 3 dimensions with right mouse button
      With mScale
         If DIMState.Buttons(1) Then
            .x = .x + DIMState.lY * 0.01
            If .x < 0 Then .x = 0
            .y = .x
            .z = .x
         End If
         D3DXMatrixScaling tMat, .x, .y, .z
         D3DXMatrixMultiply oMat, oMat, tMat
      End With
      
      ' ## set model position
      With mPosition
         D3DXMatrixTranslation tMat, .x, .y, .z
         D3DXMatrixMultiply oMat, oMat, tMat
      End With
      
      ' ## delay frames for each sequence
      If AnimSleep >= AnimSpeed Then
         ' ## advance to next frame
         Select Case MDLType
         Case 1
           cMDL.NextFrame
         Case 2
           cMD2.NextFrame
         End Select
         AnimSleep = 0
      Else
         AnimSleep = AnimSleep + 1
      End If
      
      ' ## set interpolation value between frames according to the current delay time
      Select Case MDLType
      Case 1
        cMDL.SetFrameIP AnimSleep / AnimSpeed
      Case 2
        cMD2.SetFrameIP AnimSleep / AnimSpeed
      End Select
         
      ' ## apply the world matrix
      D3DDevice.SetTransform D3DTS_WORLD, oMat
      
      ' ## render to screen
      Render
      
      ' ## fps calculation
      If TimeElapsed(tFPS, 1000) Then
         tFPS = GetTickCount()
         FPS = cFPS
         cFPS = 0
      Else
         cFPS = cFPS + 1
      End If
      
      DoEvents
   Loop
   
End Sub

Sub Render()
  
  With D3DDevice
    .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H3F3F3F, 1#, 0
    .BeginScene
        
        ' ## render model
        Select Case MDLType
        Case 1
           cMDL.Render
        Case 2
           cMD2.Render
        End Select
       
        ' ## draw frames per second and some text
        DrawFPS
              
    .EndScene
    .Present ByVal 0, ByVal 0, 0, ByVal 0
  End With
  
End Sub

' ## create a font for use with DrawText
Sub CreateFont()
   Set stdFNT = New StdFont
   stdFNT.Name = "Arial"
   stdFNT.size = 10
   stdFNT.Bold = True
   Set iFNT = stdFNT
   Set TEXTFONT = D3DX.CreateFont(D3DDevice, iFNT.hFont)
End Sub

' ## draw some infotext to the screen
Sub DrawFPS()
   Dim Msg As String
   TEXTRECT.Left = 0
   TEXTRECT.Top = 0
   TEXTRECT.Right = 250
   TEXTRECT.bottom = 100
   Msg = "Right mouse: Zoom"
   Msg = Msg & vbCrLf & "Left mouse: Animation speed"
   Msg = Msg & vbCrLf & "Cursor keys/Del/Pgdn: Rotate"
   Msg = Msg & vbCrLf & "FPS: " & FPS
   D3DX.DrawText TEXTFONT, &HFFFFFFFF, Msg, TEXTRECT, DT_TOP Or DT_LEFT
End Sub
