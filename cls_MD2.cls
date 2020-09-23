VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_MD2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ####################################################################################################################
' #
' #  MD2 Model class - Copyright (c) TPD Software
' #
' #     This class handles all MDL related stuff
' #
' #     Features:   - multiple and custom skin support
' #                 - uses precalculated normals
' #                 - animation with frame interpolation
' #                 - vertex buffering
' #                 - bitmaps supported: BMP and PCX
' #
' ####################################################################################################################

Option Explicit

' --------------- PCX related ---------------

Private Type PCXHeader
   manufacturer      As Byte
   version           As Byte
   encoding          As Byte
   bpp               As Byte
   minx              As Integer
   miny              As Integer
   maxx              As Integer
   maxy              As Integer
   horz_res          As Integer
   vert_res          As Integer
   ega_palette(47)   As Byte
   reserved          As Byte
   planes            As Byte
   bpline            As Integer
   palette_type      As Integer
   padding(57)       As Byte
End Type

' --------------- DX related ---------------

Private Const FVF_MDLVERTEX As Long = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Private cVBuffer            As Direct3DVertexBuffer8
Private cTexture()          As Direct3DTexture8
Private cMaterial           As D3DMATERIAL8
Private c3DDevice           As Direct3DDevice8

' --------------- MD2 related --------------

Private Type tHeader
    id           As String * 4
    version      As Long
    skinwidth    As Long
    skinheight   As Long
    framesize    As Long
    num_skins    As Long
    num_xyz      As Long
    num_st       As Long
    num_tris     As Long
    num_glcmds   As Long
    num_frames   As Long
    ofs_skins    As Long
    ofs_st       As Long
    ofs_tris     As Long
    ofs_frames   As Long
    ofs_glcmds   As Long
    ofs_end      As Long
End Type

Private Type tSkins
    sname        As String * 64
End Type
    
Private Type tSkinCoords
    s            As Integer
    t            As Integer
End Type
    
Private Type tVertexIndexes
    x            As Byte
    y            As Byte
    z            As Byte
    iLightNormal As Byte
End Type

Private Type tTriangleIndexes
    vindex(2)    As Integer
    tindex(2)    As Integer
End Type

Private Type tTriangleVertex
    p            As D3DVECTOR
    n            As D3DVECTOR
    t            As D3DVECTOR2
End Type
    
Private Type tAliasFrame
    ascale       As D3DVECTOR
    atrans       As D3DVECTOR
    aname        As String * 16
End Type

Private Type tAliasFrameData
    adata()      As tVertexIndexes
End Type

Private Type tFrame
    fname        As String * 16
    fdata()      As tTriangleVertex
End Type

Private Type tFrameIndex
    fname        As String * 16
    first        As Long
    last         As Long
End Type

Private Type PalColor
    r            As Byte
    g            As Byte
    b            As Byte
End Type

Private Pal(255)        As PalColor

Private hMDL            As tHeader
Private SkinNames()     As tSkins
Private SkinCoords()    As tSkinCoords
Private Triangles()     As tTriangleIndexes
Private AliasFrame      As tAliasFrame
Private AliasData       As tAliasFrameData
Private Frames()        As tFrame
Private FramesIndex()   As tFrameIndex

Private cCustSkin       As Long

Private cModelLoaded    As Boolean
Private cIP             As Boolean
Private cTextured       As Boolean
Private cSkin           As Long
Private cFrame          As Long
Private nFrame          As Long
Private cAnim           As Long
Private cIPv            As Single
Private cPos            As D3DVECTOR
Private cErr            As MDL_ERRORS

' ## Load model data
Sub Load(File As String)

  ' On Error GoTo LoadError
   
   Dim i As Long
   Dim j As Long

   If c3DDevice Is Nothing Then Exit Sub

   ' ## clear resources
   FreeResources
   
   Open File For Binary As #1
   
     cModelLoaded = True
   
     ' ## get MDL header
     Get #1, , hMDL
     If hMDL.id <> "IDP2" Then
        Close #1
        cErr = MDL_INVALID_ID
        Exit Sub
     End If
     
     If hMDL.version <> 8 Then
        Close #1
        cErr = MDL_INVALID_VERSION
        Exit Sub
     End If
     
     ' ## allocate memory for model data
     ReDim SkinCoords(hMDL.num_st - 1)
     ReDim Triangles(hMDL.num_tris - 1)
     ReDim AliasData.adata(hMDL.num_xyz - 1)
     ReDim Frames(hMDL.num_frames - 1)
     ReDim FaceIndex(hMDL.num_tris - 1)
     
     If hMDL.num_skins > 0 Then
        ReDim cTexture(hMDL.num_skins - 1)
        ReDim SkinNames(hMDL.num_skins - 1)
        ' ## read skin texture names !
        ' ## md2 stores skins externally
        Seek #1, hMDL.ofs_skins + 1
        Get #1, , SkinNames()
        
        ' ## overwrite custom skins when numskins in header > 0
        If cCustSkin > 0 Then
           ReDim SkinNamesReal(0)
           cCustSkin = 0
           cErr = MDL_CUSTOM_SKIN_OVERWRITTEN
        End If
     Else
        ' ## allocate memory for custom skins
        If cCustSkin > 0 Then
           ReDim cTexture(cCustSkin - 1)
        End If
     End If
     
     ' ## read in texture coordinates
     Seek #1, hMDL.ofs_st + 1
     Get #1, , SkinCoords()
     
     ' ## read model triangles
     Seek #1, hMDL.ofs_tris + 1
     Get #1, , Triangles()
     
     ' ## read animation frames
     Seek #1, hMDL.ofs_frames + 1
     For i = 0 To hMDL.num_frames - 1
        Get #1, , AliasFrame
        Get #1, , AliasData.adata()
        ReDim Frames(i).fdata(hMDL.num_xyz - 1)
        Frames(i).fname = AliasFrame.aname
        For j = 0 To hMDL.num_xyz - 1
           Frames(i).fdata(j).p.x = AliasFrame.ascale.y * AliasData.adata(j).y + AliasFrame.atrans.y
           Frames(i).fdata(j).p.y = AliasFrame.ascale.z * AliasData.adata(j).z + AliasFrame.atrans.z
           Frames(i).fdata(j).p.z = AliasFrame.ascale.x * AliasData.adata(j).x + AliasFrame.atrans.x
           Frames(i).fdata(j).n.x = Normals(AliasData.adata(j).iLightNormal, 1)
           Frames(i).fdata(j).n.y = Normals(AliasData.adata(j).iLightNormal, 2)
           Frames(i).fdata(j).n.z = Normals(AliasData.adata(j).iLightNormal, 0)
        Next j
     Next i
     
     ' ## create a lookup buffer for the animations start and end indexes
     GetAnimationIndexes
     
     ' ## load skins
     LoadSkins File
     
     ' ## allocate enough memory for the vertex buffer
     ' ## (structure length * num triangles * num vertices per triangle (3))
     Set cVBuffer = c3DDevice.CreateVertexBuffer(Len(Frames(0).fdata(0)) * hMDL.num_tris * 3, 0, FVF_MDLVERTEX, D3DPOOL_MANAGED)
     
   Close #1
   
  ' ## reset index variables as they change when loading a new model
   cSkin = 0
   cFrame = 0
   nFrame = 0
   cAnim = 0
   cCustSkin = 0
   cErr = MDL_OK
   
   Exit Sub

LoadError:

   Close #1
   cModelLoaded = False
   cCustSkin = 0
   cErr = MDL_LOAD_ERROR
     
End Sub

' ## load the skins
Private Sub LoadSkins(ByVal File As String)
   
   Dim abspath As String
   Dim bitmap  As String
   Dim ext     As String
   Dim i       As Long
   
   ' ## create a temporary buffer
   ReDim SkinBuffer(hMDL.skinwidth * hMDL.skinheight - 1) As Byte
   
   ' ## obtain path to the textures
   If InStr(1, File, "\") Then
      abspath = Left(File, InStrRev(File, "\"))
   Else
      abspath = ""
   End If
   
   ' ## read in all skins
   For i = 0 To UBound(SkinNames)
      bitmap = SkinNames(i).sname
      bitmap = Replace(bitmap, "\", "/")
      bitmap = Left(bitmap, InStr(1, bitmap, vbNullChar) - 1)
      bitmap = Mid(bitmap, InStrRev(bitmap, "/") + 1)
      ext = LCase(Mid(bitmap, InStrRev(bitmap, ".") + 1))
      ' ## bitmap and pcx skins are supported
      Select Case ext
        Case "bmp"
          If Len(Dir(abspath & bitmap, vbNormal)) > 0 Then
             Set cTexture(i) = D3DX.CreateTextureFromFile(c3DDevice, abspath & bitmap)
          Else
             cErr = MDL_MISSING_SKIN
          End If
        Case "pcx"
          If LoadPCX(abspath & bitmap, SkinBuffer()) Then
             Set cTexture(i) = CreateRGBTextureFromPalettizedBitmap(SkinBuffer())
          Else
             cErr = MDL_MISSING_SKIN
          End If
      End Select
   Next i
   
   ' ## clean temporary buffer
   ReDim SkinBuffer(0)
   
End Sub

' ## Make a lookup buffer for all animations in the model
' ## specific for Q2, so may be adjusted for custom models
Private Sub GetAnimationIndexes()
  
   Dim an As String
   Dim la As String
   Dim ln As Long
   Dim fi As Long
   Dim i  As Long
   
   ' ## loop trough al frames
   For i = 0 To hMDL.num_frames - 1
      ' ## strip the name
      an = Left(Frames(i).fname, InStr(1, Frames(i).fname, vbNullChar) - 1)
      ln = Len(an)
      Do While IsNumeric(Mid(an, ln, 1))
         ln = ln - 1
      Loop
      ln = ln + 1
      an = Left(an, ln)
      If an <> la Then
         ReDim Preserve FramesIndex(fi)
         ' ## for each new animation save it's starting frame
         With FramesIndex(fi)
           .fname = an
           .first = i
         End With
         fi = fi + 1
      End If
      la = an
   Next i
   ' ## now that we have the starting frames
   ' ## we can calculate the ending frames
   For i = 0 To UBound(FramesIndex) - 1
      FramesIndex(i).last = FramesIndex(i + 1).first - 1
   Next i
   ' ## set the last frame of the last animation
   FramesIndex(UBound(FramesIndex)).last = hMDL.num_frames - 1
   
End Sub

' ## when the model uses PCX files as skins this sub comes in handy
Private Function LoadPCX(File As String, BitmapBuffer() As Byte) As Boolean
    
    Dim Hdr          As PCXHeader
    Dim tpcxBuf()    As Byte
    Dim fOff         As Long
    Dim bOff         As Long
    Dim Cidx         As Byte
    Dim Repeat       As Long
    Dim ff           As Long
    
    ' ## check if skin exists
    If Len(Dir(File, vbNormal)) = 0 Then Exit Function
    
    ff = FreeFile
    
    Open File For Binary As #ff
    
       ' ## read the header
       Get #ff, , Hdr
       If Hdr.manufacturer <> 10 Then
          GoTo PCXFailure
       End If
       If Hdr.version <> 5 Then
          GoTo PCXFailure
       End If
       
       ' ## create buffers
       ReDim tpcxBuf(CLng(Hdr.maxx + 1) * CLng(Hdr.maxy + 1) - 1)
       ReDim tpcxPal(255)
       
       ' ## read pcx type
       fOff = Seek(ff)
       If Hdr.bpp = 8 And Hdr.planes = 1 Then
          Seek #ff, LOF(ff) - (UBound(tpcxPal) * 3) - 3
          Get #ff, , Cidx
          If Cidx <> 12 Then
             GoTo PCXFailure
          End If
          Get #ff, , Pal()
          Seek #ff, fOff
       Else
          GoTo PCXFailure
       End If
       
       ' ## read pcx data and decode it
       bOff = 0
       Do
           Get #ff, , Cidx
           If Cidx > 191 Then
              Repeat = Cidx - 192
              Get #ff, , Cidx
              Do
                tpcxBuf(bOff) = Cidx
                bOff = bOff + 1
                Repeat = Repeat - 1
              Loop Until Repeat = 0
           Else
              tpcxBuf(bOff) = Cidx
              bOff = bOff + 1
           End If
       Loop Until bOff >= UBound(tpcxBuf)
        
    Close #ff
    
    ' ## return bitmap data
    BitmapBuffer() = tpcxBuf()
    
    ' ## succesful load
    LoadPCX = True
    
    Exit Function
    
PCXFailure:
    Close #ff

End Function

' ## This sub coverts the model's 8 bit palette texture to an RGB equivalent
' ## which can be used as a valid texture in DX
Private Function CreateRGBTextureFromPalettizedBitmap(BitmapBuffer() As Byte) As Direct3DTexture8

   Dim tmpT As Direct3DTexture8
   Dim tmpS As Direct3DSurface8
   Dim desS As Direct3DSurface8
   Dim lr   As D3DLOCKED_RECT
   Dim Offs As Long
   Dim i    As Long

   ' ## allocate memory which can hold the new bitmap format
   ReDim nbuf(UBound(BitmapBuffer) * 3) As Byte
  
   ' ## copy the 8-bit bitmap data into the new array
   ' ## using a palette for lookup
   Offs = 0
   For i = 0 To UBound(BitmapBuffer) - 1
      ' ## directx reverses color bytes RGB => BGR
      nbuf(Offs + 0) = Pal(BitmapBuffer(i)).b
      nbuf(Offs + 1) = Pal(BitmapBuffer(i)).g
      nbuf(Offs + 2) = Pal(BitmapBuffer(i)).r
      Offs = Offs + 3
   Next i
   
   ' ## create a temporary surface in RGB format
   Set tmpS = c3DDevice.CreateImageSurface(hMDL.skinwidth, hMDL.skinheight, D3DFMT_R8G8B8)
   
   ' ## copy the array to the temporary surface
   tmpS.LockRect lr, ByVal 0, D3DLOCK_NOSYSLOCK
      DXCopyMemory ByVal lr.pBits, nbuf(0), UBound(nbuf) + 1
   tmpS.UnlockRect
   
   ' ## create an empty X8R8G8B8 texture object
   Set tmpT = c3DDevice.CreateTexture(hMDL.skinwidth, hMDL.skinheight, 1, 0, D3DFMT_X8R8G8B8, D3DPOOL_MANAGED)
   
   ' ## get surface 0 from the texture object
   Set desS = tmpT.GetSurfaceLevel(0)
   
   ' ## copy the buffer onto the texturesurface
   D3DX.LoadSurfaceFromSurface desS, ByVal 0, ByVal 0, tmpS, ByVal 0, ByVal 0, D3DX_FILTER_LINEAR, 0
   
   ' ## return and cleanup
   Set CreateRGBTextureFromPalettizedBitmap = tmpT
   Set tmpT = Nothing
   Set tmpS = Nothing
   Set desS = Nothing
   
End Function

' ## Render the model
Sub Render()

   ' ## only render when a valid model is loaded
   If cModelLoaded And Not c3DDevice Is Nothing Then
      
      ReDim tVerts(hMDL.num_tris * 3 - 1) As tTriangleVertex   ' ## model triangle buffer
      Dim i                               As Long
      Dim j                               As Long
      
      With c3DDevice
      
      ' ## set the vertex format
      .SetVertexShader FVF_MDLVERTEX
      
      ' ## set the model's material
      .SetMaterial cMaterial
      
      ' ## set the model's texture
      If cTextured And Not cTexture(cSkin) Is Nothing Then
         .SetTexture 0, cTexture(cSkin)
      Else
         .SetTexture 0, Nothing
      End If
      
      ' ## draw the models triangles
      For i = 0 To hMDL.num_tris - 1
         
         ' ## fill the buffer for each vertex
         For j = 0 To 2
            
            With tVerts(i * 3 + j)
            
               ' ## get texture coordinates for current triangle
               .t.x = SkinCoords(Triangles(i).tindex(j)).s / hMDL.skinwidth
               .t.y = SkinCoords(Triangles(i).tindex(j)).t / hMDL.skinheight

               ' ## set precalculated normal for this vector
               .n = Frames(cFrame).fdata(Triangles(i).vindex(j)).n
               
               ' ## get vertex data
               .p = Frames(cFrame).fdata(Triangles(i).vindex(j)).p
                             
               ' ## if interpolation is enabled then make an interpolation between 2 consecutive frames
               If cIP Then
                  D3DXVec3Lerp .p, .p, Frames(nFrame).fdata(Triangles(i).vindex(j)).p, cIPv
               End If
          
            End With
            
         Next j
         
      Next i
   
      ' ## copy the buffer into a DX vertex buffer
      D3DVertexBuffer8SetData cVBuffer, 0, Len(tVerts(0)) * (UBound(tVerts) + 1), 0, tVerts(0)
      ' ## let DX know we are going to render the vertex buffer
      .SetStreamSource 0, cVBuffer, Len(tVerts(0))
      ' ## draw the triangles
      .DrawPrimitive D3DPT_TRIANGLELIST, 0, hMDL.num_tris
   
      End With
   
   End If
       
End Sub

' ------------------------------------------------- Class properties ----------------------------------------------------

' ## set device where rendering takes place
Sub SetDevice(dev As Direct3DDevice8)
   Set c3DDevice = dev
End Sub

' ## get device where rendering takes place
Function GetDevice() As Direct3DDevice8
   Set GetDevice = c3DDevice
End Function

' ## set material
Sub SetMaterial(Mt As D3DMATERIAL8)
   cMaterial = Mt
End Sub

' ## get material
Function GetMaterial() As D3DMATERIAL8
   GetMaterial = cMaterial
End Sub

' ## set current skin
Sub SetSkin(Nr As Long)
   If Nr < 0 Then Nr = 0
   If Nr > UBound(SkinNames) Then Nr = UBound(SkinNames)
   cSkin = Nr
End Sub

' ## get current skin
Function GetSkin() As Long
   GetSkin = cSkin
End Function

' ## get number os skins available
Function GetCountSkins() As Long
   If cModelLoaded Then
      GetCountSkins = UBound(SkinNames) + 1
   End If
End Function

' ## set current frame
Sub SetFrame(Nr As Long)
   If Nr < 0 Then Nr = 0
   If Nr > hMDL.num_frames - 1 Then Nr = hMDL.num_frames - 1
   cFrame = Nr
End Sub

' ## get current frame
Function GetFrame() As Long
   GetFrame = cFrame
End Function

' ## advance to the next frame in the animation
' ## and take care of animation boundaries
Sub NextFrame()
   
   If cModelLoaded Then
      
      With FramesIndex(cAnim)
      
         ' ## next frame
         cFrame = cFrame + 1
         If cFrame > .last Then
            cFrame = .first
         End If
         
         ' ## obtain next frame
         If cFrame < .last Then
            nFrame = cFrame + 1
         Else
            nFrame = .first
         End If
         
         ' ## just to be sure we remain inside the frame boundaries
         If nFrame > hMDL.num_frames - 1 Then
            nFrame = .first
         End If
      
      End With
      
      ' ## reset interpolation value just to be sure it's 0 when we start a new frame
      ' ## else we get some wrong model animation
      cIPv = 0
      
   End If
      
End Sub

' ## set current animation
Sub SetAnimation(Nr As Long)
   If cModelLoaded Then
     If Nr < 0 Then Nr = 0
     If Nr > UBound(FramesIndex) Then Nr = UBound(FramesIndex)
     ' ## set the current frame to the first frame of the animation
     cFrame = FramesIndex(Nr).first
     ' ## just to be sure not running into unallocated space
     If cFrame + 1 <= hMDL.num_frames - 1 Then
        ' ## set the next frames index (for frame interpolation purposes)
        nFrame = cFrame + 1
     End If
     cIPv = 0
     cAnim = Nr
   End If
End Sub

' ## get current animation
Function GetAnimation() As Long
   GetAnimation = cAnim
End Function

' ## returns the number of animation sequences in the model
Function GetCountAnimations() As Long
   If cModelLoaded Then
      GetCountAnimations = UBound(FramesIndex) + 1
   End If
End Function

' ## returns the animation's name
Function GetAnimationName(Nr As Long) As String
   If cModelLoaded Then
      If Nr < 0 Then Nr = 0
      If Nr > UBound(FramesIndex) Then Nr = UBound(FramesIndex)
      GetAnimationName = Trim(FramesIndex(Nr).fname)
   End If
End Function

' ## enable/disable frame interpolation
Sub EnableFrameIP(v As Boolean)
   cIP = v
End Sub

' ## get enable/disable frame interpolation
Function GetEnableFrameIP() As Boolean
   GetEnableFrameIP = cIP
End Function

' ## set the interpolation value to be used while rendering
Sub SetFrameIP(v As Single)
   If v < 0 Then v = 0
   If v > 1 Then v = 1
   cIPv = v
End Sub

' ## get the interpolation value used while rendering
Function GetFrameIP() As Single
   GetFrameIP = cIPv
End Function

' ## enable/disable texture
Sub EnableTexture(v As Boolean)
   cTextured = v
End Sub

' ## enable/disable texture
Function GetEnableTexture() As Boolean
   GetEnableTexture = cTextured
End Function

' ## add a custom skin (for non quakeII models)
' ## this call will go *before* the loading of the model
' ## custom skins *must* be in the same directory as the loaded model
Sub AddCustomTexture(File As String)
   If Not cModelLoaded Then
      ReDim Preserve SkinNames(cCustSkin)
      SkinNames(cCustSkin).sname = File & vbNullChar
      cCustSkin = cCustSkin + 1
   End If
End Sub

' ## free resources
Sub Unload()
   FreeResources
   ReDim SkinNames(0)
   Set c3DDevice = Nothing
   cModelLoaded = False
End Sub

' ## error
Function GetLastError() As MDL_ERRORS
   GetLastError = cErr
   cErr = MDL_OK
End Function

Sub FreeResources()
   Dim i As Long
   On Error Resume Next
   For i = 0 To UBound(cTexture)
     If Not cTexture(i) Is Nothing Then Set cTexture(i) = Nothing
   Next i
   ReDim cTexture(0)
   ReDim SkinCoords(0)
   ReDim Triangles(0)
   ReDim AliasData.adata(0)
   ReDim Frames(0)
   ReDim FaceIndex(0)
   Set cVBuffer = Nothing
End Sub

' ## initialize some things
Private Sub Class_Initialize()
   cSkin = 0
   cFrame = 0
   nFrame = 0
   cAnim = 0
   cIP = False
   cTextured = False
   cModelLoaded = False
   cPos.x = 0
   cPos.y = 0
   cPos.z = 0
   cErr = MDL_OK
End Sub

' ## cleanup memory resources
Private Sub Class_Terminate()
   Unload
End Sub


