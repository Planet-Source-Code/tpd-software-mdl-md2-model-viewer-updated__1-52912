VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_MDL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ####################################################################################################################
' #
' #  MDL Model class - Copyright (c) TPD Software
' #
' #     This class handles all MDL related stuff
' #
' #     Features:   - multiple skin support
' #                 - uses precalculated normals
' #                 - animation with frame interpolation
' #                 - vertex buffering
' #
' ####################################################################################################################

Option Explicit

' --------------- DX related ---------------

Private Const FVF_MDLVERTEX As Long = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Private cVBuffer            As Direct3DVertexBuffer8
Private cTexture()          As Direct3DTexture8
Private cMaterial           As D3DMATERIAL8
Private c3DDevice           As Direct3DDevice8

' --------------- MDL related --------------

Private Type tHeader
    id           As String * 4
    version      As Long
    scale        As D3DVECTOR
    scaleorg     As D3DVECTOR
    boundrad     As Single
    eyepos       As D3DVECTOR
    numskins     As Long
    skinwidth    As Long
    skinheight   As Long
    numvertices  As Long
    numtriangles As Long
    numframes    As Long
    synctype     As Long
    flags        As Long
    size         As Single
End Type

Private Type tSkin
    group        As Long
    Skin()       As Byte
End Type

Private Type tVertex2D
    onseam       As Long                      ' 0 or 32 (if point is on edge)
    x            As Long
    y            As Long
End Type

Private Type tVertex3D
    facesfront   As Long
    vertindex(2) As Long
End Type

Private Type tVertexIndexes
    x            As Byte
    y            As Byte
    z            As Byte
    iLightNormal As Byte
End Type

Private Type tTriangleVertex
    p            As D3DVECTOR
    n            As D3DVECTOR
    t            As D3DVECTOR2
End Type
    
Private Type tAliasFrame
    ftype        As Long
    min          As tVertexIndexes
    max          As tVertexIndexes
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
Private SkinBuf()       As tSkin
Private Vertices2D()    As tVertex2D
Private Vertices3D()    As tVertex3D
Private AliasFrame      As tAliasFrame
Private AliasData       As tAliasFrameData
Private Frames()        As tFrame
Private FramesIndex()   As tFrameIndex

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

   On Error GoTo LoadError
   
   Dim i As Long
   Dim j As Long

   If c3DDevice Is Nothing Then Exit Sub
   
   ' ## clear resources
   FreeResources
   
   Open File For Binary As #1
   
     cModelLoaded = True
   
     ' ## get MDL header
     Get #1, , hMDL
     If hMDL.id <> "IDPO" Then
        Close #1
        cErr = MDL_INVALID_ID
        Exit Sub
     End If
     
     If hMDL.version <> 6 Then
        Close #1
        cErr = MDL_INVALID_VERSION
        Exit Sub
     End If
          
     ' ## allocate memory for model data
     ReDim cTexture(hMDL.numskins - 1)
     ReDim SkinBuf(hMDL.numskins - 1)
     ReDim Vertices2D(hMDL.numvertices - 1)
     ReDim Vertices3D(hMDL.numtriangles - 1)
     ReDim AliasData.adata(hMDL.numvertices - 1)
     ReDim Frames(hMDL.numframes - 1)
     
     ' ## read skin texture
     For i = 0 To hMDL.numskins - 1
        ReDim SkinBuf(i).Skin(hMDL.skinwidth * hMDL.skinheight - 1)
        Get #1, , SkinBuf(i).group
        Get #1, , SkinBuf(i).Skin()
     Next i
     
     ' ## create an X8R8G8B8 bitmap from the model's 8-bit texture
     For i = 0 To hMDL.numskins - 1
       Set cTexture(i) = CreateRGBTextureFromPalettizedBitmap(SkinBuf(i).Skin())
     Next i
          
     ' ## read skin vertices
     Get #1, , Vertices2D()
     ' ## read model triangles
     Get #1, , Vertices3D()
     
     ' ## read model frames and apply precalculated normals
     ' ## also change the models coordinate system to directx's coordinate system
     For i = 0 To hMDL.numframes - 1
        Get #1, , AliasFrame
        Get #1, , AliasData.adata()
        ReDim Frames(i).fdata(hMDL.numvertices - 1)
        Frames(i).fname = AliasFrame.aname
        For j = 0 To hMDL.numvertices - 1
           Frames(i).fdata(j).p.x = hMDL.scale.y * AliasData.adata(j).y + hMDL.scaleorg.y
           Frames(i).fdata(j).p.y = hMDL.scale.z * AliasData.adata(j).z + hMDL.scaleorg.z
           Frames(i).fdata(j).p.z = hMDL.scale.x * AliasData.adata(j).x + hMDL.scaleorg.x
           Frames(i).fdata(j).n.x = Normals(AliasData.adata(j).iLightNormal, 1)
           Frames(i).fdata(j).n.y = Normals(AliasData.adata(j).iLightNormal, 2)
           Frames(i).fdata(j).n.z = Normals(AliasData.adata(j).iLightNormal, 0)
        Next j
     Next i
     
     ' ## create a lookup buffer for the animations start and end indexes
     GetAnimationIndexes
     
     ' ## allocate enough memory for the vertex buffer
     ' ## (structure length * num triangles * num vertices per triangle (3))
     Set cVBuffer = c3DDevice.CreateVertexBuffer(Len(Frames(0).fdata(0)) * hMDL.numtriangles * 3, 0, FVF_MDLVERTEX, D3DPOOL_MANAGED)
     
     ' ## clear the old skinbuffer as we don't furhter need it
     ReDim SkinBuf(0)
     
   Close #1
   
   ' ## reset index variables as they change when loading a new model
   cSkin = 0
   cFrame = 0
   nFrame = 0
   cAnim = 0
   cErr = MDL_OK
   
   Exit Sub

LoadError:

   Close #1
   cModelLoaded = False
   cErr = MDL_LOAD_ERROR
     
End Sub

' ## Make a lookup buffer for all animations in the model
Private Sub GetAnimationIndexes()
  
   Dim an As String
   Dim la As String
   Dim ln As Long
   Dim fi As Long
   Dim i  As Long
   
   ' ## loop trough al frames
   For i = 0 To hMDL.numframes - 1
      ' ## strip the name
      an = Left(Frames(i).fname, InStr(1, Frames(i).fname, vbNullChar) - 1)
      ln = Len(an)
      ' ## check for asc extension
      If Right(an, 4) = ".asc" Then ln = ln - 4
      ' ## check for letter extension
      If Mid(an, ln, 1) >= "a" And Mid(an, ln, 1) <= "z" Then ln = ln - 1
      ln = ln - 1
      Do While Val(Mid(an, ln, 1)) > 0 And Val(Mid(an, ln, 1)) <= 9
         ln = ln - 1
      Loop
      ln = ln + 1
      an = Left(an, ln - 1)
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
   FramesIndex(UBound(FramesIndex)).last = hMDL.numframes - 1
   
End Sub

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
      
      ReDim tVerts(hMDL.numtriangles * 3 - 1) As tTriangleVertex  ' ## model triangle buffer
      Dim i                                   As Long
      Dim j                                   As Long
      
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
      For i = 0 To hMDL.numtriangles - 1
         
         ' ## fill the buffer for each vertex
         For j = 0 To 2
            
            With tVerts(i * 3 + j)
            
               ' ## get texture coordinates for current triangle
               .t.x = Vertices2D(Vertices3D(i).vertindex(j)).x / hMDL.skinwidth
               .t.y = Vertices2D(Vertices3D(i).vertindex(j)).y / hMDL.skinheight
            
               ' ## adjust texture coordinates
               If Vertices2D(Vertices3D(i).vertindex(j)).onseam And Vertices3D(i).facesfront = 0 Then
                  .t.x = .t.x + 0.5
               End If

               ' ## set precalculated normal for this vector
               .n = Frames(cFrame).fdata(Vertices3D(i).vertindex(j)).n
            
               ' ## get vertex data
               .p = Frames(cFrame).fdata(Vertices3D(i).vertindex(j)).p
                 
               ' ## if interpolation is enabled then make an interpolation between 2 consecutive frames
               If cIP Then
                  D3DXVec3Lerp .p, .p, Frames(nFrame).fdata(Vertices3D(i).vertindex(j)).p, cIPv
               End If
          
            End With
         
         Next j
         
      Next i
   
      ' ## copy the buffer into a DX vertex buffer
      D3DVertexBuffer8SetData cVBuffer, 0, Len(tVerts(0)) * (UBound(tVerts) + 1), 0, tVerts(0)
      ' ## let DX know we are going to render the vertex buffer
      .SetStreamSource 0, cVBuffer, Len(tVerts(0))
      ' ## draw the triangles
      .DrawPrimitive D3DPT_TRIANGLELIST, 0, hMDL.numtriangles
   
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
   If Nr > hMDL.numskins - 1 Then Nr = hMDL.numskins - 1
   cSkin = Nr
End Sub

' ## get current skin
Function GetSkin() As Long
   GetSkin = cSkin
End Function

' ## get number os skins available
Function GetCountSkins() As Long
   If cModelLoaded Then
      GetCountSkins = hMDL.numskins
   End If
End Function

' ## set current frame
Sub SetFrame(Nr As Long)
   If Nr < 0 Then Nr = 0
   If Nr > hMDL.numframes - 1 Then Nr = hMDL.numframes - 1
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
         If nFrame > hMDL.numframes - 1 Then
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
     If cFrame + 1 <= hMDL.numframes - 1 Then
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

' ## free resources
Sub Unload()
   FreeResources
   Set c3DDevice = Nothing
   cModelLoaded = False
End Sub

' ## Load palette as mdl bitmaps are 8 bit, we need the corresponding RGB values
Sub LoadPalette(File As String)
    Open File For Binary As #1
      Get #1, , Pal()
    Close #1
End Sub

' ## error
Function GetLastError() As MDL_ERRORS
   GetLastError = cErr
   cErr = MDL_OK
End Function

Private Sub FreeResources()
   Dim i As Long
   On Error Resume Next
   For i = 0 To UBound(cTexture)
     If Not cTexture(i) Is Nothing Then Set cTexture(i) = Nothing
   Next i
   ReDim cTexture(0)
   ReDim Vertices2D(0)
   ReDim Vertices3D(0)
   ReDim AliasData.adata(0)
   ReDim Frames(0)
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
