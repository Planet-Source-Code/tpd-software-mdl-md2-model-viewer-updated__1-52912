VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm3D 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDL / MD2 Viewer"
   ClientHeight    =   9360
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSkin 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox cmbAnim 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   2775
   End
   Begin VB.PictureBox Screen 
      BorderStyle     =   0  'None
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   330
      Width           =   12000
   End
   Begin MSComDlg.CommonDialog MDLLoad 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lnkModels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "download model files by clicking here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9240
      TabIndex        =   5
      Top             =   45
      Width           =   2625
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Skins"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   3
      Top             =   60
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Animation"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   60
      Width           =   855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load model"
      End
      Begin VB.Menu S1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuLight 
         Caption         =   "Enable &light"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuIpFrame 
         Caption         =   "Frame &interpolation"
         Checked         =   -1  'True
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRenderMode 
         Caption         =   "&Wireframe"
         Index           =   0
      End
      Begin VB.Menu mnuRenderMode 
         Caption         =   "&Flat shaded"
         Index           =   1
      End
      Begin VB.Menu mnuRenderMode 
         Caption         =   "&Gouraud shaded"
         Index           =   2
      End
      Begin VB.Menu mnuRenderMode 
         Caption         =   "&Textured"
         Checked         =   -1  'True
         Index           =   3
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuAboutViewer 
         Caption         =   "&MDL Viewer"
      End
   End
End
Attribute VB_Name = "frm3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAnim_Click()
   ' ## set a different animation sequence
   Select Case MDLType
   Case 1
     cMDL.SetAnimation cmbAnim.ListIndex
   Case 2
     cMD2.SetAnimation cmbAnim.ListIndex
   End Select
   Screen.SetFocus
End Sub

Private Sub Command1_Click()
   LoadPCX "E:\VBSource\MDLVIE~2\MD2\skin_lt.pcx"
End Sub

Private Sub cmbSkin_Click()
   Select Case MDLType
   Case 1
     cMDL.SetSkin cmbSkin.ListIndex
   Case 2
     cMD2.SetSkin cmbSkin.ListIndex
   End Select
   Screen.SetFocus
End Sub

Private Sub Form_Load()
  
  Dim Material As D3DMATERIAL8
  
  ' ## initialize load dialog
  With MDLLoad
     .DialogTitle = "Select model file"
     .CancelError = True
     .DefaultExt = ".mdl"
     .Filter = "MDL/MD2 files (.mdl/.md2)|*.mdl;*.md2"
     .InitDir = App.Path
     .flags = 4
  End With
  
  ' ## initialize directx
  Me.Show
  If Not DXLaunch(Screen) Then
     MsgBox "Could not initialize directx."
     End
  End If
  
  ' ## initialize MDL object
  Set cMDL = New cls_MDL
  
  ' ## initialize MD2 object
  Set cMD2 = New cls_MD2
  
  ' ## set the device to render to for this model object
  cMDL.SetDevice D3DDevice
  cMD2.SetDevice D3DDevice
  
  ' ## make some material for the model
  With Material
     .Ambient = CreateD3DColorVal(1, 1, 1, 0)
     .diffuse = .Ambient
  End With
   
  With cMDL
     ' ## enable frame interpolation
     .EnableFrameIP True
     ' ## enable model texture
     .EnableTexture True
     ' ##  apply material
     .SetMaterial Material
     ' ## load palette for this model
     .LoadPalette App.Path & "\palette.lmp"
  End With
  
  With cMD2
     ' ## enable frame interpolation
     .EnableFrameIP True
     ' ## enable model texture
     .EnableTexture True
     ' ##  apply material
     .SetMaterial Material
  End With
  
  ' ## load precalculated normals
  LoadNormals
  ' ## create font for DrawText
  CreateFont
  ' ## set a light source
  SetLight 0, D3DLIGHT_DIRECTIONAL, MakeVector(0, 0, 400), MakeVector(0, 0, -1), 500, CreateD3DColorVal(1, 1, 1, 1)
  ' ## position the camera
  SetCamera MakeVector(0, 0, 400), MakeVector(0, 0, 0), MakeVector(0, 1, 0)
  ' ## enter the main loop
  MainLoop
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ' ## end objects and DX
  cMDL.Unload
  cMD2.Unload
  DXEnd
  End
End Sub

Private Sub lnkModels_Click()
  ShellExecute Me.hwnd, vbNullString, "http://www.planetquake.com/polycount/downloads/", vbNullString, vbNullString, vbMaximized
End Sub

' ## MENU->about the author
Private Sub mnuAboutViewer_Click()
   MsgBox "MDL / MD2 (Quake 1 & 2) Model Viewer by TPD Software.", vbOKOnly Or vbInformation, "About"
End Sub

' ## MENU->exit program
Private Sub mnuExit_Click()
  Form_Unload 0
End Sub

' ## MENU->enable/disable frame interpolation
Private Sub mnuIpFrame_Click()
  With mnuIpFrame
    .Checked = Not .Checked
    Select Case MDLType
    Case 1
      cMDL.EnableFrameIP .Checked
    Case 2
      cMD2.EnableFrameIP .Checked
    End Select
  End With
End Sub

' ## MENU->enable/disable light
Private Sub mnuLight_Click()
   Dim Flag As Long
   With mnuLight
     .Checked = Not .Checked
     If .Checked Then
        Flag = 1
     Else
        Flag = 0
     End If
   End With
   If Not D3DDevice Is Nothing Then D3DDevice.SetRenderState D3DRS_LIGHTING, Flag
End Sub

' ## MENU->load a model
Private Sub mnuLoad_Click()
   On Error GoTo Cancelled
   MDLLoad.ShowOpen
   
   Dim i   As Long
   Dim ret As Boolean
   
   ' ## file is MDL
   If InStr(1, MDLLoad.FileName, ".mdl", vbTextCompare) Then
      
      ' ## this so we know we are rendering MDL's
      MDLType = 1
      
      cMDL.Load (MDLLoad.FileName)
      
      If cMDL.GetLastError = MDL_OK Then
   
         ' ## fill a combobox with the animation names
         cmbAnim.Clear
         For i = 0 To cMDL.GetCountAnimations - 1
            cmbAnim.AddItem cMDL.GetAnimationName(i)
         Next i
   
         ' ## show number of skins in the model
         cmbSkin.Clear
         For i = 0 To cMDL.GetCountSkins - 1
            cmbSkin.AddItem "Skin " & i
         Next i
         
      Else
         MsgBox MDLLoad.FileTitle & " could not be loaded properly."
      End If
   
   End If
      
   ' ## file is MD2
   If InStr(1, MDLLoad.FileName, ".md2", vbTextCompare) Then
      
      ' ## this so we know we are rendering MD2's
      MDLType = 2
      
      cMD2.Load MDLLoad.FileName
      
      If cMD2.GetLastError = MDL_OK Then
   
         ' ## fill a combobox with the animation names
         cmbAnim.Clear
         For i = 0 To cMD2.GetCountAnimations - 1
            cmbAnim.AddItem cMD2.GetAnimationName(i)
         Next i
   
         ' ## show number of skins in the model
         cmbSkin.Clear
         For i = 0 To cMD2.GetCountSkins - 1
            cmbSkin.AddItem "Skin " & i
         Next i
   
      Else
         MsgBox MDLLoad.FileTitle & " could not be loaded properly."
      End If
   
   End If
   
   ' ## reset model's view matrix
   InitializeMatrix
      
   ' ## set first animation
   If cmbAnim.ListCount > 0 Then
      cmbAnim.ListIndex = 0
   End If
   
   ' ## set first skin
   If cmbSkin.ListCount > 0 Then
      cmbSkin.ListIndex = 0
   End If
   
Cancelled:
End Sub

' ## MENU->set render mode (wireframe, flat shaded, gouraud shaded and textured)
Private Sub mnuRenderMode_Click(Index As Integer)
   For i = 0 To mnuRenderMode.Count - 1
      If i = Index Then
         mnuRenderMode(i).Checked = True
         
         If Not D3DDevice Is Nothing Then
           
           If i < 3 Then
              Select Case MDLType
              Case 1
                cMDL.EnableTexture False
              Case 2
                cMD2.EnableTexture False
              End Select
           End If
           
           If i = 0 Then
             D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_WIREFRAME
           Else
             D3DDevice.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
           End If
           
           If i = 1 Then
             D3DDevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_FLAT
           ElseIf i = 2 Then
             D3DDevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
           End If
           
           If i = 3 Then
             D3DDevice.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
             Select Case MDLType
             Case 1
               cMDL.EnableTexture mnuRenderMode(i).Checked
             Case 2
               cMD2.EnableTexture mnuRenderMode(i).Checked
             End Select
           End If
           
         End If
         
      Else
         mnuRenderMode(i).Checked = False
      End If
   Next i
End Sub
