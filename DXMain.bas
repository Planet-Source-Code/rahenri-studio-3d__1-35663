Attribute VB_Name = "DXMain"
Option Explicit
Public Declare Function GetTickCount Lib "kernel32" () As Long



Public DX As DirectX8
Public D3D As Direct3D8
Public D3dDevice As Direct3DDevice8
Public vb As Direct3DVertexBuffer8
Public D3DX As D3DX8
Public D3DCaps As D3DCAPS8
Public D3Dpp As D3DPRESENT_PARAMETERS
Public mode As D3DDISPLAYMODE


Function InitD3D(hWnd As Long, Optional devtype As CONST_D3DDEVTYPE = D3DDEVTYPE_HAL) As Boolean
    On Error GoTo erro
    
    
    
    Set DX = New DirectX8
    Set D3D = DX.Direct3DCreate
    Set D3DX = New D3DX8
    
    
    
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
    
        
    D3Dpp.Windowed = True
    D3Dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    D3Dpp.BackBufferFormat = mode.Format
    D3Dpp.BackBufferCount = 1
    D3Dpp.EnableAutoDepthStencil = 1
    D3Dpp.AutoDepthStencilFormat = D3DFMT_D16
    Set D3dDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, devtype, hWnd, _
                                      D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3Dpp)
    If D3dDevice Is Nothing Then GoTo erro
    
    'd3d effect modes
    
    With D3dDevice
    .SetRenderState D3DRS_CULLMODE, 1
    '.SetRenderState D3DRS_ZENABLE, 1
    '.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    '.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    '.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    '.SetRenderState D3DRS_SPECULARENABLE, 1
    .SetRenderState D3DRS_DITHERENABLE, 1
    '.SetRenderState D3DRS_AMBIENT, &HFFAAAAAA
    .SetRenderState D3DRS_LIGHTING, 1
    End With
    


    InitD3D = True
  
  

  
Exit Function
erro:
If devtype <> D3DDEVTYPE_REF Then
MsgBox "O modo aceleração 3d não pode ser iniciado, no entanto programa poderá funcionara no modo software."
InitD3D = InitD3D(hWnd, D3DDEVTYPE_REF)
End If
End Function






Sub Cleanup()
    Set vb = Nothing
    Set D3dDevice = Nothing
    Set D3D = Nothing
End Sub


Sub Render()


    Dim V As GlobalVertex
    Dim sizeOfVertex As Long
    
    
    If D3dDevice Is Nothing Then Exit Sub

    D3dDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &HFF0000FF, 1#, 0
    

    D3dDevice.BeginScene
    

    'render grid
    Dim Mtrl As D3DMATERIAL8
    Mtrl.emissive.R = 0.7
    Mtrl.emissive.G = 0.7
    Mtrl.emissive.B = 0.7
    Mtrl.emissive.A = 0.7
    D3dDevice.SetMaterial Mtrl
    D3dDevice.SetVertexShader modMain.GrdShader
    D3dDevice.DrawPrimitiveUP D3DPT_LINELIST, 120, Grid(0), Len(Grid(0))

    Scene.RenderObjects
             
    D3dDevice.EndScene
    
     

    D3dDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
HE:
End Sub










Function CreateColor(Red As Single, Green As Single, Blue As Single, Alfa As Single) As D3DCOLORVALUE
With CreateColor
.R = Red
.G = Green
.B = Blue
.A = Alfa
End With
End Function

Public Function ColorValue4(A As Single, R As Single, G As Single, B As Single) As D3DCOLORVALUE
    Dim C As D3DCOLORVALUE
    C.A = A
    C.R = R
    C.G = G
    C.B = B
    ColorValue4 = C
End Function



