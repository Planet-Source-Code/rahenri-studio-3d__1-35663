Attribute VB_Name = "RenderEngine"

Dim X As Long
Dim Y As Long
Dim Hit As PickInfo
Dim tHit As PickInfo
Dim World As D3DMATRIX
Dim View As D3DMATRIX
Dim Proj As D3DMATRIX
Dim T As D3DMATRIX, T2 As D3DMATRIX, T3 As D3DMATRIX
Dim P As Long
Dim N As D3DVECTOR
Dim N1 As Vector, N2 As Vector, N3 As Vector, N4 As Vector, N5 As Vector
Dim Rect() As Vector
Dim LightSource As D3DVECTOR
Dim LS As D3DVECTOR
Dim Normalizer As Single
Dim Shade As Double
Dim Shade1 As Double
Dim T1 As Long
Dim V As D3DVECTOR
Dim VP As D3DVIEWPORT8
Dim VecPos As D3DVECTOR
Dim VecDir As D3DVECTOR
Dim Vec1 As Vector
Dim Obj() As Object3D
Dim tMat() As D3DMATRIX
Dim TV1 As Vector, TV2 As Vector, TV3 As Vector


Public Function RenderScene(Window As Object, Bar As ProgressBar)

Window.Cls
Window.BackColor = vbBlack
If Scene.GetnObj = -1 Then Bar.Value = Bar.Max: Exit Function
Scene.GetObjects Obj
ReDim tMat(UBound(Obj))


For T1 = 0 To UBound(Obj)
With Obj(T1)

D3DXMatrixAffineTransformation tMat(T1), 1, cD3DVECTOR(.Center), Rota2(.Rotation), cD3DVECTOR(.Position)
End With
Next

ReDim Rect(UBound(Obj), 1)
Bar.Max = Window.Height




  LightSource.X = 0
  LightSource.Y = 0
  LightSource.Z = 0


Dim K2 As D3DVECTOR4

D3dDevice.GetTransform D3DTS_VIEW, View
D3dDevice.GetTransform D3DTS_WORLD, World
D3dDevice.GetTransform D3DTS_PROJECTION, Proj
D3dDevice.GetViewport VP



For T1 = 0 To UBound(Obj)
Rect(T1, 0).X = 2147483647
Rect(T1, 0).Y = 2147483647
Rect(T1, 0).Z = 2147483647

Rect(T1, 1).X = -2147483647
Rect(T1, 1).Y = -2147483647
Rect(T1, 1).Z = -2147483647



For P = 0 To UBound(Obj(T1).VertexBuffer.Vertex)
V = cD3DVECTOR(Obj(T1).VertexBuffer.Vertex(P).Position)

D3DXVec3Project V, V, VP, Proj, View, tMat(T1)

If V.X < Rect(T1, 0).X Then Rect(T1, 0).X = V.X
If V.Y < Rect(T1, 0).Y Then Rect(T1, 0).Y = V.Y

If V.X > Rect(T1, 1).X Then Rect(T1, 1).X = V.X
If V.Y > Rect(T1, 1).Y Then Rect(T1, 1).Y = V.Y

Next
Next



For Y = 0 To Window.Height
For X = 0 To Window.Width
Hit.HitDist = 2147483647
Hit.ObjHit = -1
For T1 = 0 To UBound(Obj)

If X >= Rect(T1, 0).X And X <= Rect(T1, 1).X And Y >= Rect(T1, 0).Y And Y <= Rect(T1, 1).Y Then
tHit = Scene.RayIntersect(CSng(X), CSng(Y), T1)
If tHit.ObjHit = 1 And tHit.HitDist < Hit.HitDist Then

Hit = tHit
Hit.ObjHit = T1

End If
End If
Next

If Hit.ObjHit > -1 Then


D3DXMatrixMultiply T, World, tMat(Hit.ObjHit)
D3DXMatrixMultiply T, T, View



N1 = Obj(Hit.ObjHit).VertexBuffer.Vertex(Obj(Hit.ObjHit).VertexBuffer.Indices(Hit.FaceIndex * 3)).Normal
N2 = Obj(Hit.ObjHit).VertexBuffer.Vertex(Obj(Hit.ObjHit).VertexBuffer.Indices((Hit.FaceIndex * 3) + 1)).Normal
N3 = Obj(Hit.ObjHit).VertexBuffer.Vertex(Obj(Hit.ObjHit).VertexBuffer.Indices((Hit.FaceIndex * 3) + 2)).Normal

TV1 = Obj(Hit.ObjHit).VertexBuffer.Vertex(Obj(Hit.ObjHit).VertexBuffer.Indices(Hit.FaceIndex * 3)).Position
TV2 = Obj(Hit.ObjHit).VertexBuffer.Vertex(Obj(Hit.ObjHit).VertexBuffer.Indices((Hit.FaceIndex * 3) + 1)).Position
TV3 = Obj(Hit.ObjHit).VertexBuffer.Vertex(Obj(Hit.ObjHit).VertexBuffer.Indices((Hit.FaceIndex * 3) + 2)).Position


MixNormals N4, N1, N2, N3, 1 - (Hit.u + Hit.V), Hit.u, Hit.V

N = cD3DVECTOR(N4)
      
      D3DXVec3TransformNormal N, N, T
      
        'compute vertex pos
        
        'Compute super smoothing
        
        Dim SP1 As Single
        

        N4 = Add(N1, N2)
        N4 = Add(N4, N3)
        
        N4.X = N4.X / 2
        N4.Y = N4.Y / 2
        N4.Z = N4.Z / 2
        
        SP1 = D3Dist(TV1, TV2) + D3Dist(TV2, TV3) + D3Dist(TV3, TV1)
        
        SP1 = SP1 / 3
      
        
      
        D3DXVec3TransformCoord VecPos, cD3DVECTOR(Hit.CheckPos), T
        D3DXVec3TransformNormal VecDir, cD3DVECTOR(Hit.CheckDir), T
        D3DXVec3Add VecPos, VecPos, Vec(VecDir.X * Hit.HitDist, VecDir.Y * Hit.HitDist, VecDir.Z * Hit.HitDist)
        
        
        
        D3DXVec3Add VecPos, VecPos, Vec(N.X * SP1, N.Y * SP1, N.Z * SP1)
        
        'compute direction from vertex to light
        
        D3DXVec3Subtract VecPos, LightSource, VecPos
        
        D3DXVec3Normalize VecPos, VecPos
        
        Shade = ((N.X * VecPos.X) + (N.Y * VecPos.Y) + (N.Z * VecPos.Z))
        

        
If Shade < 0 Then Shade = 0

If Shade > 1 Then Shade = 1


P = 255 * Shade


P = RGB(P * Obj(Hit.ObjHit).Material.Diffuse.R, P * Obj(Hit.ObjHit).Material.Diffuse.G, P * Obj(Hit.ObjHit).Material.Diffuse.B)


SetPixel Window.hdc, X, Y, P

End If

Next

DoEvents
Bar.Value = Y
Next

End Function


Private Function D3Dist(V1 As Vector, V2 As Vector) As Single
Dim C1 As Single, C2 As Single, C3 As Single
C1 = Abs(V1.Y - V2.Y)
C2 = Abs(V1.X - V2.X)
C3 = Abs(V1.Z - V2.Z)
C2 = Sqr(Q(C2) + Q(C3))
D3Dist = Sqr(Q(C1) + Q(C2))
End Function

Private Function Q(N As Single)
Q = N * N
End Function
