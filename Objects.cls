VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Objects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"





Private Type VBB
Pos As D3DVECTOR
Diffuse As Long
End Type



Private Type D3DStuff
Transformation As D3DMATRIX
RotQ As D3DQUATERNION
Mesh As D3DXMesh
End Type

Dim Obj() As Object3D
Dim nObj As Long

Dim Lights() As tLight
Dim nLights As Long

Dim Cams() As tCamera
Dim nCams As Long

Dim WithEvents View As Opt
Attribute View.VB_VarHelpID = -1



Dim Stuff() As D3DStuff
Dim DX As New D3DX8
Dim Dev As Direct3DDevice8
Dim BB(23) As VBB
Dim RenderBB As Boolean
Dim BBObj As Long
Dim ActObj As Long
Dim ActiveId As Long
Private Const BBShader = D3DFVF_XYZ Or D3DFVF_DIFFUSE

Function AddCam(Cam As tCamera) As Long
nCams = nCams + 1
ReDim Preserve Cams(nCams)
Cams(nCams) = Cam
AddCam = nCams
End Function

Function GetCam(CamId As Long) As tCamera
GetCam = Cams(CamId)
End Function

Function SetCam(CamId As Long, Cam As tCamera)
Cams(CamId) = Cam
End Function

Function ActiveCam(CamId As Long)
Dim Mat As D3DMATRIX
With Cams(CamId)
D3DXMatrixLookAtLH Mat, cD3DVECTOR(.Position), cD3DVECTOR(.Direction), cD3DVECTOR(.Orientation)
D3dDevice.SetTransform D3DTS_VIEW, Mat
D3DXMatrixPerspectiveFovLH Mat, .Fov, .Apect, 1, 10000
D3dDevice.SetTransform D3DTS_PROJECTION, Mat
ActiveId = CamId
End With
End Function

Function AddObj(Creator As Object) As Long
Randomize Timer
Dim T As Long




T = Creator.TypeId



Select Case T
Case tObject
nObj = nObj + 1
ReDim Preserve Obj(nObj)
ReDim Preserve Stuff(nObj)

D3DXQuaternionIdentity Stuff(nObj).RotQ
D3DXMatrixIdentity Stuff(nObj).Transformation
Set Obj(nObj).Creator = Creator
Obj(nObj).Creator.SetObjects Me, nObj
Obj(nObj).Creator.CreateObject

Obj(nObj).Material.Diffuse = Col(1, 1, 1, 1)


AddObj = nObj

Case tLight



Case tCam



Case tSystem



End Select


End Function

Private Sub Class_Initialize()
nObj = -1
nCams = -1
nLights = -1
End Sub


Function SetToCreator(ObjId As Long)

End Function

Function SetToModeler(ObjId) As Long

End Function

Function ApplyModifier(ObjId As Long)

End Function

Sub CallUpdateObj(ObjId As Long)

Dim T As Long
Erase Obj(ObjId).VertexBuffer.Vertex
Obj(ObjId).Creator.RequestBuffer Obj(ObjId).VertexBuffer
Set Stuff(ObjId).Mesh = Nothing
Set Stuff(ObjId).Mesh = DX.CreateMeshFVF((UBound(Obj(ObjId).VertexBuffer.Indices) + 1) / 3, UBound(Obj(ObjId).VertexBuffer.Vertex) + 1, D3DXMESH_MANAGED, D3DFVF, Dev)

If Not D3DXMeshVertexBuffer8SetData(Stuff(ObjId).Mesh, 0, Len(Obj(ObjId).VertexBuffer.Vertex(0)) * (UBound(Obj(ObjId).VertexBuffer.Vertex) + 1), 0, Obj(ObjId).VertexBuffer.Vertex(0)) = D3D_OK Then MsgBox "Error"
If Not D3DXMeshIndexBuffer8SetData(Stuff(ObjId).Mesh, 0, (UBound(Obj(ObjId).VertexBuffer.Indices) + 1) * Len(Obj(ObjId).VertexBuffer.Indices(0)), 0, Obj(ObjId).VertexBuffer.Indices(0)) = D3D_OK Then MsgBox "Error"

If Stuff(ObjId).Mesh Is Nothing Then MsgBox "Error"
If BBObj = ObjId And RenderBB Then Me.CreateBoundingBox BBObj
'D3DX.ComputeNormals Mesh(ObjId)
End Sub

Sub RenderObjects()
Dim T As Long
Dim A As GlobalVertex
Dim Mat As D3DMATRIX
Dim fMat As D3DMATRIX

'Active lights
If nLights = -1 Then
DefaultLight
Else



End If

With Dev
.SetVertexShader D3DFVF
.GetTransform D3DTS_WORLD, Mat
For T = 0 To nObj

D3DXMatrixMultiply fMat, Mat, Stuff(T).Transformation

.SetTransform D3DTS_WORLD, fMat
.SetMaterial cMtrl(Obj(T).Material)

'.DrawPrimitiveUP D3DPT_TRIANGLELIST, (UBound(Obj(T).VertexBuffer.Vertex) + 1) / 3, Obj(T).VertexBuffer.Vertex(0), Len(A)

Stuff(T).Mesh.DrawSubset 0

If RenderBB And BBObj = T Then
Dim Mtrl As D3DMATERIAL8
Mtrl.emissive.R = 1
Mtrl.emissive.G = 1
Mtrl.emissive.B = 1
Mtrl.emissive.A = 1
.SetMaterial Mtrl
.SetVertexShader BBShader
.DrawPrimitiveUP D3DPT_LINELIST, 12, BB(0), Len(BB(0))
End If

Next



.SetTransform D3DTS_WORLD, Mat
End With
End Sub

Function SetObjRotation(ObjId As Long, Rot As Vector4)
With Obj(ObjId)

.Rotation = Rot
Stuff(ObjId).RotQ = Rota2(.Rotation)
D3DXMatrixAffineTransformation Stuff(ObjId).Transformation, 1, cD3DVECTOR(.Center), Rota2(.Rotation), cD3DVECTOR(.Position)

End With
End Function

Function GetObjRotation(ObjId As Long) As Vector4
GetObjRotation = Obj(ObjId).Rotation
End Function

Function SetObjPos(ObjId As Long, Pos As Vector)
With Obj(ObjId)
.Position = Pos
D3DXMatrixAffineTransformation Transformation(ObjId), 1, cD3DVECTOR(.Cent), Rota(.Rot.X, .Rot.Y, .Rot.Z), cD3DVECTOR(.Pos)
End With
End Function

Function GetObjPos(ObjId As Long) As Vector
With Obj(ObjId).Position
GetObjPos.X = .X
GetObjPos.Y = .Y
GetObjPos.Z = .Z
End With
End Function

Function RotateObject(ObjId As Long, Rot As Vector)
With Obj(ObjId)


Dim T As D3DMATRIX
Dim Q As D3DQUATERNION
Dim Q1 As D3DQUATERNION
Dim T1 As D3DVECTOR
Dim A As Single

Q = Stuff(ObjId).RotQ
Q1 = Rota(Rot)
D3DXQuaternionMultiply Q, Q, Q1
Stuff(ObjId).RotQ = Q

D3DXQuaternionToAxisAngle Q, T1, A



.Rotation.X = T1.X
.Rotation.Y = T1.Y
.Rotation.Z = T1.Z
.Rotation.W = A

D3DXMatrixAffineTransformation Stuff(ObjId).Transformation, 1, cD3DVECTOR(.Center), Rota2(.Rotation), cD3DVECTOR(.Position)


End With
End Function

Sub SetDevice(Device As Direct3DDevice8)
Set Dev = Device
End Sub


Function MouseHit(X As Single, Y As Single) As PickInfo
On Error GoTo erro
    Dim ViewPort As D3DVIEWPORT8
    Dim World As D3DMATRIX
    Dim Proj As D3DMATRIX
    Dim View As D3DMATRIX
    Dim vIn As D3DVECTOR, vNear As D3DVECTOR, vFar As D3DVECTOR, vDir As D3DVECTOR
    Dim bHit As Boolean, i As Long
    Dim T As Long
    Dim currentMatrix As D3DMATRIX
    Dim NewWorldMatrix As D3DMATRIX
    Dim H
Dim Info As PickInfo
Dim Out As PickInfo
Dim Hit As Long
Dim fHit As Boolean

Out.ObjHit = -1

D3DXMatrixIdentity World
Dev.GetViewport ViewPort
Dev.GetTransform D3DTS_PROJECTION, Proj
Dev.GetTransform D3DTS_VIEW, View

For T = 0 To nObj
currentMatrix = Stuff(T).Transformation
D3DXMatrixMultiply NewWorldMatrix, currentMatrix, World

vIn.X = X: vIn.Y = Y
vIn.Z = 0
D3DXVec3Unproject vNear, vIn, ViewPort, Proj, View, NewWorldMatrix
vIn.Z = 1
D3DXVec3Unproject vFar, vIn, ViewPort, Proj, View, NewWorldMatrix
D3DXVec3Subtract vDir, vFar, vNear
Info.ObjHit = T
DX.Intersect Stuff(T).Mesh, vNear, vDir, Hit, Info.FaceIndex, Info.u, Info.V, Info.HitDist, 0
Info.CheckPos = cVector(vNear)
Info.CheckDir = cVector(vDir)

If Hit <> 0 Then
If fHit = False Then

Out = Info

fHit = True
Else

If Info.HitDist < Out.HitDist Then

Out = Info

End If
End If
End If
Next


MouseHit = Out
Exit Function
erro:
MouseHit.ObjHit = -1

End Function


Sub CreateBoundingBox(ObjId As Long)
Dim Max As D3DVECTOR
Dim Min As D3DVECTOR
Dim T As Long

DX.ComputeBoundingBoxFromMesh Stuff(ObjId).Mesh, Min, Max

Min.X = Min.X - Abs((Min.X / 30))
Min.Y = Min.Y - Abs((Min.Y / 30))
Min.Z = Min.Z - Abs((Min.Z / 30))

Max.X = Max.X + Abs((Max.X / 30))
Max.Y = Max.Y + Abs((Max.Y / 30))
Max.Z = Max.Z + Abs((Max.Z / 30))


BB(0).Pos = Vec(Max.X, Max.Y, Max.Z)
BB(1).Pos = Vec(Min.X, Max.Y, Max.Z)

BB(2).Pos = Vec(Max.X, Min.Y, Max.Z)
BB(3).Pos = Vec(Min.X, Min.Y, Max.Z)

BB(4).Pos = Vec(Max.X, Max.Y, Max.Z)
BB(5).Pos = Vec(Max.X, Min.Y, Max.Z)

BB(6).Pos = Vec(Min.X, Max.Y, Max.Z)
BB(7).Pos = Vec(Min.X, Min.Y, Max.Z)



BB(8).Pos = Vec(Max.X, Max.Y, Min.Z)
BB(9).Pos = Vec(Min.X, Max.Y, Min.Z)

BB(10).Pos = Vec(Max.X, Min.Y, Min.Z)
BB(11).Pos = Vec(Min.X, Min.Y, Min.Z)

BB(12).Pos = Vec(Max.X, Max.Y, Min.Z)
BB(13).Pos = Vec(Max.X, Min.Y, Min.Z)

BB(14).Pos = Vec(Min.X, Max.Y, Min.Z)
BB(15).Pos = Vec(Min.X, Min.Y, Min.Z)




BB(16).Pos = Vec(Max.X, Max.Y, Max.Z)
BB(17).Pos = Vec(Max.X, Max.Y, Min.Z)

BB(18).Pos = Vec(Max.X, Min.Y, Max.Z)
BB(19).Pos = Vec(Max.X, Min.Y, Min.Z)



BB(20).Pos = Vec(Min.X, Max.Y, Max.Z)
BB(21).Pos = Vec(Min.X, Max.Y, Min.Z)

BB(22).Pos = Vec(Min.X, Min.Y, Max.Z)
BB(23).Pos = Vec(Min.X, Min.Y, Min.Z)


For T = 0 To 23

BB(T).Diffuse = &HFFFFFFFF


Next
RenderBB = True
BBObj = ObjId
End Sub


Sub DestroyBoundingBox()
RenderBB = False
End Sub

Function MovObject(ObjId As Long, Mov As Vector)
With Obj(ObjId)

.Position = Add(Mov, .Position)

Dim T As D3DMATRIX
Dim Q As D3DQUATERNION

Q = Stuff(ObjId).RotQ

D3DXMatrixAffineTransformation Stuff(ObjId).Transformation, 1, cD3DVECTOR(.Center), Q, cD3DVECTOR(.Position)

End With
End Function


Sub ActiveObj(ObjId As Long, Flags As Boolean)
Dim T() As RollsInfo
Dim P As Long
Dim P2 As Long
Obj(ObjId).Creator.GetArgs T()
View.Clear
For P = 0 To UBound(T)

View.AddRoll T(P).Nome

For P2 = 0 To T(P).nItms
View.AddItem P, T(P).Items(P2).Nome, T(P).Items(P2).Default, T(P).Items(P2).Numb
Next

Next
ActObj = ObjId
View.Update
End Sub

Sub Deactive()

ActObj = -1
View.Clear
View.Update
End Sub

Function CreateName(pName As String) As String
Dim T As String
Dim N As Long
Dim TN As Long
R:

N = N + 1

T = Format(N, "00")

For TN = 0 To nObj

If Obj(TN).Creator.GetName = pName & T Then GoTo R

Next

CreateName = pName & T

End Function


Function SetOptViewer(OptV As Object)
Set View = OptV

End Function



Private Sub View_Change(Index As Integer, Value As String, NewValue As String)
Dim T As Long
Dim T1 As Long
Dim R() As RollsInfo
Dim Ci As Long
If ActObj < 0 Then Exit Sub
'only number filter

Ci = -1

Obj(ActObj).Creator.GetArgs R

For T = 0 To UBound(R)
For T1 = 0 To UBound(R(T).Items)


Ci = Ci + 1

If Ci = Index Then GoTo Fim




Next
Next
Fim:


If R(T).Items(T1).Numb Then

If Not IsNumeric(Value) Then
NewValue = Obj(ActObj).Creator.GetArg(Index)
Exit Sub
End If
End If
Obj(ActObj).Creator.SetArg Index, Value, NewValue
View.Communicate
End Sub

Function GetObjects(Objects() As Object3D)
Objects = Obj
End Function


Public Function RayIntersect(X As Single, Y As Single, ObjId As Long) As PickInfo
    Dim vIn As D3DVECTOR, vNear As D3DVECTOR, vFar As D3DVECTOR, vDir As D3DVECTOR
    Dim VP As D3DVIEWPORT8
    Dim World As D3DMATRIX
    Dim View As D3DMATRIX
    Dim Proj As D3DMATRIX

D3dDevice.GetViewport VP
D3dDevice.GetTransform D3DTS_PROJECTION, Proj
D3dDevice.GetTransform D3DTS_VIEW, View
World = Stuff(ObjId).Transformation


Dim Info As PickInfo

Dim Hit As Long

vIn.X = X: vIn.Y = Y
vIn.Z = 0
D3DXVec3Unproject vNear, vIn, VP, Proj, View, World
vIn.Z = 1
D3DXVec3Unproject vFar, vIn, VP, Proj, View, World
D3DXVec3Subtract vDir, vFar, vNear
D3DXVec3Normalize vDir, vDir
DX.Intersect Stuff(ObjId).Mesh, vNear, vDir, Hit, Info.FaceIndex, Info.u, Info.V, Info.HitDist, 0

Info.CheckPos = cVector(vNear)
Info.CheckDir = cVector(vDir)

Info.ObjHit = Hit

RayIntersect = Info


End Function






Sub RemoveObj(ObjId As Long)
Dim T As Long
Dim C As Long
nObj = nObj - 1
If UBound(Obj) = 0 Then
Erase Obj
Erase Stuff
Exit Sub
End If

On Error Resume Next

Dim tObj() As Object3D

Dim TStuff() As D3DStuff

ReDim tObj(UBound(Obj) - 1)
ReDim TStuff(UBound(Obj) - 1)

Set Obj(ObjId).Creator = Nothing
Set Stuff(ObjId).Mesh = Nothing


C = -1

For T = 0 To UBound(Obj)
If Not T = ObjId Then
C = C + 1
tObj(C) = Obj(T)
TStuff(C) = Stuff(T)
End If


Next

Obj = tObj
Stuff = TStuff

End Sub

Public Function GetnObj() As Long
GetnObj = nObj
End Function

Sub DefaultLight()
    Dim Light As D3DLIGHT8
    Light.Type = D3DLIGHT_POINT

    Light.Diffuse.R = 1
    Light.Diffuse.G = 1
    Light.Diffuse.B = 1

    
    Light.Position = cD3DVECTOR(Cams(ActiveId).Position)

    Light.Range = 10000

    Light.Attenuation0 = 1.05
    D3dDevice.SetLight 0, Light
    D3dDevice.LightEnable 0, True

    D3dDevice.SetRenderState D3DRS_LIGHTING, 1

    D3dDevice.SetRenderState D3DRS_AMBIENT, &HFFAAAAAA

End Sub
