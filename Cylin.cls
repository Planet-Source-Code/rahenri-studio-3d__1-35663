VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cylin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class module creates a Cylinder

Dim Obj As Buffer
Dim Tmp As Buffer
Dim SegX As Long
Dim SegY As Long
Dim Nome As String
Dim Di As Vector
Dim sObj As Objects
Dim LocalId As Long
Dim Args() As RollsInfo
Dim R1 As Single, R2 As Single
Dim nRolls As Long
Function CreateObject()

Dim X As Long
Dim Y As Long
Dim R As Single
Dim sr As Single
Dim nVec As Single
Dim M As D3DXMesh

Set M = D3DX.CreateCylinder(D3dDevice, R1, R2, 2, SegX, SegY, Nothing)
Set M = M.CloneMeshFVF(D3DXMESH_MANAGED, D3DFVF, D3dDevice)

ReDim Obj.Vertex(M.GetNumVertices - 1)
D3DVertexBuffer8GetData M.GetVertexBuffer, 0, Len(Obj.Vertex(0)) * M.GetNumVertices, 0, Obj.Vertex(0)

ReDim Obj.Indices((M.GetNumFaces * 3) - 1)

D3DIndexBuffer8GetData M.GetIndexBuffer, 0, ((M.GetNumFaces * 3)) * 2, 0, Obj.Indices(0)

ReDimObj

End Function


Function ReDimObj()
Dim T As Long
Dim N As Vector

Tmp = Obj

For T = 0 To UBound(Obj.Vertex)


Tmp.Vertex(T) = Obj.Vertex(T)
Tmp.Vertex(T).Position.X = Tmp.Vertex(T).Position.X * Di.X
Tmp.Vertex(T).Position.Y = Tmp.Vertex(T).Position.Y * Di.Y
Tmp.Vertex(T).Position.Z = Tmp.Vertex(T).Position.Z * Di.Z


Next

Dim V1 As Vector
Dim V2 As Vector
Dim V3 As Vector

Dim No() As Norm
ReDim No(UBound(Obj.Vertex))
Dim nArray As Long


For T = 0 To UBound(Obj.Indices) Step 3

V1 = Obj.Vertex(Obj.Indices(T)).Position
V2 = Obj.Vertex(Obj.Indices(T + 1)).Position
V3 = Obj.Vertex(Obj.Indices(T + 2)).Position

N = CalcNormal(V1, V2, V3)


No(Obj.Indices(T)).Nor = Add(N, No(Obj.Indices(T)).Nor)
No(Obj.Indices(T)).nN = No(Obj.Indices(T)).nN + 1

No(Obj.Indices(T + 1)).Nor = Add(N, No(Obj.Indices(T + 1)).Nor)
No(Obj.Indices(T + 1)).nN = No(Obj.Indices(T + 1)).nN + 1

No(Obj.Indices(T + 2)).Nor = Add(N, No(Obj.Indices(T + 2)).Nor)
No(Obj.Indices(T + 2)).nN = No(Obj.Indices(T + 2)).nN + 1




Next
Dim T2 As Long

For T = 0 To UBound(No)
N = No(T).Nor



N.X = N.X / No(T).nN
N.Y = N.Y / No(T).nN
N.Z = N.Z / No(T).nN

Tmp.Vertex(T).Normal = N
Next

sObj.CallUpdateObj LocalId

End Function

Sub SetObjects(sObj1 As Objects, myId As Long)
Set sObj = sObj1
LocalId = myId
Nome = sObj.CreateName("Rosca")
Di = Vec3(30, 30, 30)
SegX = 16
SegY = 8
R1 = 1
R2 = 1
CreateArgs
End Sub





Function GetArgs(Out() As RollsInfo)
Out = Args
End Function

Public Sub SetArg(Index As Integer, Value As String, NewValue As String)

Select Case Index
Case 0
If Value <> "" Then
Nome = Value
Else
NewValue = Nome
End If

Case 1
If Int(Value) > 0 Then
SegX = Int(Value)
Else
NewValue = SegX
End If
Me.CreateObject
Case 2
If Int(Value) > 0 Then
SegY = Int(Value)
Else
NewValue = SegY
End If
Me.CreateObject
Case 3
If CSng(Value) >= 0 Then
R1 = CSng(Value)
Else
NewValue = R1
End If
Me.CreateObject
Case 4
If CSng(Value) >= 0 Then
R2 = CSng(Value)
Else
NewValue = R2
End If
Me.CreateObject
Case 5
If CSng(Value) > 0 Then
Di.X = CSng(Value)
Else
NewValue = Di.X
End If
Me.ReDimObj
Case 6
If CSng(Value) > 0 Then
Di.Y = CSng(Value)
Else
NewValue = Di.Y
End If
Me.ReDimObj
Case 7
If CSng(Value) > 0 Then
Di.Z = CSng(Value)
Else
NewValue = Di.Z
End If
Me.ReDimObj
End Select
CreateArgs
End Sub

Public Function GetArg(Index As Integer) As String

Select Case Index
Case 0
GetArg = Nome
Case 1
GetArg = Trim(SegX)
Case 2
GetArg = Trim(SegY)
Case 3
GetArg = Trim(R1)
Case 4
GetArg = Trim(R2)
Case 5
GetArg = Trim(Di.X)
Case 6
GetArg = Trim(Di.Y)
Case 7
GetArg = Trim(Di.Z)
End If
End Select

End Function

Function RequestBuffer(rBuf As Buffer)

Dim T As Long

rBuf = Tmp

End Function



Function GetName() As String
GetName = Nome
End Function

Private Sub AddRoll(Nome As String)
nRolls = nRolls + 1
ReDim Preserve Args(nRolls)
Args(nRolls).Nome = Nome
Args(nRolls).nItms = -1
End Sub

Private Sub AddItem(Nome As String, Value As String, Numb As Boolean)
Args(nRolls).nItms = Args(nRolls).nItms + 1
ReDim Preserve Args(nRolls).Items(Args(nRolls).nItms)
Args(nRolls).Items(Args(nRolls).nItms).Nome = Nome
Args(nRolls).Items(Args(nRolls).nItms).Default = Value
Args(nRolls).Items(Args(nRolls).nItms).Numb = Numb
End Sub

Private Sub CreateArgs()
nRolls = -1
Erase Args
AddRoll "Identificação"
AddItem "Nome", Nome, False
AddRoll "Criação"
AddItem "Segmentos X", Trim(SegX), True
AddItem "Segmentos Y", Trim(SegY), True
AddItem "Raio 1", Trim(R1), True
AddItem "Raio 2", Trim(R2), True
AddRoll "Dimensão"
AddItem "Largura", Trim(Di.X), True
AddItem "Altura", Trim(Di.Y), True
AddItem "Profundidade", Trim(Di.Z), True
End Sub


Public Property Get TypeId() As ObjType

TypeId = tObject

End Property
