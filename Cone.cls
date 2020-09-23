VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'this class module creates a Cone
Dim Obj As Buffer
Dim Tmp As Buffer
Dim SegX As Long
Dim SegY As Long
Dim SegZ As Long
Dim WithEvents View As Opt
Attribute View.VB_VarHelpID = -1
Dim Nome As String
Dim Di As Vector
Dim sObj As Objects
Dim LocalId As Long
Dim Activated As Boolean


Function CreateObject()
Dim X As Long
Dim Y As Long
Dim R As Single
Dim sr As Single
Dim nVec As Single
Dim M As D3DXMesh



Set M = D3DX.CreateCylinder(D3dDevice, 1, 0, 2, SegX, SegY, Nothing)
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
View.Communicate
End Function

Sub SetObjects(sObj1 As Objects, myId As Long)
Set sObj = sObj1
LocalId = myId
Nome = sObj.CreateName("Cone")
Di = Vec3(30, 30, 30)
SegX = 16
SegY = 16
SegX = 16
End Sub



Function SetOptViewer(OptV As Object)
Set View = OptV

End Function

Function ShowOpt()
View.Clear
View.AddRoll "Identificação"
View.AddItem 0, "Nome", Nome, False
View.AddRoll "Criação"
View.AddItem 1, "Segmentos X", Trim(SegX), True
View.AddItem 1, "Segmentos Y", Trim(SegY), True
View.AddItem 1, "Segmentos Z", Trim(SegY), True
View.AddRoll "Dimensão"
View.AddItem 2, "Largura", Trim(Di.X), True
View.AddItem 2, "Altura", Trim(Di.Y), True
View.AddItem 2, "Profundidade", Trim(Di.Z), True
View.Update
End Function

Private Sub View_Change(Index As Integer, Value As String, NewValue As String)

If Not Activated Then Exit Sub
If Index = 1 Or Index = 2 Then
NewValue = Trim(Str(CInt(Value)))
ElseIf Index > 2 Then
NewValue = Trim(CSng(Value))
End If

Select Case Index
Case 0
Nome = Value
Case 1
If CSng(NewValue) > 0 Then
SegX = NewValue
CreateObject
Else
NewValue = SegX
End If
Case 2
If CSng(NewValue) > 0 Then
SegY = NewValue
CreateObject
Else
NewValue = SegY
End If
Case 3
If CSng(NewValue) > 0 Then
SegZ = NewValue
CreateObject
Else
NewValue = SegZ
End If
Case 4
Di.X = NewValue
ReDimObj
Case 5
Di.Y = NewValue
ReDimObj
Case 6
Di.Z = NewValue
ReDimObj
End Select

End Sub

Function RequestBuffer(rBuf As Buffer)

Dim T As Long
rBuf = Tmp

End Function

Sub Active(Flags As Boolean)
If Flags <> Activated Then
Activated = Flags

If Not Flags Then

View.Clear

Else

ShowOpt

End If
End If
End Sub

Function GetName() As String
GetName = Nome
End Function


