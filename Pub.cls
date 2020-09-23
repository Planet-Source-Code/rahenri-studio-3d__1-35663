VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Public Enum ObjType
tObject = 1
tLight = 2
tCamera = 3
tSystem = 4
End Enum

Public Enum LightType
Point = 1
Spot = 2
Directional = 3
End Enum


Public Type Vector
X As Single
Y As Single
Z As Single
End Type

Public Type Vector4
X As Single
Y As Single
Z As Single
W As Single
End Type

Public Type GlobalVertex
    Position As Vector
    Normal As Vector
    tu As Single
    tv As Single
End Type

Public Type Buffer
Vertex() As GlobalVertex
Indices() As Integer
End Type

Type Color
A As Single
R As Single
G As Single
B As Single
End Type


Type rMtrl
Diffuse As Color
Emissive As Color
Specular As Color
Ambient As Color
Power As Long
End Type

Public Type Object3D
Material As rMtrl
Creator As Object
Modeler As Object
Modifier() As Object
Position As Vector
Rotation As Vector4
Center As Vector
VertexBuffer As Buffer
End Type

Public Type PickInfo
FaceIndex As Long
HitDist As Single
ObjHit As Long
CheckPos As Vector
CheckDir As Vector
U As Single
V As Single
End Type

Public Type Norm
Nor As Vector
nN As Long
End Type

Public Type ItmData
Nome As String
Default As String
Numb As Boolean
End Type

Public Type RollsInfo
Items() As ItmData
Nome As String
nItms As Long
End Type

Public Type tLight
Position As Vector
Diffuse As Color
Specular As Color
CastShadows As Boolean
ShadowColor As Color
LightStyle As LightType
Direction As Vector
Cone As Single
InCone As Single
DecayStarts As Single
Decay As Single
CastShadow As Boolean
End Type

Public Type tCamera
Position As Vector
Direction As Vector
Orientation As Vector
Fov As Single
Apect As Single
End Type
