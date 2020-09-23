VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RayLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class mudule controls RayLights
Dim Lights() As tLight
Dim nLights As Long
Dim AmbientLight As Long

Public Function AddLight(Light As tLight) As Long
nLights = nLights + 1
ReDim Preserve Lights(nLights)
Lights(nLights) = Light
AddLight = nLights
End Function

Public Function GetLight(LightId As Long) As tLight
GetLight = Lights(LightId)
End Function

Public Function SetLight(LightId As Long, sLight As tLight)
Lights(LightId) = sLight
End Function

Public Function GetAmbientLight() As Long
GetAmbientLight = AmbientLight
End Function

Public Function SetAmbientLight(AL As Long)
AmbientLight = AL
End Function


Sub SetupLights()

    Dim Light As D3DLIGHT8
    Light.Type = D3DLIGHT_POINT

    Light.Diffuse.R = 1
    Light.Diffuse.G = 1
    Light.Diffuse.B = 1
    'light.specular = ColorValue4(1, 1, 1, 1)
    
    Light.Position = Vec(0, 50, 200)

    Light.Range = 10000

    Light.Attenuation0 = 1.05
    D3dDevice.SetLight 0, Light
    D3dDevice.LightEnable 0, True

    D3dDevice.SetRenderState D3DRS_LIGHTING, 1

    D3dDevice.SetRenderState D3DRS_AMBIENT, &HFFAAAAAA
    
End Sub


Public Property Get TypeId() As ObjType

TypeId = tLight

End Property
