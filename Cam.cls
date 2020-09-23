VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Camera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim Cam As tCamera


Public Function SetCam(CamId As Long, sCam As tCamera)
Cams(CamId) = sCam
End Function

Public Function GetCam() As tCamera

End Function


Private Sub Class_Initialize()
nCams = -1
End Sub


Public Property Get TypeId() As ObjType

TypeId = tCamera

End Property


