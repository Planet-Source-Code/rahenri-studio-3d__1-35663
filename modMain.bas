Attribute VB_Name = "modMain"
Public Const GrdShader = D3DFVF_XYZ
Public Grid() As Vector


Sub Main()

frmSplash.Show
DoEvents

Dim X As Long
Dim Y As Long
Dim C As Long
Dim T As tLight
Dim T2 As tCamera
Set Scene = New Objects




ReDim Grid(240)

For X = 0 To 59

Grid(C).X = (-29.5 * 5)
Grid(C).Z = (X * 5) - (29.5 * 5)
C = C + 1
Grid(C).X = (29.5 * 5)
Grid(C).Z = (X * 5) - (29.5 * 5)
C = C + 1
Grid(C).X = (X * 5) - (29.5 * 5)
Grid(C).Z = (-29.5 * 5)
C = C + 1
Grid(C).X = (X * 5) - (29.5 * 5)
Grid(C).Z = (29.5 * 5)
C = C + 1
Next



'after everything is set, start the programm
frmMain.Show


If Not InitD3D(frmMain.View.hWnd, D3DDEVTYPE_HAL) Then MsgBox "Não foi possível iniciar a DirectX": End
On Error Resume Next
T2.Position = Vec3(0, 50, 200)
T2.Direction = Vec3(0, 0, 0)
T2.Apect = frmMain.View.ScaleHeight / frmMain.View.ScaleWidth
T2.Fov = PI / 4
T2.Orientation = Vec3(0, 1, 0)

Scene.AddCam T2
Scene.ActiveCam 0


DoEvents

Scene.SetOptViewer frmMain.ObjC
Scene.SetDevice D3dDevice

'frmMain.SetFocus
DoEvents

Unload frmSplash
Render

End Sub
