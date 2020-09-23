VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "3D Eagle"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   662
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CC 
      Left            =   9960
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   12632256
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   13157576
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0164
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":297C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sel"
            Description     =   "Selecionar"
            ImageIndex      =   3
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Mov"
            Description     =   "Mover"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Rot"
            Description     =   "Rotacionar"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "X"
            Description     =   "X"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Y"
            Description     =   "Y"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Z"
            Description     =   "Z"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "XZ"
            Description     =   "XZ"
            ImageIndex      =   7
            Style           =   2
            Value           =   1
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox View 
      BackColor       =   &H00FF0000&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   109
      TabIndex        =   7
      Top             =   600
      Width           =   1695
   End
   Begin TabDlg.SSTab Tools 
      Height          =   10410
      Left            =   12000
      TabIndex        =   0
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   18362
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frmMain.frx":4224
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ObjC"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frmMain.frx":4576
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   " "
      TabPicture(2)   =   "frmMain.frx":48C8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Check1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Check2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Command13"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "pBar"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Command4"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command4 
         Caption         =   "Salvar"
         Height          =   375
         Left            =   -74880
         TabIndex        =   15
         Top             =   2520
         Width           =   2775
      End
      Begin MSComctlLib.ProgressBar pBar 
         Height          =   375
         Left            =   -74880
         TabIndex        =   13
         Top             =   2040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin Eagle3D.Opt ObjC 
         Height          =   2655
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4683
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Renderizar"
         Height          =   435
         Left            =   -74880
         TabIndex        =   11
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sombreamento suave"
         Height          =   255
         Left            =   -74880
         TabIndex        =   10
         Top             =   960
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Cores suaves"
         Height          =   255
         Left            =   -74880
         TabIndex        =   9
         Top             =   600
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Caption         =   "Objetos"
         Height          =   2415
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3015
         Begin VB.CommandButton Command2 
            Height          =   255
            Left            =   1680
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Rosca"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Bule"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1680
            Width           =   1335
         End
         Begin VB.CommandButton Command5 
            Caption         =   "Cilindro"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   600
            Width           =   1335
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Caixa"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Esfera"
            Height          =   255
            Left            =   120
            TabIndex        =   2
            Top             =   960
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Const aX = 1
Const aY = 2
Const aZ = 3
Const aXZ = 4



Dim Mouse As Vector
Dim Action As Long
Dim T As Vector
Dim ObjHit As PickInfo
Dim MousePos As POINTAPI
Dim MovAxis As Long

Private Sub Check1_Click()
D3dDevice.SetRenderState D3DRS_DITHERENABLE, Check1.Value
Render
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
D3dDevice.SetRenderState D3DRS_SHADEMODE, 2
Else
D3dDevice.SetRenderState D3DRS_SHADEMODE, 1
End If
Render
End Sub


Private Sub Command1_Click()
Dim T As Long
Dim Sphere As SphereCreator
Set Sphere = New SphereCreator
T = Scene.AddObj(Sphere)
Scene.SetToCreator T
Render
End Sub

Private Sub Command11_Click()
Dim T As Long
Dim Torus As New Rosca


T = Scene.AddObj(Torus)
Scene.SetToCreator T
Render
End Sub

Private Sub Command13_Click()
RenderScene View, pBar
End Sub

Private Sub Command3_Click()
Dim T As Long
Dim Caixa As New Box


T = Scene.AddObj(Caixa)
Scene.SetToCreator T
Render
End Sub

Private Sub Command4_Click()
On Error GoTo erro
CC.ShowSave


If UBound(Dir(CC.FileName)) = UBound(CC.FileTitle) Then

If MsgBox("Esse arquivo já existe, deseja substituí-lo?", vbYesNo) = vbYes Then

Kill CC.FileName
Else
GoTo erro
End If
End If


SavePicture View.Image, CC.FileName
MsgBox "Arquivo salvo"
erro:
End Sub

Private Sub Command5_Click()
Dim T As Long
Dim Cylinder As New Cylin


T = Scene.AddObj(Cylinder)
Scene.SetToCreator T
Render
End Sub

Private Sub Command9_Click()
Dim T As Long
Dim Pot As New Teapot


T = Scene.AddObj(Pot)
Scene.SetToCreator T
Render
End Sub

Private Sub Form_Load()
ObjHit.ObjHit = -1
MovAxis = 4
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
Tools.Left = MeWi - Tools.Width - 8
Tools.Height = MeHe - 27 - 40
ObjC.Height = (Tools.Height * Screen.TwipsPerPixelY) - ObjC.Top - (45 * Screen.TwipsPerPixelY)
View.Height = MeHe - 27 - 40
View.Width = Tools.Left - 1
ElseIf Me.WindowState = 2 Then
Tools.Left = MeWi - Tools.Width - 9
Tools.Height = MeHe - 28 - 40
ObjC.Height = (Tools.Height * Screen.TwipsPerPixelY) - ObjC.Top - (45 * Screen.TwipsPerPixelY)
View.Height = MeHe - 28 - 40
View.Width = Tools.Left - 1
End If

Render
End Sub

Function MeWi() As Long
MeWi = Me.Width / Screen.TwipsPerPixelX
End Function

Function MeHe() As Long
MeHe = Me.Height / Screen.TwipsPerPixelY
End Function

Private Sub Form_Unload(Cancel As Integer)
Cleanup
End
End Sub

Private Sub Objc_Com()
DoEvents
Render
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "Mov"

Action = 1

Case "Rot"

Action = 2
Case "Sel"
Action = 0
Case "X"
MovAxis = aX
Case "Y"
MovAxis = aY

Case "Z"
MovAxis = aZ

Case "XZ"
MovAxis = aXZ

End Select
End Sub

Private Sub View_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then

If ObjHit.HitDist > -1 Then

Scene.RemoveObj ObjHit.ObjHit
Render
End If
End If

End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Mouse.X = X
Mouse.Y = Y
Dim LH As Long

LH = ObjHit.ObjHit
ObjHit = Scene.MouseHit(X, Y)
If ObjHit.ObjHit > -1 Then



If LH <> ObjHit.ObjHit Then
Scene.ActiveObj ObjHit.ObjHit, True
Scene.CreateBoundingBox ObjHit.ObjHit
End If
GetCursorPos MousePos
Render
Else
Scene.Deactive
Scene.DestroyBoundingBox
Render
End If
End Sub

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = 1 And ObjHit.ObjHit >= 0 Then
Dim T As POINTAPI
GetCursorPos T
If MousePos.X <> T.X Or MousePos.Y <> T.Y Then


Dim B As Vector


Dim TX As Single
Dim TY As Single


Select Case Action
Case 1
TX = ((MousePos.X - T.X) / 3)
TY = -((MousePos.Y - T.Y) / 3)

Select Case MovAxis
Case aX
B = Vec3(TX, 0, 0)
Case aY
B = Vec3(0, -TY, 0)
Case aZ
B = Vec3(0, 0, TY)
Case aXZ
B = Vec3(TX, 0, TY)
End Select

Scene.MovObject ObjHit.ObjHit, B
Case 2
TX = ((MousePos.X - T.X) / 150)
TY = -((MousePos.Y - T.Y) / 150)

Select Case MovAxis
Case aX
B = Vec3(0, TY, 0)
Case aY
B = Vec3(TX, 0, 0)
Case aZ
B = Vec3(0, 0, TX)
Case aXZ
B = Vec3(TX, TY, 0)
End Select

Scene.RotateObject ObjHit.ObjHit, B
Case Else
'don't botter rendering


Exit Sub
End Select
Render


SetCursorPos MousePos.X, MousePos.Y
End If
End If
End Sub
