VERSION 5.00
Begin VB.UserControl Opt 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   ScaleHeight     =   3600
   ScaleWidth      =   4950
   Begin VB.VScrollBar Scrool 
      Height          =   3615
      Left            =   4680
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox lbl 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   0
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Frame Frame 
         Caption         =   "+/- TabCaption"
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   360
         Width           =   855
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label sizer 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   2520
      TabIndex        =   5
      Top             =   3120
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "Opt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim RollsWidth As Long

Private Type Item
OnlyNumber As Boolean
Index As Long
End Type

Private Type Roll
Caption As String
nItems As Long
Items() As Item
lblWidth As Long
End Type

Dim Rolls() As Roll
Dim nRolls As Long
Dim CurrentIndex As Long

Event Change(Index As Integer, Value As String, NewValue As String)

Event Com()
Private Sub Scrool_Change()
Picture1.Top = -Scrool.Value * 15
End Sub

Private Sub Scrool_Scroll()
Picture1.Top = -Scrool.Value * 15
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim k As String
If KeyAscii = 13 Then
RaiseEvent Change(Index, txt(Index).Text, k)
If k <> "" Then txt(Index).Text = k
KeyAscii = 0
End If

If txt(Index).Tag = True Then
If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Or KeyAscii = 44 Or KeyAscii = 46) Then KeyAscii = 0
End If


End Sub

Private Sub txt_LostFocus(Index As Integer)
Dim k As String
RaiseEvent Change(Index, txt(Index).Text, k)
If k <> "" Then txt(Index).Text = k
End Sub

Private Sub UserControl_Initialize()
nRolls = -1
CurrentIndex = -1
End Sub

Private Sub UserControl_Resize()
RollsWidth = UserControl.Width - Scrool.Width
Scrool.Left = RollsWidth
Scrool.Height = UserControl.Height
Picture1.Width = RollsWidth
Update
End Sub

Function Update()
Dim count As Long
Dim count1 As Long
Dim Height As Long
Dim Top As Long
Dim count2 As Long
For count = 0 To nRolls

Height = 180 + ((Rolls(count).nItems + 1) * 360)

If Frame(count).Height <> Height Then Frame(count).Height = Height
If Frame(count).Top <> Top Then Frame(count).Top = Top
If Frame(count).Width <> RollsWidth Then Frame(count).Width = RollsWidth
If Frame(count).Visible = False Then Frame(count).Visible = True

For count2 = 0 To Rolls(count).nItems
lbl(Rolls(count).Items(count2).Index).Top = Top + 210 + (count2 * 360)
lbl(Rolls(count).Items(count2).Index).Visible = True
lbl(Rolls(count).Items(count2).Index).Left = 60
lbl(Rolls(count).Items(count2).Index).Width = Rolls(count).lblWidth
lbl(Rolls(count).Items(count2).Index).ZOrder 0
txt(Rolls(count).Items(count2).Index).Left = 90 + Rolls(count).lblWidth
txt(Rolls(count).Items(count2).Index).Width = Frame(count).Width - 165 - Rolls(count).lblWidth
txt(Rolls(count).Items(count2).Index).Top = Top + 180 + (count2 * 360)
txt(Rolls(count).Items(count2).Index).Visible = True
txt(Rolls(count).Items(count2).Index).ZOrder 0
Next

Top = Top + Height + 45


Next
If Top > 0 Then
Picture1.Height = Top - 45
Else
Picture1.Height = 0
End If
Height = Top - 45
Height = Height - UserControl.Height
If Height < 0 Then
Scrool.Enabled = False
Else
Scrool.Enabled = True
Scrool.Max = Height / 15
End If
End Function

Function AddRoll(Caption As String) As Long
nRolls = nRolls + 1
ReDim Preserve Rolls(nRolls)
Rolls(nRolls).nItems = -1
If Not nRolls = 0 Then
Load Frame(nRolls)
End If
Frame(nRolls).Caption = Caption
End Function

Function AddItem(RollId As Long, Caption As String, DeafaultValue As String, OnlyNumber As Boolean) As Long
With Rolls(RollId)
.nItems = .nItems + 1
ReDim Preserve .Items(.nItems)
.Items(.nItems).OnlyNumber = OnlyNumber
CurrentIndex = CurrentIndex + 1
AddItem = CurrentIndex
.Items(.nItems).Index = CurrentIndex

If CurrentIndex > 0 Then
Load txt(CurrentIndex)
Load lbl(CurrentIndex)
End If
txt(CurrentIndex).Text = DeafaultValue
txt(CurrentIndex).Tag = OnlyNumber
sizer.Caption = Caption
If sizer.Width > .lblWidth Then .lblWidth = sizer.Width
lbl(CurrentIndex).Text = Caption
lbl(CurrentIndex).Tag = OnlyNumber
End With
End Function

Function Clear()
Dim C As Long
For C = 1 To nRolls
Unload Frame(C)
Next

For C = 1 To CurrentIndex

Unload txt(C)
Unload lbl(C)



Next
Frame(0).Visible = False
txt(0).Visible = False
lbl(0).Visible = False
ReDim Rolls(0)
CurrentIndex = -1
nRolls = -1
End Function

Sub Communicate()
RaiseEvent Com
End Sub
