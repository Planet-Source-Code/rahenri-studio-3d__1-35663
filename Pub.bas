Attribute VB_Name = "Pub"
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public Scene As Objects
Public Const PI = 3.14159265358979
Public Const D3DFVF = (D3DFVF_XYZ Or D3DFVF_NORMAL Or D3DFVF_TEX1)

Public Function Vec3(X As Single, Y As Single, Z As Single) As Vector
Vec3.X = X
Vec3.Y = Y
Vec3.Z = Z
End Function

Public Function Vec4(X As Single, Y As Single, Z As Single, W As Single) As Vector4
Vec4.X = X
Vec4.Y = Y
Vec4.Z = Z
Vec4.W = W
End Function

Public Function cD3DVECTOR(V As Vector) As D3DVECTOR
cD3DVECTOR.X = V.X
cD3DVECTOR.Y = V.Y
cD3DVECTOR.Z = V.Z
End Function


Public Function cVector(V As D3DVECTOR) As Vector
cVector.X = V.X
cVector.Y = V.Y
cVector.Z = V.Z
End Function

Public Function CalcNormal(V0 As Vector, V1 As Vector, V2 As Vector) As Vector
Dim Side1 As Vector
Dim Side2 As Vector
Dim Normalizer As Single
Dim TVec As Vector
On Error Resume Next
    Side1.X = (V1.X - V0.X)
    Side1.Y = (V1.Y - V0.Y)
    Side1.Z = (V1.Z - V0.Z)
   
    Side2.X = (V2.X - V1.X)
    Side2.Y = (V2.Y - V1.Y)
    Side2.Z = (V2.Z - V1.Z)
    
    TVec.X = (Side1.Y * Side2.Z) - (Side1.Z * Side2.Y)
    TVec.Y = (Side1.Z * Side2.X) - (Side1.X * Side2.Z)
    TVec.Z = (Side1.X * Side2.Y) - (Side1.Y * Side2.X)
    
    
    Normalizer = Sqr((TVec.X * TVec.X) + _
                    (TVec.Y * TVec.Y) + _
                    (TVec.Z * TVec.Z))
    CalcNormal.X = TVec.X / Normalizer
    CalcNormal.Y = TVec.Y / Normalizer
    CalcNormal.Z = TVec.Z / Normalizer
    
End Function

Public Function ColorValue4(A As Single, R As Single, G As Single, B As Single) As D3DCOLORVALUE

    ColorValue4.A = A
    ColorValue4.R = R
    ColorValue4.G = G
    ColorValue4.B = B

End Function

Public Function Col(A As Single, R As Single, G As Single, B As Single) As Color

    Col.A = A
    Col.R = R
    Col.G = G
    Col.B = B

End Function

Public Function cMtrl(Mtrl As rMtrl) As D3DMATERIAL8

cMtrl.Ambient = Col1(Mtrl.Ambient)
cMtrl.emissive = Col1(Mtrl.emissive)
cMtrl.Diffuse = Col1(Mtrl.Diffuse)
cMtrl.specular = Col1(Mtrl.specular)
cMtrl.power = Mtrl.power

End Function

Public Function Col1(Cor As Color) As D3DCOLORVALUE

    Col1.A = Cor.A
    Col1.R = Cor.R
    Col1.G = Cor.G
    Col1.B = Cor.B

End Function

Function Rota(V As Vector) As D3DQUATERNION
    Dim quat As D3DQUATERNION
    D3DXQuaternionRotationYawPitchRoll quat, V.X, V.Y, V.Z
    
    Rota = quat
End Function

Function Rota2(V As Vector4) As D3DQUATERNION
    Dim quat As D3DQUATERNION
    D3DXQuaternionRotationAxis quat, Vec(V.X, V.Y, V.Z), V.W
    
    Rota2 = quat
End Function

Function Vec(X As Single, Y As Single, Z As Single) As D3DVECTOR
Vec.X = X
Vec.Y = Y
Vec.Z = Z
End Function

Function Add(V1 As Vector, V2 As Vector) As Vector
Add.X = V1.X + V2.X
Add.Y = V1.Y + V2.Y
Add.Z = V1.Z + V2.Z
End Function

Function VecMultiply(V1 As Vector, M As Single) As Vector
VecMultiply.X = V1.X * M
VecMultiply.Y = V1.Y * M
VecMultiply.Z = V1.Z * M
End Function

Public Function MixNormals(outN As Vector, N1 As Vector, N2 As Vector, N3 As Vector, F1 As Single, F2 As Single, F3 As Single)

outN.X = (N1.X * F1) + (N2.X * F2) + (N3.X * F3)
outN.Y = (N1.Y * F1) + (N2.Y * F2) + (N3.Y * F3)
outN.Z = (N1.Z * F1) + (N2.Z * F2) + (N3.Z * F3)

End Function

Public Function GetLineSize(Vec As Vector) As Single
Dim C2 As Single
C2 = Sqr(Q(Vec.X) + Q(Vec.Z))
GetLineSize = Sqr(Q(Vec.Y) + Q(C2))
End Function

Public Function Q(N As Single) As Single
Q = N * N
End Function

Public Function Mix(outColor As Long, inColor As Long, inColor1 As Long, offSet As Double)
Dim oC As Color
Dim oC1 As Color
Dim D As Long
Dim number
number = inColor
oC.R = number Mod 256
number = number - (number Mod 256)
number = number / 256
oC.G = number Mod 256
number = number - (number Mod 256)
number = number / 256
oC.B = number Mod 256

number = inColor1
oC1.R = number Mod 256
number = number - (number Mod 256)
number = number / 256
oC1.G = number Mod 256
number = number - (number Mod 256)
number = number / 256
oC1.B = number Mod 256


D = oC.R - oC1.R
D = D * offSet
oC.R = oC.R - D

D = oC.G - oC1.G
D = D * offSet
oC.G = oC.G - D

D = oC.B - oC1.B
D = D * offSet
oC.B = oC.B - D

outColor = RGB(oC.R, oC.G, oC.B)
End Function


