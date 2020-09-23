Attribute VB_Name = "Module1"
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ELLIPSE Lib "gdi32" Alias "Ellipse" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Type PointAPI
    X As Double
    Y As Double
End Type
Public Type RGBTRI
    R As Long
    G As Long
    B As Long
End Type
Public Enum particleType
    tDefault = 0
    tShell = 1
    tStatic = 2
    dupe = 3
End Enum
Public Type Particle
    CLocation As PointAPI
    Speed As PointAPI
    Offset As PointAPI
    Decay As Integer
    Acceleration As Double
    nType As particleType
    Exploded As Boolean
    Color As RGBTRI
    TypeCol As Byte

End Type
Public Type settings
    CLRscrn As Boolean
    MaxParticles As Long
    ENDok As Boolean
    Showcolor As Boolean
    FPS As Long
    ELLIPSE As Boolean
    RelativeParticleSize As Byte
End Type
Public Particles() As Particle
Public LL As settings
Public Sub ExplodeParticles(Location As PointAPI _
            , ParticleCount As Integer, life As Byte, Speed As Integer, Optional ByVal nExclude As Integer)
Dim inactiveParticles() As Integer
Dim scol As Byte
ReDim inactiveParticles(1)
For i = 0 To UBound(Particles)
    With Particles(i)
        If .Decay >= 255 Then
        ReDim Preserve inactiveParticles(UBound(inactiveParticles) + 1)
        inactiveParticles(UBound(inactiveParticles)) = i
        End If
    End With
Next
'On Error Resume Next
Do Until UBound(inactiveParticles) >= ParticleCount
    ReDim Preserve Particles(UBound(Particles) + 1)
        Particles(UBound(Particles)).Decay = 0
    ReDim Preserve inactiveParticles(UBound(inactiveParticles) + 1)
    inactiveParticles(UBound(inactiveParticles)) = UBound(Particles)

DoEvents
Loop

'ReDim Preserve Particles(UBound(Particles) + ParticleCount)
Dim Curparticle As Integer
Dim colo As RGBTRI
Dim colo1 As RGBTRI
    If Rnd * 10 <= 5 Then
        scol = 1
        colo.R = 255
        colo.G = 0
        colo.B = 0
    Else
        If Rnd * 10 <= 5 Then
            scol = 2
            colo.R = 0
            colo.G = 255
            colo.B = 0
        Else
            scol = 3
            colo.R = 0
            colo.G = 0
            colo.B = 255
        End If
    End If
For Curparticle = 0 To ParticleCount
    With Particles(inactiveParticles(Curparticle))
        .CLocation = Location
        colo1.R = colo.R
        colo1.G = colo.G
        colo1.B = colo.B
        .Speed.X = IIf(Rnd * 10 <= 5, Rnd * Speed, -(Rnd * Speed))
        .Speed.Y = IIf(Rnd * 10 <= 5, Rnd * Speed, -(Rnd * Speed))
        .Acceleration = Val("0." & (Rnd * 100))
        .Offset.X = Rnd * 4
        .Offset.Y = Rnd * 4
        .Decay = Abs(255 - life)
        .nType = 0
        .TypeCol = scol
        .Exploded = False
        .Color = colo1
    End With
Next
End Sub

Public Sub ShootParticle(nLocation As PointAPI, life As Byte, Optional Lrandom As Boolean, Optional Stagnant As Boolean)
Dim UseThisParticle As Long
For i = 0 To UBound(Particles)
    With Particles(i)
        If .Decay = 255 Then
            UseThisParticle = i
            GoTo UTP:
        End If
    End With
Next
UseThisParticle = UBound(Particles) + 1
DoEvents
On Error Resume Next
ReDim Preserve Particles(UseThisParticle)
UTP:
With Particles(UseThisParticle)
    .CLocation = nLocation
    .Decay = 0
    If Lrandom Then
     .Speed.Y = IIf(Rnd * 10 <= 5, Rnd * 3, -(Rnd * 3))
     .Speed.X = IIf(Rnd * 10 <= 5, Rnd * 3, -(Rnd * 3))
    Else
    .Speed.X = 0
    .Speed.Y = -(2)
    End If
    If Stagnant Then
    .nType = 2
    Else
    .nType = 1
    End If
    .Offset.X = Rnd * 2
    .Offset.Y = 0
    .Acceleration = IIf(Rnd * 2 = 1, 0.3, 0.5)
    .Exploded = False
End With
End Sub

Public Sub init()
Dim Curparticle As Integer

Dim CL As PointAPI
Dim CB As Byte
Dim sze As Long

aa:
Do
If LL.CLRscrn Then
Form1.Cls
Form1.AutoRedraw = True
Else
Form1.AutoRedraw = False
End If
On Error Resume Next

For Curparticle = 0 To UBound(Particles)
    With Particles(Curparticle)
        
        CL = .CLocation
        If .nType = 0 Then
        .CLocation.X = .CLocation.X + .Speed.X + IIf(.Speed.X < 0, -(.Acceleration), (.Acceleration))
        .CLocation.Y = .CLocation.Y + .Speed.Y + IIf(.Speed.Y < 0, -(.Acceleration), (.Acceleration))
        Else
        .CLocation.X = .CLocation.X + .Speed.X + IIf(Rnd * 10 <= 5, Rnd * 1, -(Rnd * 1))
        .CLocation.Y = .CLocation.Y + .Speed.Y + IIf(Rnd * 10 <= 5, Rnd * 1, -(Rnd * 1)) - .Acceleration
        End If
        sze = Abs(255 - .Decay)
        .Acceleration = .Acceleration - 0.01
        If .Decay < 255 Then
            CB = Abs(255 - .Decay)
            .Decay = .Decay + 2
            If .nType = 0 Then
                If LL.Showcolor Then
                If Not (LL.ELLIPSE) Then
                .Color.R = Abs(.Color.R - Rnd * 2)
                .Color.G = Abs(.Color.G - Rnd * 2)
                .Color.B = Abs(.Color.B - Rnd * 2)
                SetPixelV Form1.hdc, .CLocation.X _
            , .CLocation.Y, RGB(.Color.R, .Color.G, .Color.B)
            Else
            .Color.R = Abs(.Color.R - Rnd * 4)
                .Color.G = Abs(.Color.G - Rnd * 4)
                .Color.B = Abs(.Color.B - Rnd * 4)
            Form1.FillColor = RGB(.Color.R, .Color.G, .Color.B)
            Form1.ForeColor = RGB(.Color.R, .Color.G, .Color.B)
            ELLIPSE Form1.hdc, .CLocation.X, _
            .CLocation.Y, .CLocation.X + sze / LL.RelativeParticleSize, _
            .CLocation.Y + sze / LL.RelativeParticleSize
                End If
                Else
                If Not (LL.ELLIPSE) Then
                SetPixelV Form1.hdc, .CLocation.X _
            , .CLocation.Y, RGB(CB, CB, CB)
            Else
            Form1.FillColor = RGB(CB, CB, CB)
            Form1.ForeColor = RGB(CB, CB, CB)
            ELLIPSE Form1.hdc, .CLocation.X, _
            .CLocation.Y, .CLocation.X + sze / LL.RelativeParticleSize, _
            .CLocation.Y + sze / LL.RelativeParticleSize
                End If
                End If
            Else
            If Not (LL.ELLIPSE) Then
            SetPixelV Form1.hdc, .CLocation.X _
            , .CLocation.Y, RGB(255, 255, 255)
            Else
            ELLIPSE Form1.hdc, .CLocation.X, _
            .CLocation.Y, .CLocation.X + sze / LL.RelativeParticleSize, _
            .CLocation.Y + sze / LL.RelativeParticleSize
            End If
            End If
            
        Else
            If .nType = 1 Then
                If .Exploded = False Then
                .Decay = 0
                ExplodeParticles CL, Rnd * (UBound(Particles) / 4), 200, (Rnd * 3) + 1
                'ShootParticle .CLocation, 100, True
                .Decay = 255
                .Exploded = True
                End If
            Else

            End If
        End If
    End With
Next

LL.FPS = LL.FPS + 1
DoEvents
Sleep 1
Loop Until LL.ENDok = True
End
End Sub
Public Sub ExplodeCURParticles(LOC1 As PointAPI)
For i = 0 To UBound(Particles)
With Particles(UseThisParticle)
    If .Decay = 255 Then
    .nType = 1
    .CLocation = LOC1
    .Decay = 0
    If Lrandom Then
     .Speed.Y = IIf(Rnd * 10 <= 5, Rnd * 3, -(Rnd * 3))
     .Speed.X = IIf(Rnd * 10 <= 5, Rnd * 2, -(Rnd * 2))
    Else
    .Speed.X = Rnd * 2
    .Speed.Y = -(Rnd * 2)
    End If
    .Offset.Y = 0
    .Acceleration = 0.3
    .Exploded = True
    End If
End With
Next

End Sub

Public Sub DupeParticle(LOC1 As PointAPI)
For i = 0 To UBound(Particles)
With Particles(UseThisParticle)
    If .Decay = 255 Then
    .nType = 1
    .CLocation = LOC1
    .Decay = 0

     .Speed.Y = IIf(Rnd * 10 <= 5, Rnd * 3, -(Rnd * 3))
     .Speed.X = IIf(Rnd * 10 <= 5, Rnd * 2, -(Rnd * 2))

    .Acceleration = 0.3
    .Exploded = True
    .nType = 3
    End If
End With
Next
End Sub
