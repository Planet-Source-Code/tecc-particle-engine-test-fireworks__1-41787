VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5715
   ClientLeft      =   1815
   ClientTop       =   1635
   ClientWidth     =   6585
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   Begin VB.Timer rndSettings 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1200
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   2880
      Top             =   3960
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   2400
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyC
        LL.CLRscrn = Not (LL.CLRscrn)
    Case vbKeyEscape
        LL.ENDok = True
    Case vbKeyO
        LL.Showcolor = Not (LL.Showcolor)
    Case vbKeyE
        LL.ELLIPSE = Not (LL.ELLIPSE)
    Case vbKeyA
        Timer2.Enabled = Not (Timer2.Enabled)
    Case vbKeyR
        rndSettings.Enabled = Not (rndSettings.Enabled)
    Case vbKeyAdd
        If LL.RelativeParticleSize > 1 Then
        LL.RelativeParticleSize = LL.RelativeParticleSize - 1
        End If
        Me.Caption = "Particle size: " & Abs(255 - LL.RelativeParticleSize)
    Case vbKeySubtract
        If LL.RelativeParticleSize < 254 Then
        LL.RelativeParticleSize = LL.RelativeParticleSize + 1
        End If
        Me.Caption = "Particle size: " & Abs(255 - LL.RelativeParticleSize)
End Select
End Sub

Private Sub Form_Load()
ReDim Particles(200)
For i = 0 To UBound(Particles)
    Particles(i).Decay = 255
Next
Me.Show
LL.RelativeParticleSize = 4
Dim LOC1 As PointAPI
LL.CLRscrn = True
ExplodeParticles LOC1, 33, 70, 5
MsgBox "Use these keys: " & vbCrLf & vbCrLf & " C: Toggle Clear screen" & vbCrLf & " O: Toggle color" & vbCrLf & " Escape: Proper exit" & vbCrLf & " E: Toggle draw mode(pixels or ellipse)" & vbCrLf & " A: Toggle Auto-fire" & vbCrLf & " R: Random settings toggle" & vbCrLf & " + or - (numpad): Increase or decrease max particle size" & vbCrLf & vbCrLf & "Mouse has several functions also, play around and have fun"
init
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim LOC1 As PointAPI
LOC1.X = X
LOC1.Y = Y
If Button = 1 Then
        LOC1.X = X
        LOC1.Y = Y
    ExplodeParticles LOC1, 300, 255, 2
ElseIf Button = 2 Then
    ShootParticle LOC1, 255
Else
    ShootParticle LOC1, 255, False, True
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Dim loc1 As PointAPI
'loc1.X = X
'loc1.Y = Y
'ShootParticle loc1, 50, True
End Sub

Private Sub rndSettings_Timer()
Randomize

If Rnd * 10 <= 5 Then
LL.CLRscrn = True
Else
LL.CLRscrn = False
End If
If Rnd * 10 <= 5 Then
LL.ELLIPSE = True
Else
LL.ELLIPSE = False
End If
If Rnd * 10 <= 5 Then
LL.Showcolor = True
Else
LL.Showcolor = False
End If
LL.RelativeParticleSize = Int(Rnd * 100)
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Me.Caption = UBound(Particles) & "   FPS: " & LL.FPS & " - Particle Size: " & Abs(255 - LL.RelativeParticleSize)
LL.FPS = 0
End Sub

Private Sub Timer2_Timer()
Dim SPL As PointAPI
SPL.X = Rnd * Me.ScaleWidth
SPL.Y = Me.ScaleHeight
ShootParticle SPL, 100
End Sub
