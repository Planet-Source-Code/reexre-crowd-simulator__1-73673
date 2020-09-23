Attribute VB_Name = "modCrowd"
'Author :Roberto Mior
'     reexre@gmail.com
'
'If you use source code or part of it please cite the author
'You can use this code however you like providing the above credits remain intact


Option Explicit

Type tHUMAN

    X              As Single
    Y              As Single
    A              As Single
    incA           As Single

    ATar           As Single

    V              As Single
    Vstd           As Single
    Vmul           As Single

    tX             As Single
    tY             As Single

    GridX          As Long
    GridY          As Long

    
    color          As Long

    rX             As Single
    rY             As Single

    TrailX(0 To 50) As Long
    TrailY(0 To 50) As Long

    CosA As Single
    SinA As Single
    vX As Single
    vY As Single
    
End Type

'global Const TrailStep As Long = 20
'global Const TrailLen As Long = 0 '1
Global TrailStep As Long
Global TrailLen As Long


Global NH          As Long
Global H()         As tHUMAN
Global Const r     As Single = 5
Global Const GridView As Long = 60 '30




Global Const Rsq2  As Single = (6 + 6) ^ 2
Global Const MinDIST As Single = 6 + 6

Global rZOOM As Long


Global Const PIh = 1.5707963267949
Global Const PI = 3.14159265358979
Global Const PI2 = 6.28318530717959
Global Const Amax = 3.14159265358979 * 0.15

Global MaxX        As Long
Global MaxY        As Long


Global ZOOM        As Single

Global CT          As Long


Public Function Atan2(ByVal dX As Single, ByVal dY As Single) As Single
'This Should return Angle

    Dim Theta      As Single

    If (Abs(dX) < 0.0000001) Then
        If (Abs(dY) < 0.0000001) Then
            Theta = 0
        ElseIf (dY > 0) Then
            Theta = PIh
            'theta = PI / 2
        Else
            Theta = -PIh
            'theta = -PI / 2
        End If
    Else
        Theta = Atn(dY / dX)

        If (dX < 0) Then
            If (dY >= 0) Then
                Theta = Theta + PI
            Else
                Theta = Theta - PI
            End If
        End If
    End If


    Atan2 = Theta

    'While Atan2 < 0: Atan2 = Atan2 + PI: Wend
    'While Atan2 > PI: Atan2 = Atan2 - PI: Wend

End Function
Public Function AngleDIFF(A1 As Single, A2 As Single) As Single
'double difference = secondAngle - firstAngle;
'while (difference < -180) difference += 360;
'while (difference > 180) difference -= 360;
'return difference;

    AngleDIFF = A2 - A1
    While AngleDIFF < -PI
        AngleDIFF = AngleDIFF + PI2
    Wend
    While AngleDIFF > PI
        AngleDIFF = AngleDIFF - PI2
    Wend



    '''' this is to have values between 0 and 1
    'AngleDIFF = AngleDIFF + PI
    'AngleDIFF = AngleDIFF / (PI * 2)


End Function

Public Sub InitH()
    Dim I          As Long
    Dim AA         As Single


    NH = frmMain.txtNH

    ReDim H(NH)

    For I = 1 To NH
        With H(I)

            
            ReTarget I, True
            ReColor I


        End With
    Next

End Sub


Public Sub DRAW()
    Dim x1          As Long
    Dim y1          As Long
    Dim x2          As Long
    Dim y2          As Long
    
    Dim I          As Long
    Dim j          As Long

    BitBlt frmMain.PIC.Hdc, 0, 0, frmMain.PIC.Width, frmMain.PIC.Height, frmMain.PIC.Hdc, 0, 0, vbBlackness

    For I = 1 To NH
        With H(I)
            x1 = .X * ZOOM
            y1 = .Y * ZOOM
            
            

            'FastLine frmMain.PIC.Hdc, .X \ 1, .Y \ 1, X \ 1, Y \ 1, 1, .color
            'MyCircle frmMain.PIC.Hdc, .X \ 1, .Y \ 1, r, 1, .color

           'x2 = (.X + .CosA * r) * ZOOM
           'Y2 = (.Y + .SinA * r) * ZOOM
           ' FastLine frmMain.PIC.Hdc, X1, Y1, X2, Y2, 1, .color
           
            MyCircle frmMain.PIC.Hdc, x1, y1, rZOOM, 1, .color

            For j = 0 To TrailLen
                SetPixel frmMain.PIC.Hdc, .TrailX(j), .TrailY(j), .color
            Next

        End With
    Next
    frmMain.PIC.Refresh

End Sub


Public Sub MoveMentOLD()
    Const KK = 11                 ' '15                 '20                 '25 '10

    Const TurnSpeed = 0.07

    Dim I          As Long
    Dim j          As Long
    Dim dX         As Single
    Dim dY         As Single
    Dim VX1        As Single
    Dim VY1        As Single
    Dim vX2        As Single
    Dim vY2        As Single
    Dim rX1        As Single
    Dim rY1        As Single
    Dim rX2        As Single
    Dim rY2        As Single
    Dim D          As Single
    Dim DD         As Single

    Dim A1         As Single
    Dim A2         As Single
    Dim OOangle    As Single
    Dim BearingA   As Single
    Dim ABSBearingA As Single

    Dim NHm1       As Long

    For I = 1 To NH
        H(I).Vmul = 1
    Next

    NHm1 = NH - 1
    For I = 1 To NHm1
        For j = I + 1 To NH

            If Abs(H(I).GridX - H(j).GridX) <= 1 Then
                If Abs(H(I).GridY - H(j).GridY) <= 1 Then


                    dX = H(j).X - H(I).X
                    dY = H(j).Y - H(I).Y

                    D = (dX * dX + dY * dY)
                    If D < Rsq2 Then
                        '                    Stop

                        D = Sqr(D)
                        DD = (MinDIST - D) / D

                        H(I).X = H(I).X - dX * DD
                        H(I).X = H(I).X - dX * DD
                        H(j).X = H(j).X + dX * DD
                        H(j).Y = H(j).Y + dY * DD
                        If H(I).V > H(j).V Then
                            H(I).Vmul = H(I).Vmul * 0.99
                        Else
                            H(j).Vmul = H(j).Vmul * 0.99
                        End If
                    End If
                    OOangle = Atan2(dX, dY)

                    'GoTo skip1
                    BearingA = (AngleDIFF(OOangle, H(I).A))
                    ABSBearingA = Abs(BearingA)
                    If (ABSBearingA) < PIh Then
                        VX1 = Cos(H(I).A) * (H(I).V + 0.5)
                        VY1 = Sin(H(I).A) * (H(I).V + 0.5)
                        vX2 = Cos(H(j).A) * (H(j).V + 0.5)
                        vY2 = Sin(H(j).A) * (H(j).V + 0.5)


                        If ABSBearingA < Amax Then
                            DD = (4 * Sqr(dX * dX + dY * dY) / (Sqr(GridView * GridView + GridView * GridView)))
                            If DD < 1 Then H(I).Vmul = H(I).Vmul * DD
                        End If

                        rX1 = dX + vX2 * KK - VX1 * KK
                        rY1 = dY + vY2 * KK - VY1 * KK
                        A1 = AngleDIFF(Atan2(rX1, rY1), OOangle)

                        If Sgn(A1) = Sgn(AngleDIFF(OOangle, H(I).A)) Then
                            'If AngleDIFF(H(I).A, OOangle) < 0.05 Then
                            H(I).incA = H(I).incA + A1 * TurnSpeed
                            'FastLine frmMain.PIC.Hdc, H(I).X \ 1, H(I).Y \ 1, _
                             H(I).X \ 1 + 20 * Cos(H(I).incA * 10 + H(I).A), H(I).Y \ 1 + 20 * Sin(H(I).incA * 10 + H(I).A), 1, vbYellow
                            'FastLine frmMain.PIC.Hdc, H(I).X \ 1, H(I).Y \ 1, H(J).X \ 1, H(J).Y \ 1, 1, vbRed

                        End If
                    End If

skip1:

                    'GoTo skip2
                    BearingA = (AngleDIFF(OOangle, H(j).A + PI))
                    ABSBearingA = Abs(BearingA)
                    If (ABSBearingA) < PIh Then
                        VX1 = Cos(H(I).A) * (H(I).V + 0.5)
                        VY1 = Sin(H(I).A) * (H(I).V + 0.5)
                        vX2 = Cos(H(j).A) * (H(j).V + 0.5)
                        vY2 = Sin(H(j).A) * (H(j).V + 0.5)

                        If ABSBearingA < Amax Then
                            DD = (4 * Sqr(dX * dX + dY * dY) / (Sqr(GridView * GridView + GridView * GridView)))
                            If DD < 1 Then H(j).Vmul = H(j).Vmul * DD
                        End If

                        rX2 = -dX + VX1 * KK - vX2 * KK
                        rY2 = -dY + VY1 * KK - vY2 * KK
                        A1 = AngleDIFF(Atan2(rX2, rY2), OOangle + PI)

                        If Sgn(A1) = Sgn(AngleDIFF(OOangle + PI, H(j).A)) Then
                            'If AngleDIFF(H(J).A, OOangle) < 0.05 Then

                            H(j).incA = H(j).incA + A1 * TurnSpeed

                            'FastLine frmMain.PIC.Hdc, H(J).X \ 1, H(J).Y \ 1, _
                             H(J).X \ 1 + 20 * Cos(H(J).incA * 10 + H(J).A), H(J).Y \ 1 + 20 * Sin(H(J).incA * 10 + H(J).A), 1, vbYellow
                            'FastLine frmMain.PIC.Hdc, H(I).X \ 1, H(I).Y \ 1, H(J).X \ 1, H(J).Y \ 1, 1, vbRed

                        End If


                    End If

skip2:

                End If
            End If
        Next
    Next
    frmMain.PIC.Refresh

    For I = 1 To NH
        H(I).A = H(I).A + H(I).incA
        H(I).incA = 0
        H(I).V = H(I).V * H(I).Vmul
        H(I).Vmul = 1
    Next

    For I = 1 To NH
        With H(I)

            .X = .X + Cos(.A) * .V
            .Y = .Y + Sin(.A) * .V
            .V = .V * 0.98 + .Vstd * 0.02    '0.005
            .A = .A + AngleDIFF(.A, .ATar) * 0.05    ' 0.035    '0.03    '0.025    '0.02

            dX = .tX - .X
            dY = .tY - .Y
            .ATar = Atan2(dX, dY)
            If (dX * dX + dY * dY) < 100 Then
                'If .X > MaxX * 0.5 Then
                '    .tX = 50
                'Else
                '    .tX = MaxX - 50
                'End If
                ''.tX = Rnd * MaxX
                '.tY = Rnd * MaxY
                '.tY = MaxY / 2 - 300 + Rnd * 600

                Select Case I \ 4 Mod 4
                    Case 0
                        .Y = MaxY / 2 - 200 + Rnd * 400
                        .X = 10
                        .tY = .Y
                        .tX = MaxX - 10
                    Case 1
                        .Y = MaxY / 2 - 200 + Rnd * 400
                        .X = MaxX - 10
                        .tY = .Y
                        .tX = 10
                    Case 2
                        .X = MaxX / 2 - 200 + Rnd * 400
                        .Y = 10
                        .tX = .X
                        .tY = MaxY - 10
                    Case 3
                        .X = MaxX / 2 - 200 + Rnd * 400
                        .Y = MaxY - 10
                        .tX = .X
                        .tY = 10

                End Select

                'If I Mod 2 = 0 Then
                '    .X = MaxX / 2 - 200 + Rnd * 400
                '    .Y = MaxY - 10
                '    .tX = .X
                '    .tY = 10
                'Else
                '    .Y = MaxY / 2 - 200 + Rnd * 400
                '    .X = 10
                '    .tY = .Y
                '    .tX = MaxX - 10
                'End If

                'OOangle = Rnd * PI2
                '.X = MaxX / 2 + Cos(OOangle) * MaxX / 2
                '.Y = MaxY / 2 + Sin(OOangle) * MaxX / 2
                '.tX = MaxX / 2 - Cos(OOangle) * MaxX / 2
                '.tY = MaxY / 2 - Sin(OOangle) * MaxX / 2

            End If

            .GridX = .X \ (GridView)
            .GridY = .Y \ (GridView)
        End With
    Next


End Sub
Public Sub ADDTrail()
    Dim I          As Long
    For I = 1 To NH
        With H(I)
            .TrailX(CT) = .X * ZOOM
            .TrailY(CT) = .Y * ZOOM
        End With
    Next
    CT = CT + 1
    CT = CT Mod (TrailLen + 1)
End Sub
Public Sub MoveMent()

    Const KK       As Single = 15    '11                 ' '15                 '20                 '25 '10

    Const BrakeAmount As Single = 1    '1.025    '1.05

    Const TurnSpeed As Single = 1.25    '2    '1.5

    Dim I          As Long
    Dim j          As Long
    Dim dX         As Single
    Dim dY         As Single
    Dim Vxi        As Single
    Dim Vyi        As Single
    Dim Vxj        As Single
    Dim Vyj        As Single
    Dim rXi        As Single
    Dim rYi        As Single
    Dim rXj        As Single
    Dim rYj        As Single
    Dim D          As Single
    Dim DD         As Single
    Dim D2         As Single

    Dim A1         As Single
    Dim A2         As Single
    Dim OOangle    As Single
    Dim BearingA   As Single
    Dim ABSBearingA As Single

    Dim Xi         As Single
    Dim Yi         As Single
    Dim Xj         As Single
    Dim Yj         As Single

    Dim dX2        As Single
    Dim dY2        As Single
    Dim Bi         As Single
    Dim Bj         As Single
    Dim Bi2        As Single
    Dim Bj2        As Single
    Dim BiD        As Single
    Dim BjD        As Single
    Dim TTi        As Single
    Dim Brake      As Single

    Dim NHm1       As Long

    For I = 1 To NH
        With H(I)
            .Vmul = 1
            .CosA = Cos(H(I).A)
            .SinA = Sin(H(I).A)
            .vX = .V * .CosA
            .vY = .V * .SinA
        End With
    Next

    NHm1 = NH - 1
    For I = 1 To NHm1
        For j = I + 1 To NH

            If Abs(H(I).GridX - H(j).GridX) <= 1 Then
                If Abs(H(I).GridY - H(j).GridY) <= 1 Then



                    dX = H(j).X - H(I).X
                    dY = H(j).Y - H(I).Y
                    'If Abs(dX) <= MinDIST Then
                    '    If Abs(dY) <= MinDIST Then

                            D = (dX * dX + dY * dY)
                            If D < Rsq2 Then
                                D = Sqr(D)
                                DD = (MinDIST - D) / D
                                H(I).X = H(I).X - dX * DD
                                H(I).Y = H(I).Y - dY * DD
                                H(j).X = H(j).X + dX * DD
                                H(j).Y = H(j).Y + dY * DD
                                If H(I).V > H(j).V Then
                                    H(I).Vmul = H(I).Vmul * 0.9999
                                Else
                                    H(j).Vmul = H(j).Vmul * 0.9999
                                End If
                            End If
                    '    End If
                    'End If



                    '-----------------------------------------------------------------------------------

                    Vxi = H(I).vX
                    Vyi = H(I).vY
                    Vxj = H(j).vX
                    Vyj = H(j).vY

                    Xi = H(I).X + Vxi * BrakeAmount
                    Yi = H(I).Y + Vyi * BrakeAmount
                    Xj = H(j).X + Vxj * BrakeAmount
                    Yj = H(j).Y + Vyj * BrakeAmount

                    dX2 = Xj - Xi
                    dY2 = Yj - Yi
                    D = (dX * dX + dY * dY)
                    D2 = (dX2 * dX2 + dY2 * dY2)
                    If D2 < D Then
                        Brake = Sqr(D2) / Sqr(D)
                        TTi = 1 - Brake
                        Brake = Brake ^ 0.6
                        OOangle = Atan2(dX, dY)
                        Bi = (AngleDIFF(OOangle, H(I).A))
                        Bj = (AngleDIFF(OOangle + PI, H(j).A))

                        rXi = dX + Vxj * KK - Vxi * KK
                        rYi = dY + Vyj * KK - Vyi * KK
                        rXj = -dX + Vxi * KK - Vxj * KK
                        rYj = -dY + Vyi * KK - Vyj * KK

                        Bi2 = AngleDIFF(OOangle, Atan2(rXi, rYi))
                        Bj2 = AngleDIFF(OOangle + PI, Atan2(rXj, rYj))
                        BiD = AngleDIFF(Bi, Bi2)
                        BjD = AngleDIFF(Bj, Bj2)


                        If Abs(BiD) < PIh Then
                            H(I).incA = H(I).incA - TurnSpeed * TTi * BiD
                            If Brake < 1 Then H(I).Vmul = H(I).Vmul * Brake


                            'FastLine frmMain.PIC.Hdc, H(I).X \ 1, H(I).Y \ 1, H(J).X \ 1, H(J).Y \ 1, 1, vbRed
                            'FastLine frmMain.PIC.Hdc, H(I).X \ 1, H(I).Y \ 1, H(I).X + rXi, H(I).Y + rYi, 1, vbGreen
                            'frmMain.PIC.Refresh
                        End If

                        If Abs(BjD) < PIh Then    '
                            H(j).incA = H(j).incA - TurnSpeed * TTi * BjD
                            If Brake < 1 Then H(j).Vmul = H(j).Vmul * Brake

                            'FastLine frmMain.PIC.Hdc, H(I).X \ 1, H(I).Y \ 1, H(J).X \ 1, H(J).Y \ 1, 1, vbRed
                            'FastLine frmMain.PIC.Hdc, H(J).X \ 1, H(J).Y \ 1, H(J).X + rXj, H(J).Y + rYj, 1, vbYellow
                            'frmMain.PIC.Refresh
                        End If

                    End If

                    '-----------------------------------------------------------------------------------

                End If
            End If
        Next
    Next
    frmMain.PIC.Refresh

    For I = 1 To NH
        H(I).A = H(I).A + H(I).incA
        H(I).incA = 0
        H(I).V = H(I).V * H(I).Vmul
        H(I).Vmul = 1
    Next

    For I = 1 To NH
        With H(I)

            .X = .X + Cos(.A) * .V
            .Y = .Y + Sin(.A) * .V
            .V = .V * 0.98 + .Vstd * 0.02    '0.005
            .A = .A + AngleDIFF(.A, .ATar) * 0.035    ' 0.035    '0.03    '0.025    '0.02

            dX = .tX - .X
            dY = .tY - .Y
            .ATar = Atan2(dX, dY)
            If (dX * dX + dY * dY) < 100 Then

                ReTarget I, True
            End If

            .GridX = .X \ GridView
            .GridY = .Y \ GridView
          
            '            Stop

        End With
    Next


End Sub
Public Sub ReTarget(I As Long, StartToo As Boolean)
    Dim AA         As Single
    Dim X          As Single
    Dim Y          As Single

    With H(I)
        Select Case frmMain.cmbTargetMode.ListIndex
            Case 0
                AA = Rnd * PI2
                If StartToo Then .X = MaxX / 2 + Cos(AA) * MaxX / 2
                If StartToo Then .Y = MaxY / 2 + Sin(AA) * MaxX / 2
                .tX = MaxX / 2 - Cos(AA) * MaxX / 2
                .tY = MaxY / 2 - Sin(AA) * MaxX / 2
                .A = AA + PI
                .Vstd = 1 + 0.2 * Rnd    ' Rnd * 0.2

                ReColor I
            Case 1
                If I Mod 2 = 0 Then
                    Y = MaxY / 2 - 200 + Rnd * 400
                    X = MaxX - 10
                    If StartToo Then .Y = Y: .X = X
                    .tY = Y
                    .tX = 10
                Else
                    Y = MaxY / 2 - 200 + Rnd * 400
                    X = 10
                    If StartToo Then .Y = Y: .X = X
                    .tY = Y
                    .tX = MaxX - 10
                End If
            Case 2
                If I Mod 2 = 0 Then
                    X = MaxX / 2 - 200 + Rnd * 400
                    Y = MaxY - 10
                    If StartToo Then .Y = Y: .X = X
                    .tX = X
                    .tY = 10
                Else
                    Y = MaxY / 2 - 200 + Rnd * 400
                    X = 10
                    If StartToo Then .Y = Y: .X = X
                    .tY = Y
                    .tX = MaxX - 10
                End If
            Case 3
                Select Case I \ 4 Mod 4
                    Case 0
                        Y = MaxY / 2 - 200 + Rnd * 400
                        X = 10
                        .tY = Y
                        .tX = MaxX - 10
                    Case 1
                        Y = MaxY / 2 - 200 + Rnd * 400
                        X = MaxX - 10
                        .tY = Y
                        .tX = 10
                    Case 2
                        X = MaxX / 2 - 200 + Rnd * 400
                        Y = 10
                        .tX = X
                        .tY = MaxY - 10
                    Case 3
                        X = MaxX / 2 - 200 + Rnd * 400
                        Y = MaxY - 10
                        .tX = X
                        .tY = 10
                End Select
                 If StartToo Then .Y = Y: .X = X
            Case 4
                Select Case I \ 4 Mod 4
                    Case 0
                        Y = MaxY * Rnd
                        X = 10
                        .tY = Y
                        .tX = MaxX - 10
                    Case 1
                        Y = MaxY * Rnd
                        X = MaxX - 10
                        .tY = Y
                        .tX = 10
                    Case 2
                        X = MaxX * Rnd
                        Y = 10
                        .tX = X
                        .tY = MaxY - 10
                    Case 3
                        X = MaxX * Rnd
                        Y = MaxY - 10
                        .tX = X
                        .tY = 10
                End Select
                If StartToo Then .Y = Y: .X = X


        End Select
    End With

End Sub
Public Sub ReColor(I As Long)
    Dim AA         As Single

    With H(I)
        Select Case frmMain.cmbTargetMode.ListIndex
            Case 0
                AA = Atan2(.X - .tX, .Y - .tY)
                If AA < 0 Then AA = AA + PI2
                AA = 64 + 255 * AA / PI2
                .color = RGB(AA, (64 + AA) Mod 255, (127 + AA) Mod 255)

            Case 1, 2
                If I Mod 2 = 0 Then
                    .color = RGB(200 + Rnd * 150, 120, 120)
                Else
                    .color = RGB(120, 200 + Rnd * 150, 120)
                End If

            Case 3, 4
                Select Case I \ 4 Mod 4
                    Case 0
                        .color = RGB(0, 255, 0)
                    Case 1
                        .color = RGB(255, 0, 0)

                    Case 2
                        .color = RGB(255, 255, 0)
                    Case 3
                        .color = RGB(255, 0, 255)
                End Select
        End Select


    End With

End Sub
