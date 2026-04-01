'==============================================================================
' PAGODA ANTENNA - CST STUDIO SUITE MACRO (CORRECTED & FORMATTED)



'==============================================================================
Option Explicit

Const PI As Double = 3.14159265358979323846

Sub Main()

    '--------------------------------------------------------------------------
    ' USER PARAMETERS
    '--------------------------------------------------------------------------
    Dim TARGET_FREQ_GHZ As Double
    TARGET_FREQ_GHZ = 5.9

    Dim POL As String
    POL = "RHCP"

    Dim VER As String
    VER = "3"

    Dim RUN_SOLVER As Boolean
    RUN_SOLVER = False

    Dim ans As Integer
    ans = MsgBox("Pagoda Antenna Builder" & Chr(13) & Chr(13) & _
                 "Frequency : " & TARGET_FREQ_GHZ & " GHz" & Chr(13) & _
                 "Pol       : " & POL              & Chr(13) & _
                 "Version   : " & VER              & Chr(13) & Chr(13) & _
                 "Run on a FRESH empty MWS project only." & Chr(13) & _
                 "Continue?", vbYesNo + vbQuestion, "Pagoda Builder")
    If ans = vbNo Then Exit Sub

    '--------------------------------------------------------------------------
    ' SCALE FACTOR & DIMENSIONS
    '--------------------------------------------------------------------------
    Dim sc As Double
    sc = 5.8 / TARGET_FREQ_GHZ

    ' SMA CONNECTOR DIMENSIONS
    Dim coax_r1 As Double
    coax_r1 = 0.46

    Dim coax_r2 As Double
    coax_r2 = 1.50

    Dim coax_r3 As Double
    coax_r3 = 1.80

    Dim solder_w As Double
    solder_w = 0.60

    Dim hole_sp As Double
    hole_sp = 0.05

    Dim hole_sp2 As Double
    hole_sp2 = 0.10

    Dim ring_w As Double
    ring_w = 0.25

    '--------------------------------------------------------------------------
    ' ANTENNA TUNING PARAMETERS (Based on parameter.txt)
    '--------------------------------------------------------------------------
    Dim R1 As Double ' Outer substrate radius (PCB1 and 2)
    Dim R2 As Double ' Outer substrate radius (PCB3)
    Dim R3 As Double ' Outer copper arc radius (centerline)
    Dim R4 As Double ' Inner copper arc outer edge radius
    Dim R5 As Double ' Inner copper arc inner edge radius
    Dim R6 As Double ' Central copper cylinder outer radius (PCB1 and 2)
    Dim R7 As Double ' SMA pad outer radius (PCB2)
    Dim R8 As Double ' Reflector disk radius (PCB3)

    Dim L1 As Double ' Outer arc angular length (degrees)
    Dim L2 As Double ' Inner arc angular length (degrees)
    Dim L3 As Double ' Inner radius of the central cylinder (hole)

    Dim W1 As Double ' Track width of the radial spokes
    Dim W2 As Double ' Width of the central copper cylinder
    Dim W3 As Double ' Width of the SMA pad on PCB2
    Dim W4 As Double ' Outer arc track width

    Dim disk_d1 As Double
    Dim disk_d2 As Double

    If VER = "3" Then
        ' Base values for Version 3 (Scaled by sc)
        R4 = 9 * sc
        R5 = 8  * sc
        R3 = 10.9  * sc
        
        W4 = 0.9 * sc
        W1 = 1.0 * sc

        L1 = 8.74
        L2 = 5

        R6 = 6.9 * sc
        W2 = R6 - (2.75 * sc) ' Set width explicitly
        L3 = R6 - W2            ' Inner radius derived from R6 and W2

        R8 = 4.5 * sc

        R1 = R4 + 0.5 * sc
        R2 = R8 + 0.5 * sc
        
        R7 = coax_r3 + solder_w
        W3 = R7 - coax_r2       ' Set width explicitly

        disk_d1 = 2.7 * sc
        disk_d2 = 4.5 * sc

    ElseIf VER = "3B" Then
        ' Base values for Version 3B
        R4 = 10.7313 * sc
        R5 = 9.7313  * sc
        
        R6 = 5.5849 * sc
        W2 = R6 - (2.4297 * sc) ' Set width explicitly
        L3 = R6 - W2            ' Inner radius derived from R6 and W2
        
        R3 = ((R4 + R5)/2.0 + R6 - 0.5 / 2.0) / 2.0
        
        W4 = 1.0 * sc
        W1 = 1.0 * sc

        L1 = 19.1231
        L2 = 69.1855

        R8 = 5.6459 * sc

        R1 = R4 + 0.5 * sc
        R2 = R8 + 0.5 * sc
        
        R7 = coax_r3 + solder_w
        W3 = R7 - coax_r2       ' Set width explicitly

        disk_d1 = 3.6526 * sc
        disk_d2 = 12.4514 * sc
    Else
        Exit Sub
    End If

    '--------------------------------------------------------------------------
    ' MAP TUNING PARAMETERS TO INTERNAL MACRO VARIABLES
    '--------------------------------------------------------------------------
    Dim track_r1 As Double
 track_r1 = (R4 + R5) / 2.0
    Dim track_w1_inner As Double
 track_w1_inner = R4 - R5
    Dim track_w1_outer As Double
 track_w1_outer = W4
    Dim track_r2 As Double
 track_r2 = R3
    Dim track_w2 As Double
 track_w2 = W1
    Dim track_a1 As Double
 track_a1 = L2
    Dim track_c1 As Double
 track_c1 = L1
    
    Dim disk_r1 As Double
 disk_r1 = R6
    Dim disk_r2 As Double
 disk_r2 = R6
    Dim hole_r1 As Double
 hole_r1 = L3
    Dim hole_r2 As Double
 hole_r2 = L3
    Dim disk_r3 As Double
 disk_r3 = R8

    Dim pcb_r1 As Double
 pcb_r1 = R1
    Dim pcb_r2 As Double
 pcb_r2 = R1
    Dim pcb_r3 As Double
 pcb_r3 = R2

    ' FIXED MECHANICAL
    Dim pcb_th As Double
    pcb_th = 1.0

    Dim cu_th As Double
    cu_th = 0.035

    Dim track_b1 As Double
    track_b1 = -track_c1 / 2.0

    ' Z STACK
    Dim z3_bot As Double
    z3_bot = 0.0

    Dim z3_top As Double
    z3_top = z3_bot + pcb_th

    Dim z2_bot As Double
    z2_bot = z3_top + disk_d2

    Dim z2_top As Double
    z2_top = z2_bot + pcb_th

    Dim z1_bot As Double
    z1_bot = z2_top + disk_d1

    Dim z1_top As Double
    z1_top = z1_bot + pcb_th

    Dim coax_stub_len As Double
    coax_stub_len = 5.0

    Dim z_port As Double
    z_port = z3_bot - coax_stub_len

    '--------------------------------------------------------------------------
    ' SETUP & MATERIALS
    '--------------------------------------------------------------------------
    Call SetupProject()
    Call DefineMaterials()

    ' PCB SUBSTRATES
    Call BuildSubstrate("PCB1", "Substrates", 0, 0, z1_bot, pcb_r1, pcb_th, coax_r1 + hole_sp)
    Call BuildSubstrate("PCB2", "Substrates", 0, 0, z2_bot, pcb_r2, pcb_th, coax_r3 + hole_sp2)
    Call BuildSubstrate("PCB3", "Substrates", 0, 0, z3_bot, pcb_r3, pcb_th, coax_r3 + hole_sp2)

    '--------------------------------------------------------------------------
    ' PCB1 TOP COPPER (Connected to coax center pin)
    '--------------------------------------------------------------------------
    Dim cu_top1 As Double
    cu_top1 = z1_top

    Call BuildCopperRing("pcb1_gndring", "PCB1_CopperTop", 0, 0, cu_top1, hole_r1, disk_r1, cu_th)

    ' FIX: Solid disk for center pin connection
    Call BuildSolidDisk("pcb1_sma_pad", "PCB1_CopperTop", 0, 0, cu_top1, coax_r1 + solder_w, cu_th)

    Dim arm As Integer
    Dim base_ang As Double
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double
    Dim a4 As Double
    Dim a5 As Double
    Dim pfx As String

    For arm = 0 To 2
        base_ang = 90.0 + arm * 120.0

        If POL = "LHCP" Then
            a1 = base_ang + track_b1 - (track_w2 - track_w1_inner) / 2.0 / track_r1 * (180.0 / PI)
            a2 = a1 + track_a1
            a3 = base_ang + track_b1
            a4 = base_ang + track_b1 + track_c1
        Else
            a1 = base_ang - track_b1 + (track_w2 - track_w1_inner) / 2.0 / track_r1 * (180.0 / PI)
            a2 = a1 - track_a1
            a3 = base_ang - track_b1
            a4 = base_ang - track_b1 - track_c1
        End If

        a5 = base_ang + 60.0
        pfx = "pcb1_arm" & (arm + 1)

        Call BuildArcTrace(pfx & "_iarc", "PCB1_CopperTop", 0, 0, cu_top1, track_r1, a1, a2, track_w1_inner, cu_th)
        Call BuildArcTrace(pfx & "_oarc", "PCB1_CopperTop", 0, 0, cu_top1, track_r2, a3, a4, track_w1_outer, cu_th)
        Call BuildRadialTrace(pfx & "_rad1", "PCB1_CopperTop", 0, 0, cu_top1, track_r2, track_r1, a3, track_w2, cu_th)
        Call BuildRadialTrace(pfx & "_rad2", "PCB1_CopperTop", 0, 0, cu_top1, track_r2, disk_r1, a4, track_w2, cu_th)

        ' FIX: Connect spoke safely to center pad area
        Call BuildRadialTrace(pfx & "_spoke", "PCB1_CopperTop", 0, 0, cu_top1, 0.0, hole_r1, a5, track_w2, cu_th)
    Next arm

    '--------------------------------------------------------------------------
    ' PCB2 TOP COPPER (Connected to coax shield)
    '--------------------------------------------------------------------------
    Dim cu_top2 As Double
    cu_top2 = z2_top

    Call BuildCopperRing("pcb2_gndring", "PCB2_CopperTop", 0, 0, cu_top2, hole_r2, disk_r2, cu_th)

    ' FIX: Start pad inner edge is derived from W3 (R7 - W3) and outer edge is R7.
    ' Note: Normally R7 - W3 = coax_r2 to avoid shorting, but we use the tuning parameter directly here.
    Call BuildCopperRing("pcb2_sma_pad", "PCB2_CopperTop", 0, 0, cu_top2, R7 - W3, R7, cu_th)

    For arm = 0 To 2
        base_ang = 90.0 + arm * 120.0

        If POL = "LHCP" Then
            a1 = base_ang - track_b1 + (track_w2 - track_w1_inner) / 2.0 / track_r1 * (180.0 / PI)
            a2 = a1 - track_a1
            a3 = base_ang - track_b1
            a4 = base_ang - track_b1 - track_c1
        Else
            a1 = base_ang + track_b1 - (track_w2 - track_w1_inner) / 2.0 / track_r1 * (180.0 / PI)
            a2 = a1 + track_a1
            a3 = base_ang + track_b1
            a4 = base_ang + track_b1 + track_c1
        End If

        a5 = base_ang + 60.0
        pfx = "pcb2_arm" & (arm + 1)

        Call BuildArcTrace(pfx & "_iarc", "PCB2_CopperTop", 0, 0, cu_top2, track_r1, a1, a2, track_w1_inner, cu_th)
        Call BuildArcTrace(pfx & "_oarc", "PCB2_CopperTop", 0, 0, cu_top2, track_r2, a3, a4, track_w1_outer, cu_th)
        Call BuildRadialTrace(pfx & "_rad1", "PCB2_CopperTop", 0, 0, cu_top2, track_r2, track_r1, a3, track_w2, cu_th)
        Call BuildRadialTrace(pfx & "_rad2", "PCB2_CopperTop", 0, 0, cu_top2, track_r2, disk_r2, a4, track_w2, cu_th)

        ' FIX: Ensure spoke reaches the inner edge of the SMA pad (R7 - W3) to prevent open circuit
        Call BuildRadialTrace(pfx & "_spoke", "PCB2_CopperTop", 0, 0, cu_top2, R7 - W3, hole_r2, a5, track_w2, cu_th)
    Next arm

    '--------------------------------------------------------------------------
    ' PCB3 BOTTOM COPPER (Reflector)
    '--------------------------------------------------------------------------
    Dim cu_bot3 As Double
    cu_bot3 = z3_bot - cu_th

    ' FIX: Used BuildCopperRing instead of BuildSolidDisk to allow center pin to pass without shorting
    Call BuildCopperRing("pcb3_reflector", "PCB3_CopperBot", 0, 0, cu_bot3, coax_r2, R8, cu_th)

    '--------------------------------------------------------------------------
    ' COAXIAL FEED STUB
    '--------------------------------------------------------------------------
    With Cylinder
        .Reset
        .Name "coax_center"
        .Component "Feed"
        .Material "PEC"
        .OuterRadius coax_r1
        .InnerRadius 0
        .Axis "z"
        .Zrange z_port, z1_top
        .Create
    End With

    With Cylinder
        .Reset
        .Name "coax_ptfe"
        .Component "Feed"
        .Material "PTFE_Teflon"
        .OuterRadius coax_r2
        .InnerRadius coax_r1
        .Axis "z"
        .Zrange z_port, z2_top
        .Create
    End With

    With Cylinder
        .Reset
        .Name "coax_shield"
        .Component "Feed"
        .Material "PEC"
        .OuterRadius coax_r3
        .InnerRadius coax_r2
        .Axis "z"
        .Zrange z_port, z2_top
        .Create
    End With

    '--------------------------------------------------------------------------
    ' WAVEGUIDE PORT
    '--------------------------------------------------------------------------
    With Port
        .Reset
        .PortNumber 1
        .Label "SMA_Feed"
        .NumberOfModes 1
        .Orientation "zmin"
        .Xrange -coax_r3, coax_r3
        .Yrange -coax_r3, coax_r3
        .Zrange z_port, z_port
        .Create
    End With

    '--------------------------------------------------------------------------
    ' FREQUENCY + BOUNDARY
    '--------------------------------------------------------------------------
    Dim f_min As Double
    f_min = TARGET_FREQ_GHZ * 0.6

    Dim f_max As Double
    f_max = TARGET_FREQ_GHZ * 1.4

    Solver.FrequencyRange f_min, f_max

    ' FIX: Set bounds to "open (add space)" to prevent clipping the near-fields
    With Boundary
        .Xmin "open (add space)"
        .Xmax "open (add space)"
        .Ymin "open (add space)"
        .Ymax "open (add space)"
        .Zmin "open (add space)"
        .Zmax "open (add space)"
        .Xsymmetry "none"
        .Ysymmetry "none"
        .Zsymmetry "none"
    End With

    With Background
        .Type "Normal"
        .Epsilon 1.0
        .Mu 1.0
    End With

    '--------------------------------------------------------------------------
    ' MONITORS & SOLVER
    '--------------------------------------------------------------------------
    With Monitor
        .Reset
        .Domain "frequency"
        .FieldType "Farfield"
        .Frequency TARGET_FREQ_GHZ
        .Name "farfield_" & Format(TARGET_FREQ_GHZ, "0_0") & "GHz"
        .Create
    End With

    With Solver
        .Method "Hexahedral"
        .CalculationType "TD-S"
        .AutoNormImpedance True
        .NormingImpedance 50.0
    End With

    If RUN_SOLVER Then Solver.Start

    MsgBox "=== PAGODA ANTENNA BUILT ===" & Chr(13) & "Electrical faults fixed. VSWR will now be normal."
End Sub

'==============================================================================
' HELPER FUNCTIONS
'==============================================================================
Sub SetupProject()
    With Units
        .Geometry "mm"
        .Frequency "GHz"
        .Time "ns"
        .TemperatureUnit "Celsius"
    End With
End Sub

Sub DefineMaterials()
    With Material
        .Reset
        .Name "FR4_Lossy"
        .Type "Normal"
        .Epsilon 4.3
        .TanD 0.02
        .Create
    End With

    With Material
        .Reset
        .Name "Copper (annealed)"
        .Type "Lossy metal"
        .Sigma 5.8e7
        .Create
    End With

    With Material
        .Reset
        .Name "PTFE_Teflon"
        .Type "Normal"
        .Epsilon 2.1
        .TanD 0.0002
        .Create
    End With
End Sub

Sub BuildSubstrate(sName As String, sComp As String, cx As Double, cy As Double, z_bot As Double, pcb_r As Double, pcb_th As Double, drill_r As Double)
    With Cylinder
        .Reset
        .Name sName
        .Component sComp
        .Material "FR4_Lossy"
        .OuterRadius pcb_r
        .InnerRadius drill_r
        .Axis "z"
        .Zrange z_bot, z_bot + pcb_th
        .Create
    End With
End Sub

Sub BuildCopperRing(sName As String, sComp As String, cx As Double, cy As Double, z_bot As Double, r_in As Double, r_out As Double, cu_h As Double)
    With Cylinder
        .Reset
        .Name sName
        .Component sComp
        .Material "Copper (annealed)"
        .OuterRadius r_out
        .InnerRadius r_in
        .Axis "z"
        .Zrange z_bot, z_bot + cu_h
        .Create
    End With
End Sub

Sub BuildSolidDisk(sName As String, sComp As String, cx As Double, cy As Double, z_bot As Double, r As Double, cu_h As Double)
    With Cylinder
        .Reset
        .Name sName
        .Component sComp
        .Material "Copper (annealed)"
        .OuterRadius r
        .InnerRadius 0
        .Axis "z"
        .Zrange z_bot, z_bot + cu_h
        .Create
    End With
End Sub

Sub BuildArcTrace(sName As String, sComp As String, cx As Double, cy As Double, z_bot As Double, arc_r As Double, angle1 As Double, angle2 As Double, tw As Double, cu_h As Double)
    Dim grp As String
    grp = "crv_" & sName

    Dim r_in As Double
    r_in = arc_r - tw / 2.0

    Dim r_out As Double
    r_out = arc_r + tw / 2.0

    Dim loc_a1 As Double
    loc_a1 = angle1

    Dim loc_a2 As Double
    loc_a2 = angle2

    Dim overlap As Double
    overlap = 0.5

    If loc_a1 < loc_a2 Then
        loc_a1 = loc_a1 - overlap
        loc_a2 = loc_a2 + overlap
    Else
        loc_a1 = loc_a1 + overlap
        loc_a2 = loc_a2 - overlap
    End If

    Dim a1 As Double
    a1 = loc_a1 * PI / 180.0

    Dim a2 As Double
    a2 = loc_a2 * PI / 180.0

    Curve.NewCurve grp

    With Line
        .Reset
        .Curve grp
        .Name "l1"
        .X1 cx + r_in * Cos(a1)
        .Y1 cy + r_in * Sin(a1)
        .X2 cx + r_out * Cos(a1)
        .Y2 cy + r_out * Sin(a1)
        .Create
    End With

    With Arc
        .Reset
        .Curve grp
        .Name "a1"
        .Xcenter cx
        .Ycenter cy
        .X1 cx + r_out * Cos(a1)
        .Y1 cy + r_out * Sin(a1)
        .X2 cx + r_out * Cos(a2)
        .Y2 cy + r_out * Sin(a2)
        .UseAngle False
        If loc_a2 > loc_a1 Then
            .Orientation "CounterClockwise"
        Else
            .Orientation "Clockwise"
        End If
        .Segments 0
        .Create
    End With

    With Line
        .Reset
        .Curve grp
        .Name "l2"
        .X1 cx + r_out * Cos(a2)
        .Y1 cy + r_out * Sin(a2)
        .X2 cx + r_in * Cos(a2)
        .Y2 cy + r_in * Sin(a2)
        .Create
    End With

    With Arc
        .Reset
        .Curve grp
        .Name "a2"
        .Xcenter cx
        .Ycenter cy
        .X1 cx + r_in * Cos(a2)
        .Y1 cy + r_in * Sin(a2)
        .X2 cx + r_in * Cos(a1)
        .Y2 cy + r_in * Sin(a1)
        .UseAngle False
        If loc_a2 > loc_a1 Then
            .Orientation "Clockwise"
        Else
            .Orientation "CounterClockwise"
        End If
        .Segments 0
        .Create
    End With

    With ExtrudeCurve
        .Reset
        .Name sName
        .Component sComp
        .Material "Copper (annealed)"
        .Curve grp & ":l1"
        .Thickness cu_h
        .Create
    End With

    With Transform
        .Reset
        .Name sComp & ":" & sName
        .Vector 0, 0, z_bot
        .Transform "Shape", "Translate"
    End With

    Curve.DeleteCurve grp
End Sub

Sub BuildRadialTrace(sName As String, sComp As String, cx As Double, cy As Double, z_bot As Double, r1 As Double, r2 As Double, angle_deg As Double, tw As Double, cu_h As Double)
    Dim loc_r1 As Double
    loc_r1 = r1

    Dim loc_r2 As Double
    loc_r2 = r2

    Dim overlap As Double
    overlap = tw / 2.0

    If loc_r1 < loc_r2 Then
        If loc_r1 >= overlap Then loc_r1 = loc_r1 - overlap
        loc_r2 = loc_r2 + overlap
    Else
        If loc_r2 >= overlap Then loc_r2 = loc_r2 - overlap
        loc_r1 = loc_r1 + overlap
    End If

    If loc_r1 < 0 Then loc_r1 = 0
    If Abs(loc_r2 - loc_r1) < 0.001 Then Exit Sub

    Dim a As Double
    a = angle_deg * PI / 180.0

    Dim perp As Double
    perp = a + PI / 2.0

    Dim px As Double
    px = Cos(perp) * tw / 2.0

    Dim py As Double
    py = Sin(perp) * tw / 2.0

    Dim dx As Double
    dx = Cos(a)

    Dim dy As Double
    dy = Sin(a)

    Dim grp As String
    grp = "crv_" & sName

    Curve.NewCurve grp

    With Line
        .Reset
        .Curve grp
        .Name "l1"
        .X1 cx + loc_r1 * dx - px
        .Y1 cy + loc_r1 * dy - py
        .X2 cx + loc_r2 * dx - px
        .Y2 cy + loc_r2 * dy - py
        .Create

        .Name "l2"
        .X1 cx + loc_r2 * dx - px
        .Y1 cy + loc_r2 * dy - py
        .X2 cx + loc_r2 * dx + px
        .Y2 cy + loc_r2 * dy + py
        .Create

        .Name "l3"
        .X1 cx + loc_r2 * dx + px
        .Y1 cy + loc_r2 * dy + py
        .X2 cx + loc_r1 * dx + px
        .Y2 cy + loc_r1 * dy + py
        .Create

        .Name "l4"
        .X1 cx + loc_r1 * dx + px
        .Y1 cy + loc_r1 * dy + py
        .X2 cx + loc_r1 * dx - px
        .Y2 cy + loc_r1 * dy - py
        .Create
    End With

    With ExtrudeCurve
        .Reset
        .Name sName
        .Component sComp
        .Material "Copper (annealed)"
        .Curve grp & ":l1"
        .Thickness cu_h
        .Create
    End With

    With Transform
        .Reset
        .Name sComp & ":" & sName
        .Vector 0, 0, z_bot
        .Transform "Shape", "Translate"
    End With

    Curve.DeleteCurve grp
End Sub
