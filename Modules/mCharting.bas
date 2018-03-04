Attribute VB_Name = "mCharting"
Option Explicit

'Input: x(1:N,1:n_dimension), n_dimension data array with N observations
'       x_class(1:N) class label of each observation
'       factor_label(1:n_dimension), name of each dimension
'Output: data printed to vRng in chartable format
'        factor_legend, if supplied it returns text label of each factor in chartable format
Sub Parallel_Coordinates_Chart(vRng As Range, x As Variant, x_class As Variant, _
        factor_label As Variant, Optional reorder As Boolean = False, Optional factor_legend As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long, n_class As Long
Dim vArr As Variant, class_list As Variant
Dim x_class_idx() As Long, jArr() As Long, ordering() As Long, class_size() As Long
Dim x_shift() As Double, x_scale() As Double, x_correl() As Double
Dim strfactors() As String
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    Call modMath.Normalize_x(x, x_shift, x_scale, "MINMAX") 'Normalize raw data
    Call modMath.Unique_Items(x_class, x_class_idx, class_list, n_class, class_size)
    Erase x_shift, x_scale
    
    ReDim strfactors(1 To n_dimension)
    For j = 1 To n_dimension
        strfactors(j) = factor_label(j)
    Next j
    
    'Find the size of the largest class so we can determine output size
    k = 0
    For i = 1 To n_class
        If class_size(i) > k Then k = class_size(i)
    Next i
    
    'Allocate row and column headings
    ReDim vArr(0 To k * (n_dimension + 1), 0 To n_class)
    For i = 1 To n_class
        vArr(0, i) = class_list(i)
    Next i
    For i = 1 To k
        n = (i - 1) * (n_dimension + 1)
        For j = 1 To n_dimension
            vArr(n + j, 0) = j
        Next j
    Next i
    
    If reorder = False Then
        ordering = modMath.index_array(1, n_dimension)
    Else
        Dim tree1 As cHierarchical
        Set tree1 = New cHierarchical
        With tree1
            x_correl = modMath.Correl_Matrix(x)
            For i = 1 To n_dimension
                For j = 1 To n_dimension
                    x_correl(i, j) = Sqr(1 - x_correl(i, j) ^ 2)
                Next j
            Next i
            Call .NNChainLinkage(strfactors, x_correl)
            Call .MOLO_Ordering
            ordering = .leaf_order
            Call .Reset
        End With
        Set tree1 = Nothing
        Erase x_correl
    End If
    
    ReDim jArr(1 To n_class)
    For i = 1 To n_raw
        k = x_class_idx(i)
        jArr(k) = jArr(k) + 1
        n = (jArr(k) - 1) * (n_dimension + 1)
        For j = 1 To n_dimension
            vArr(n + j, k) = x(i, ordering(j))
        Next j
    Next i
    If IsMissing(factor_legend) = False Then
        ReDim factor_legend(1 To n_dimension, 1 To 3)
        For j = 1 To n_dimension
            factor_legend(j, 1) = j
            factor_legend(j, 2) = -0.1
            factor_legend(j, 3) = strfactors(ordering(j))
        Next j
    End If
    Erase ordering
        
    With vRng
        m = UBound(vArr, 1)
        n = UBound(vArr, 2)
        Range(.Offset(0, 0), .Offset(m, n)).Value = vArr
    End With
    Erase vArr, x_class_idx, jArr, class_list
End Sub


'Input: s, chart series, e.g. mywkbk.Sheets("Sheet2").ChartObjects("NETWORK_CHART").Chart.SeriesCollection(1)
'       vArr(1 to N), reference values to size each dot
Sub Label_scatter_plot(s As Series, vArr As Variant)
Dim i As Long, n As Long
    n = UBound(vArr, 1)
    With s
        If .HasDataLabels = True Then
            .DataLabels.Delete
            .ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False
        Else
            .ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False
        End If
        For i = 1 To n
            .Points(i).DataLabel.Text = vArr(i)
        Next i
    End With
End Sub

Sub Resize_scatter_plot(s As Series, vArr As Variant, Optional min_size As Long = 2, _
    Optional max_size As Long = 30, Optional isReverse As Boolean = False)
Dim i As Long, n As Long
Dim tmp_x As Double, tmp_y As Double, tmp_rng As Double
    n = UBound(vArr, 1)
    If min_size < 2 Then min_size = 2
    With s
        tmp_x = Exp(70)
        tmp_y = -tmp_x
        For i = 1 To n
            If vArr(i) < tmp_x Then tmp_x = vArr(i)
            If vArr(i) > tmp_y Then tmp_y = vArr(i)
        Next i
        tmp_rng = tmp_y - tmp_x
        If tmp_rng > 0 Then
            If isReverse = False Then
                For i = 1 To n
                    .Points(i).MarkerSize = Round(min_size + max_size * (vArr(i) - tmp_x) / tmp_rng, 0)
                Next i
            Else
                For i = 1 To n
                    .Points(i).MarkerSize = Round(min_size + max_size * (tmp_y - vArr(i)) / tmp_rng, 0)
                Next i
            End If
        End If
    End With
End Sub

Sub Color_scatter_plot(s As Series, vArr As Variant, Optional grayscale As Long = 0, _
        Optional isReverse As Boolean = False)
Dim i As Long, j As Long, n As Long, vR As Long, vG As Long, vB As Long
Dim tmp_max As Double, tmp_min As Double, tmp_x As Double
Dim vtmp As Variant
n = UBound(vArr)
tmp_max = -Exp(70)
tmp_min = Exp(70)
For Each vtmp In vArr
    If vtmp > tmp_max Then tmp_max = vtmp
    If vtmp < tmp_min Then tmp_min = vtmp
Next vtmp
With s
    For i = 1 To n
        If isReverse = False Then
            tmp_x = (vArr(i) - tmp_min) / (tmp_max - tmp_min)
        Else
            tmp_x = (tmp_max - vArr(i)) / (tmp_max - tmp_min)
        End If
        If grayscale = 0 Then
            Call Color_Scale(tmp_x, vR, vG, vB)
        Else
            Call Gray_Scale(tmp_x, vR)
            vG = vR: vB = vR
        End If
        .Points(i).Format.Fill.ForeColor.RGB = RGB(vR, vG, vB)
    Next i
End With
End Sub

'Color Scheme: Min=Red, Med=Yellow, Max=Blue
'Input: x is a real number between 0 and 1
'Output: vR,vG,vB are integers from 0 to 255
Private Sub Color_Scale(x As Double, vR As Long, vG As Long, vB As Long)
If x <= 0.5 Then
    vR = 255
    vG = Int(510 * x)
    vB = 0
Else
    vR = Int(-510 * (x - 1))
    vG = vR
    vB = Int(510 * x - 255)
End If
End Sub

'Color SchemeL Min=white, Max=Black
'Input: x is a real number between 0 and 1
'Output: vR are integers from 0 to 255
Private Sub Gray_Scale(x As Double, vR As Long)
vR = Int(-255 * (1 - Exp(-(x - 1) * 5)) / (1 + Exp(-(x - 1) * 5)))
End Sub


Sub Test_Projection()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim xyz As Variant, vArr As Variant, uArr As Variant, vGrid As Variant
Dim mywkbk As Workbook

Set mywkbk = ActiveWorkbook

With mywkbk.Sheets("Sheet6")
    .Range("I6:P100000").Clear
    
'    n = .Range("A100000").End(xlUp).Row - 5
'    xyz = .Range("A6").Resize(n, 3).Value
'    vArr = Projection3D(xyz, uArr, True, , , , 2)
'    .Range("I6").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
'
'    Erase xyz, vArr
'    n = .Range("E100000").End(xlUp).Row - 5
'    xyz = .Range("E6").Resize(n, 3).Value
'    vArr = Projection3D(xyz, uArr, , True)
'    .Range("L6").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr

    n = .Range("E100000").End(xlUp).Row - 5
    xyz = .Range("E6").Resize(n, 3).Value
    vArr = Projection3D(xyz, uArr, True, , 1, 1, 3, , , , , , , vGrid, True)
    .Range("L6").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
    .Range("O6").Resize(UBound(vGrid, 1), UBound(vGrid, 2)).Value = vGrid


End With

Set mywkbk = Nothing




End Sub


Function Projection3D(xyz As Variant, Optional vTransform As Variant, Optional save_transform As Boolean = False, _
            Optional use_transform As Boolean = False, _
            Optional cam_x As Double = 1, Optional cam_y As Double = 1, Optional cam_z As Double = 1, _
            Optional theta_x As Double = 0, Optional theta_y As Double = 0, Optional theta_z As Double = 0, _
            Optional pan_x As Double = 0, Optional pan_y As Double = 0, Optional pan_z As Double = 1, _
            Optional vGrid As Variant, Optional output_grid As Boolean = False) As Variant
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim vArr As Variant, uArr As Variant, d() As Double
Dim A() As Double, B() As Double
Dim cam_pos() As Double, tmp_x As Double, x_min As Double, x_max As Double
Dim xyz_med() As Double, xyz_rng() As Double, cam_xyz() As Double
Dim pi As Double
    If UBound(xyz, 2) <> 3 Then
        Debug.Print "Projection3D: input data is not 3 dimensional."
        Exit Function
    End If
    n = UBound(xyz, 1)
    
    If use_transform = True Then
        ReDim cam_pos(1 To 3)
        ReDim A(1 To 3, 1 To 3)
        For j = 1 To 3
            cam_pos(j) = vTransform(j, 1)
            For i = 1 To 3
                A(i, j) = vTransform(i, j + 1)
            Next i
        Next j
        pan_x = vTransform(1, 5)
        pan_y = vTransform(2, 5)
        pan_z = vTransform(3, 5)
        ReDim d(1 To n, 1 To 3)
        For j = 1 To 3
            For i = 1 To n
                d(i, j) = xyz(i, j) - cam_pos(j)
            Next i
        Next j
        d = modMath.M_Dot(d, A, , 1)
        ReDim vArr(1 To n, 1 To 2)
        For i = 1 To n
            If Not IsEmpty(xyz(i, 1)) And Not IsError(xyz(i, 1)) Then
                vArr(i, 1) = pan_z * d(i, 1) / d(i, 3) - pan_x
                vArr(i, 2) = pan_z * d(i, 2) / d(i, 3) - pan_y
            End If
        Next i
        Projection3D = vArr
        If output_grid = True Then
            vGrid = Projection3D_Grid(xyz, cam_pos, A, pan_x, pan_y, pan_z, 4)
        End If
        Erase vArr, cam_pos, A, d
        Exit Function
    End If
    
    'Set camera position
    pi = 3.14159265358979
    ReDim cam_pos(1 To 3)
    ReDim cam_xyz(1 To 3)
    cam_xyz(1) = cam_x: cam_xyz(2) = cam_y: cam_xyz(3) = cam_z
    For j = 1 To 3
        k = 0
        x_min = Exp(70)
        x_max = -x_min
        ReDim d(1 To n)
        For i = 1 To n
            If Not IsEmpty(xyz(i, j)) And Not IsError(xyz(i, j)) Then
                k = k + 1
                d(k) = xyz(i, j)
            End If
        Next i
        ReDim Preserve d(1 To k)
        d = modMath.fQuartile(d)
        If j = 1 Or j = 2 Then
            cam_pos(j) = d(2) + (d(4) - d(0)) * cam_xyz(j)
            If cam_xyz(j) > 0 And cam_pos(j) < d(4) Then cam_pos(j) = d(4) + (d(4) - d(0)) * cam_xyz(j)
            If cam_xyz(j) < 0 And cam_pos(j) > d(0) Then cam_pos(j) = d(0) + (d(4) - d(0)) * cam_xyz(j)
        ElseIf j = 3 Then
            cam_pos(j) = d(2) - (d(4) - d(0)) * cam_xyz(j)
            If cam_xyz(j) > 0 And cam_pos(j) > d(0) Then cam_pos(j) = d(0) - (d(4) - d(0)) * cam_xyz(j)
            If cam_xyz(j) < 0 And cam_pos(j) < d(4) Then cam_pos(j) = d(4) - (d(4) - d(0)) * cam_xyz(j)
        End If
    Next j
    Erase d
    
    'Find camera transformation matrix
    ReDim A(1 To 3, 1 To 3)
    ReDim B(1 To 3, 1 To 3)
    A(2, 2) = Cos(theta_x * pi / 2)
    A(2, 3) = Sin(theta_x * pi / 2)
    A(1, 1) = 1: A(3, 3) = A(2, 2): A(3, 2) = -A(2, 3)
    B(1, 1) = Cos(theta_y * pi / 2)
    B(1, 3) = -Sin(theta_y * pi / 2)
    B(2, 2) = 1: B(3, 1) = -B(1, 3): B(3, 3) = B(1, 1)
    A = modMath.M_Dot(A, B)
    ReDim B(1 To 3, 1 To 3)
    B(1, 1) = Cos(theta_z * pi * 2)
    B(1, 2) = Sin(theta_z * pi * 2)
    B(3, 3) = 1: B(2, 2) = B(1, 1): B(2, 1) = -B(1, 2)
    A = modMath.M_Dot(A, B)
    Erase B
    
    ReDim d(1 To n, 1 To 3)
    For j = 1 To 3
        tmp_x = cam_pos(j)
        For i = 1 To n
            d(i, j) = xyz(i, j) - tmp_x
        Next i
    Next j
    d = modMath.M_Dot(d, A, , 1)
    
    k = 0
    ReDim vArr(1 To n, 1 To 2)
    For i = 1 To n
        If Not IsEmpty(xyz(i, 1)) And Not IsError(xyz(i, 1)) Then
            vArr(i, 1) = pan_z * d(i, 1) / d(i, 3) - pan_x
            vArr(i, 2) = pan_z * d(i, 2) / d(i, 3) - pan_y
        End If
    Next i
    Projection3D = vArr
    Erase vArr, d, B
    
    If output_grid = True Then
        vGrid = Projection3D_Grid(xyz, cam_pos, A, pan_x, pan_y, pan_z, 4)
    End If
    
    If save_transform = True Then
        ReDim vTransform(1 To 3, 1 To 5)
        For j = 1 To 3
            vTransform(j, 1) = cam_pos(j)
        Next j
        For j = 1 To 3
            For i = 1 To 3
                vTransform(i, j + 1) = A(i, j)
            Next i
        Next j
        vTransform(1, 5) = pan_x
        vTransform(2, 5) = pan_y
        vTransform(3, 5) = pan_z
    End If
    Erase A, cam_pos
    
End Function


Private Function Projection3D_Grid(xyz As Variant, cam_pos() As Double, cam_matrix() As Double, _
            pan_x As Double, pan_y As Double, pan_z As Double, Optional n_grid As Long = 4) As Variant
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim x_max() As Double, x_min() As Double, x() As Double, tmp_x As Double
Dim vArr As Variant, d As Variant, uArr As Variant
Dim INFINITY
        INFINITY = Exp(70)
        n = UBound(xyz, 1)
        ReDim x_max(1 To 3)
        ReDim x_min(1 To 3)
        For j = 1 To 3
            x_max(j) = -INFINITY
            x_min(j) = INFINITY
            For i = 1 To n
                If Not IsEmpty(xyz(i, j)) And Not IsError(xyz(i, j)) Then
                    If xyz(i, j) > x_max(j) Then x_max(j) = xyz(i, j)
                    If xyz(i, j) < x_min(j) Then x_min(j) = xyz(i, j)
                End If
            Next i
            tmp_x = (x_max(j) - x_min(j)) * 0.1
            x_max(j) = x_max(j) + tmp_x
            x_min(j) = x_min(j) - tmp_x
        Next j
        
        m = 0
        ReDim d(1 To 3, 1 To 1)
        For i = 0 To n_grid
            m = m + 2
            ReDim Preserve d(1 To 3, 1 To m)
            tmp_x = x_min(1) + i * (x_max(1) - x_min(1)) / n_grid
            d(1, m - 1) = tmp_x: d(1, m) = tmp_x
            d(2, m - 1) = x_min(2): d(2, m) = x_max(2)
            d(3, m - 1) = x_max(3): d(3, m) = x_max(3)
            m = m + 1
            
            m = m + 2
            ReDim Preserve d(1 To 3, 1 To m)
            tmp_x = x_min(2) + i * (x_max(2) - x_min(2)) / n_grid
            d(1, m - 1) = x_min(1): d(1, m) = x_max(1)
            d(2, m - 1) = tmp_x: d(2, m) = tmp_x
            d(3, m - 1) = x_max(3): d(3, m) = x_max(3)
            m = m + 1

            m = m + 2
            ReDim Preserve d(1 To 3, 1 To m)
            tmp_x = x_min(1) + i * (x_max(1) - x_min(1)) / n_grid
            d(1, m - 1) = tmp_x: d(1, m) = tmp_x
            d(2, m - 1) = x_min(2): d(2, m) = x_min(2)
            d(3, m - 1) = x_min(3): d(3, m) = x_max(3)
            m = m + 1

            m = m + 2
            ReDim Preserve d(1 To 3, 1 To m)
            tmp_x = x_min(3) + i * (x_max(3) - x_min(3)) / n_grid
            d(1, m - 1) = x_min(1): d(1, m) = x_max(1)
            d(2, m - 1) = x_min(2): d(2, m) = x_min(2)
            d(3, m - 1) = tmp_x: d(3, m) = tmp_x
            m = m + 1
            
            m = m + 2
            ReDim Preserve d(1 To 3, 1 To m)
            ReDim Preserve d(1 To 3, 1 To m)
            tmp_x = x_min(2) + i * (x_max(2) - x_min(2)) / n_grid
            d(1, m - 1) = x_min(1): d(1, m) = x_min(1)
            d(2, m - 1) = tmp_x: d(2, m) = tmp_x
            d(3, m - 1) = x_min(3): d(3, m) = x_max(3)
            m = m + 1

            m = m + 2
            ReDim Preserve d(1 To 3, 1 To m)
            tmp_x = x_min(3) + i * (x_max(3) - x_min(3)) / n_grid
            d(1, m - 1) = x_min(1): d(1, m) = x_min(1)
            d(2, m - 1) = x_min(2): d(2, m) = x_max(2)
            d(3, m - 1) = tmp_x: d(3, m) = tmp_x
            m = m + 1

        Next i
        

        For j = 1 To 3
            For i = 1 To UBound(d, 2)
               If i Mod 3 <> 0 Then d(j, i) = d(j, i) - cam_pos(j)
            Next i
        Next j
        d = modMath.M_Dot(cam_matrix, d)
        ReDim vArr(1 To UBound(d, 2), 1 To 2)
        For i = 1 To UBound(d, 2)
            If i Mod 3 <> 0 Then
                vArr(i, 1) = pan_z * d(1, i) / d(3, i) - pan_x
                vArr(i, 2) = pan_z * d(2, i) / d(3, i) - pan_y
            End If
        Next i
        Projection3D_Grid = vArr
        Erase vArr
End Function
