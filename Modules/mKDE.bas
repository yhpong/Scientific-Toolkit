Attribute VB_Name = "mKDE"
Option Explicit

'=================================================================
'2-Dimensional kernel density estimate using gaussian kernel
'=================================================================
'Input:
'x(1:N,1:2),     N observations of 2-dimensional data
'n_x, n_y,       number of grids to in x-y corrdinates
's_x, s_y, rho,  size and correlation of Gaussian kernel
'x_fix, y_fix,   fix positions of x or y where the pdf is calculated
'Output:
'x_pdf(1:n_y,1:n_x), if x_fix and y_fix are both null, 2D-matrix of probability densities
'x_pdf(1:n_y,1:2),   if only x_fix is given, probability densities along fixed x
'x_pdf(1:n_x,1:2),   if only y_fix is given, probability densities along fixed y
'x_pdf,              if x_fix and y_fix are both given, probability densities at the fixed point
'x_minmax(1:2,1:D),  min-max coordinates of corners cells
Sub KDE_2D(x As Variant, n_x As Long, n_y As Long, x_pdf As Variant, _
                Optional s_x As Variant = Null, Optional s_y As Variant = Null, Optional rho As Variant = Null, _
                Optional x_fix As Variant = Null, Optional y_fix As Variant = Null, _
                Optional x_minmax As Variant = Null)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long
Dim ii As Long
Dim x_mean() As Double, x_sd() As Double, x_min() As Double, x_max() As Double
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim tmp_xx As Double, tmp_yy As Double, tmp_zz As Double
Dim pi2 As Double, kernel_norm As Double
Dim dx() As Double
Dim is_x_fix As Long
    pi2 = 6.28318530717959
    
    is_x_fix = 0
    If (Not IsNull(x_fix)) And IsNull(y_fix) Then
        is_x_fix = 1
    ElseIf IsNull(x_fix) And Not IsNull(y_fix) Then
        is_x_fix = 2
    ElseIf Not IsNull(x_fix) And Not IsNull(y_fix) Then
        is_x_fix = 3
    End If
    
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim x_mean(1 To n_dimension)
    ReDim x_sd(1 To n_dimension)
    ReDim x_min(1 To n_dimension)
    ReDim x_max(1 To n_dimension)
    ReDim dx(1 To n_dimension)
    If Not IsNull(x_minmax) Then ReDim x_minmax(1 To 2, 1 To n_dimension)
    For k = 1 To n_dimension
        tmp_x = 0: tmp_y = 0
        x_min(k) = Exp(70)
        x_max(k) = -Exp(70)
        For i = 1 To n
            tmp_x = tmp_x + x(i, k)
            tmp_y = tmp_y + x(i, k) ^ 2
            If x(i, k) < x_min(k) Then x_min(k) = x(i, k)
            If x(i, k) > x_max(k) Then x_max(k) = x(i, k)
        Next i
        x_sd(k) = Sqr((tmp_y - tmp_x * (tmp_x / n)) / (n - 1))
        x_mean(k) = tmp_x / n
        If Not IsNull(x_minmax) Then
            x_minmax(1, k) = x_min(k)
            x_minmax(2, k) = x_max(k)
        End If
        If k = 1 Then
            m = n_x
        ElseIf k = 2 Then
            m = n_y
        End If
        If m > 1 Then dx(k) = (x_max(k) - x_min(k)) / (m - 1)
    Next k
    
    
    If VBA.IsNull(s_x) Then s_x = 1# / (n ^ (1# / 6))   'Silverman's rule of thumb
    If VBA.IsNull(s_y) Then s_y = 1# / (n ^ (1# / 6))   'Silverman's rule of thumb
    If VBA.IsNull(rho) Then
        rho = 0
        For i = 1 To n
            rho = rho + (x(i, 1) - x_mean(1)) * (x(i, 2) - x_mean(2))
        Next i
        rho = rho / ((n - 1) * x_sd(1) * x_sd(2))
    End If
    
    kernel_norm = n * pi2 * Sqr(1 - rho * rho) * s_x * s_y * x_sd(1) * x_sd(2) 'Gaussian normalization constant
    If is_x_fix = 0 Then
    
        ReDim x_pdf(1 To n_y, 1 To n_x)
        For j = 1 To n_y
            tmp_y = x_min(2) + (n_y - j) * dx(2)
            For i = 1 To n_x
                tmp_x = x_min(1) + (i - 1) * dx(1)
                tmp_z = 0
                For ii = 1 To n
                    tmp_xx = (tmp_x - x(ii, 1)) / (x_sd(1) * s_x)
                    tmp_yy = (tmp_y - x(ii, 2)) / (x_sd(2) * s_y)
                    tmp_xx = tmp_xx ^ 2 + tmp_yy ^ 2 - 2 * rho * tmp_xx * tmp_yy
                    tmp_z = tmp_z + Exp(-0.5 * tmp_xx / (1 - rho ^ 2))
                Next ii
                x_pdf(j, i) = tmp_z / kernel_norm
            Next i
        Next j
        
    ElseIf is_x_fix = 1 Then
    
        ReDim x_pdf(1 To n_y, 1 To 2)
        tmp_x = x_fix
        For j = 1 To n_y
            tmp_y = x_min(2) + (n_y - j) * dx(2)
            x_pdf(j, 1) = tmp_y
            tmp_z = 0
            For ii = 1 To n
                tmp_xx = (tmp_x - x(ii, 1)) / (x_sd(1) * s_x)
                tmp_yy = (tmp_y - x(ii, 2)) / (x_sd(2) * s_y)
                tmp_xx = tmp_xx ^ 2 + tmp_yy ^ 2 - 2 * rho * tmp_xx * tmp_yy
                tmp_z = tmp_z + Exp(-0.5 * tmp_xx / (1 - rho ^ 2))
            Next ii
            x_pdf(j, 2) = tmp_z / kernel_norm
        Next j

    ElseIf is_x_fix = 2 Then
    
        ReDim x_pdf(1 To n_x, 1 To 2)
        tmp_y = y_fix
        For i = 1 To n_x
            tmp_x = x_min(1) + (i - 1) * dx(1)
            x_pdf(i, 1) = tmp_x
            tmp_z = 0
            For ii = 1 To n
                tmp_xx = (tmp_x - x(ii, 1)) / (x_sd(1) * s_x)
                tmp_yy = (tmp_y - x(ii, 2)) / (x_sd(2) * s_y)
                tmp_xx = tmp_xx ^ 2 + tmp_yy ^ 2 - 2 * rho * tmp_xx * tmp_yy
                tmp_z = tmp_z + Exp(-0.5 * tmp_xx / (1 - rho ^ 2))
            Next ii
            x_pdf(i, 2) = tmp_z / kernel_norm
        Next i

    ElseIf is_x_fix = 3 Then
    
        x_pdf = 0
        tmp_x = x_fix
        tmp_y = y_fix
        For ii = 1 To n
            tmp_xx = (tmp_x - x(ii, 1)) / (x_sd(1) * s_x)
            tmp_yy = (tmp_y - x(ii, 2)) / (x_sd(2) * s_y)
            tmp_zz = tmp_xx ^ 2 + tmp_yy ^ 2 - 2 * rho * tmp_xx * tmp_yy
            x_pdf = x_pdf + Exp(-0.5 * tmp_zz / (1 - rho ^ 2))
        Next ii
        x_pdf = x_pdf / kernel_norm
        
    End If
    
End Sub


'====================================================================
'Use cross validation to find optimal bandwith matrix to be used in KDE_2D()
'====================================================================
'Input:  x(1:N,1:2), N observations of 2-dimensional data
'Output: s_x_opt, s_y_opt and rho_opt are optimal values of bandwidth and orientation to be used in KDE_2D()
Sub KDE_2D_CrossValidate(x As Variant, s_x_opt As Variant, s_y_opt As Variant, rho_opt As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long
Dim iterate As Long, iterate_sx As Long, iterate_sy As Long, iterate_rho As Long
Dim x_mean() As Double, x_sd() As Double
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim pi2 As Double, kernel_norm As Double, INFINITY As Double
Dim x_ISE As Double, x_ISE_min As Double, x_ISE_min_cur As Double
Dim s_x As Double, s_y As Double, rho As Double
Dim s_x_min As Double, s_y_min As Double, rho_min As Double
Dim s_x_max As Double, s_y_max As Double, rho_max As Double

    INFINITY = Exp(70)
    pi2 = 6.28318530717959
    
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim x_mean(1 To n_dimension)
    ReDim x_sd(1 To n_dimension)
    For k = 1 To n_dimension
        tmp_x = 0: tmp_y = 0
        For i = 1 To n
            tmp_x = tmp_x + x(i, k)
            tmp_y = tmp_y + x(i, k) ^ 2
        Next i
        x_sd(k) = Sqr((tmp_y - tmp_x * (tmp_x / n)) / (n - 1))
        x_mean(k) = tmp_x / n
    Next k

    s_x = 1# / (n ^ (1# / 6))   'Silverman's rule of thumb
    s_y = 1# / (n ^ (1# / 6))   'Silverman's rule of thumb
    rho = 0                     'Sample correlation
    For i = 1 To n
        rho = rho + (x(i, 1) - x_mean(1)) * (x(i, 2) - x_mean(2))
    Next i
    rho = rho / ((n - 1) * x_sd(1) * x_sd(2))

    s_x_min = s_x / 2: s_x_max = s_x * 2
    s_y_min = s_x / 2: s_y_max = s_y * 2
    rho_min = 0: rho_max = rho * 2
    If rho_max >= 1 Then
        rho_max = 0.9
    ElseIf rho_max <= -1 Then
        rho_max = -0.9
    End If
    
    x_ISE_min = INFINITY
    For iterate_rho = 1 To 10
        For iterate_sx = 1 To 10
            x_ISE_min_cur = INFINITY
            DoEvents
            Application.StatusBar = "KDE_2D_CrossValidate: " & iterate_rho & "/10," & iterate_sx & "/10"
            For iterate_sy = 1 To 10
        
                s_x = s_x_min + (iterate_sx - 1) * (s_x_max - s_x_min) / (10 - 1)
                s_y = s_y_min + (iterate_sy - 1) * (s_y_max - s_y_min) / (10 - 1)
                rho = rho_min + (iterate_rho - 1) * (rho_max - rho_min) / (10 - 1)
                kernel_norm = n * pi2 * Sqr(1 - rho * rho) * s_x * s_y * x_sd(1) * x_sd(2)
    
                '(Integrate f^2 dxdy) - (2/n) sum f_leave_one_out
                x_ISE = 0
                tmp_z = 0
                For i = 1 To n - 1
                    For j = i + 1 To n
                        tmp_x = (x(i, 1) - x(j, 1)) / (x_sd(1) * s_x)
                        tmp_y = (x(i, 2) - x(j, 2)) / (x_sd(2) * s_y)
                        tmp_x = tmp_x ^ 2 + tmp_y ^ 2 - 2 * rho * tmp_x * tmp_y
                        x_ISE = x_ISE + Exp(-0.25 * tmp_x / (1 - rho ^ 2))
                        tmp_z = tmp_z + Exp(-0.5 * tmp_x / (1 - rho ^ 2))
                    Next j
                Next i
                x_ISE = (0.5 + x_ISE / n) / kernel_norm - 4 * tmp_z / ((n - 1) * kernel_norm)
                
                If x_ISE < x_ISE_min Then
                    x_ISE_min = x_ISE
                    s_x_opt = s_x
                    s_y_opt = s_y
                    rho_opt = rho
                End If
    
                If x_ISE <= x_ISE_min_cur Then
                    x_ISE_min_cur = x_ISE
                Else
                    Exit For
                End If
    
            Next iterate_sy
        Next iterate_sx
    Next iterate_rho
    Debug.Print "(s_x,s_y,rho)=" & Round(s_x_opt, 4) & ", " & Round(s_y_opt, 4) & ", " & Round(rho_opt, 4)
    Application.StatusBar = False
End Sub
