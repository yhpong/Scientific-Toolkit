Attribute VB_Name = "mTSSeg"
Option Explicit

'==================================================
'"Segmenting Time Series: A Survey and Novel Approach"
'Eamonn Keogh, Selina Chu, David Hart, Michael Pazzani
'==================================================

'===========================================
'Sliding window approach to linear segmentation
'can be used in online data
'============================================
'Input: x(1:N), time series of length N
'       x_threshold, segmentation error must be lower than this threshold
'       step_size, steps to wait until a new segment is evaluated
'       errType, cost function of segmentation, SSE is sum of square, SSE_NORM is sqr(SSE)/(max-min)
'       fitType, "INTERPOL": directly joining two points, "REGRESSION": least square fit between two data points
'Output: i_anchor(1:m+1), m starting points of each segment, the m+1-th entry is simply N for convenience
Sub Seg_Sliding(i_anchor() As Long, x() As Double, Optional x_threshold As Double = 1, Optional step_size As Long = 1, _
                Optional errType As String = "SSE_NORM", Optional fitType As String = "INTERPOL")
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_anchor As Long
Dim x_err As Double
Dim x_tmp() As Double
Dim prev_anchor As Long

    n = UBound(x, 1)
    
    ReDim i_anchor(1 To 1)
    n_anchor = 1
    i_anchor(1) = 1
    prev_anchor = 1
    
    ReDim x_tmp(1 To 1)
    x_tmp(1) = x(1)
    k = 1
    
    For i = 2 To n Step step_size
        
        ReDim Preserve x_tmp(1 To k + step_size)
        For j = 1 To step_size
            x_tmp(k + j) = x(prev_anchor + k + j - 1)
        Next j
        k = k + step_size
        
        If calc_seg_err(x_tmp, errType, fitType) >= x_threshold Then
            n_anchor = n_anchor + 1
            ReDim Preserve i_anchor(1 To n_anchor)
            i_anchor(n_anchor) = i - 1
            prev_anchor = i - 1
            
            ReDim x_tmp(1 To 1)
            x_tmp(1) = x(i - 1)
            k = 1
        End If
        
    Next i
    
    ReDim Preserve i_anchor(1 To n_anchor + 1)
    i_anchor(n_anchor + 1) = n
    
End Sub



'===========================================
'Bottom-up approach to linear segmentation
'============================================
'Input: x(1:N), time series of length N
'       x_threshold, segmentation error must be lower than this threshold, defined as % of maximum segmentation error if joining start and end points directly
'       n_segment, target number of segments, override x_threshold if provided.
'       errType, cost function of segmentation, SSE is sum of square, SSE_NORM is sqr(SSE)/(max-min)
'       fitType, "INTERPOL": directly joining two points, "REGRESSION": least square fit between two data points
'       sign_penalty, penalize a merge if slopes of the two segments do not have the same signs
'Output: i_anchor(1:m+1), m starting points of each segment, the m+1-th entry is simply N for convenience
Sub Seg_BottomUp(i_anchor() As Long, x() As Double, Optional x_threshold As Double = 0.05, Optional n_segment As Long = -1, _
                Optional errType As String = "SSE", Optional fitType As String = "INTERPOL", Optional sign_penalty As Double = 0)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_anchor As Long
Dim x_tmp() As Double, y_tmp() As Double
Dim prev_anchor As Long
Dim x_err() As Double, x_cost_min As Double, i_min As Long, x_cost() As Double
Dim x_err_total As Double, x_err_max As Double
Dim isStop As Boolean
Dim x_penalty As Double, tmp_x As Double, tmp_y As Double, tmp_z As Double

    n = UBound(x, 1)
    
    'Initilize all consecutive points as segments
    n_anchor = n - 1
    x_err_total = 0
    ReDim i_anchor(1 To n_anchor)
    ReDim x_err(1 To n_anchor)
    For i = 1 To n_anchor
        i_anchor(i) = i
    Next i
    
    'Error if joining start and end by a straight line
    x_err_max = calc_seg_err(x, errType, fitType)

    'Calculate merge cost of each segment with the one after it
    x_cost_min = Exp(70)
    ReDim x_cost(1 To n_anchor - 1)
    For i = 1 To n_anchor - 1

        If i = n_anchor - 1 Then
            m = n - i_anchor(i) + 1
        Else
            m = i_anchor(i + 2) - i_anchor(i) + 1
        End If
        ReDim x_tmp(1 To m)
        For j = 1 To m
            x_tmp(j) = x(i_anchor(i) + j - 1)
        Next j

        x_cost(i) = calc_seg_err(x_tmp, errType, fitType) - x_err(i) - x_err(i + 1)
        
        If x_cost(i) < x_cost_min Then
            x_cost_min = x_cost(i)
            i_min = i
        End If
        
    Next i
    
    If n_segment > 0 Then
        isStop = n_anchor <= n_segment
    Else
        isStop = x_err_total > (x_threshold * x_err_max)
    End If
    
    'Merge segment with lowest merge cost until stopping criteriea is met
    Do While isStop = False

        'Merge i_min and i_min+1
        x_err_total = x_err_total - x_err(i_min) - x_err(i_min + 1)
        For i = i_min + 1 To n_anchor - 1
            i_anchor(i) = i_anchor(i + 1)
            x_err(i) = x_err(i + 1)
        Next i
        For i = i_min + 1 To n_anchor - 2
            x_cost(i) = x_cost(i + 1)
        Next i
        n_anchor = n_anchor - 1
        If n_anchor = 1 Then Exit Do
        ReDim Preserve i_anchor(1 To n_anchor)
        ReDim Preserve x_err(1 To n_anchor)
        ReDim Preserve x_cost(1 To n_anchor - 1)
        
        'Calculate new error in merged segment
        If i_min = n_anchor Then
            m = n - i_anchor(i_min) + 1
        Else
            m = i_anchor(i_min + 1) - i_anchor(i_min) + 1
        End If
        ReDim x_tmp(1 To m)
        For j = 1 To m
            x_tmp(j) = x(i_anchor(i_min) + j - 1)
        Next j
        x_err(i_min) = calc_seg_err(x_tmp, errType, fitType)
        x_err_total = x_err_total + x_err(i_min)
        
        'Update merge cost of i_min-1 and i_min
        For i = IIf(i_min = 1, 1, i_min - 1) To IIf(i_min = n_anchor, n_anchor - 1, i_min)
            If i = n_anchor - 1 Then
                k = n
            Else
                k = i_anchor(i + 2)
            End If
            
            'extra penalty if signs of two slopes are different
            tmp_y = (x(k) - x(i_anchor(i + 1))) / (k - i_anchor(i + 1))
            tmp_x = (x(i_anchor(i + 1)) - x(i_anchor(i))) / (i_anchor(i + 1) - i_anchor(i))
            If Sgn(tmp_y) = Sgn(tmp_x) Then
                tmp_z = 0
            Else
                tmp_z = sign_penalty * x_err_total '* Abs(tmp_y - tmp_x) / (Abs(tmp_x) + Abs(tmp_y))
            End If
            
            m = k - i_anchor(i) + 1
            ReDim x_tmp(1 To m)
            For j = 1 To m
                x_tmp(j) = x(i_anchor(i) + j - 1)
            Next j
            x_cost(i) = calc_seg_err(x_tmp, errType, fitType) - x_err(i) - x_err(i + 1) + tmp_z
        Next i
        
        'New segments with mininum merge cost
        x_cost_min = Exp(70)
        For i = 1 To n_anchor - 1
            If x_cost(i) < x_cost_min Then
                x_cost_min = x_cost(i)
                i_min = i
            End If
        Next i
        
        If n_segment > 0 Then
            isStop = n_anchor <= n_segment
        Else
            isStop = x_err_total > (x_threshold * x_err_max)
        End If
        
    Loop
    
    ReDim Preserve i_anchor(1 To n_anchor + 1)
    i_anchor(n_anchor + 1) = n
    
End Sub


Private Function calc_seg_err(x() As Double, Optional errType As String = "SSE", Optional fitType As String = "INTERPOL") As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim x_err As Double
Dim x_mean As Double, x_max As Double, x_min As Double, i_mean As Double
Dim x_slope As Double, x_intercept As Double, t_mean As Double

    n = UBound(x, 1)
    
    'Two data points only, trivial solution
    If n <= 2 Then
        calc_seg_err = 0
        Exit Function
    End If
    
    x_mean = 0
    x_min = Exp(70)
    x_max = -Exp(70)
    For i = 1 To n
        x_mean = x_mean + x(i)
        If x(i) > x_max Then x_max = x(i)
        If x(i) < x_min Then x_min = x(i)
    Next i
    x_mean = x_mean / n
    
    'Flat line, trivial solution
    If (x_max = x_min) Then
        calc_seg_err = 0
        Exit Function
    End If
    
    If UCase(fitType) = "INTERPOL" Then
        x_err = 0
        x_slope = (x(n) - x(1)) / (n - 1)
        For i = 1 To n
            tmp_x = (i - 1) * x_slope + x(1)
            x_err = x_err + (x(i) - tmp_x) ^ 2
        Next i
    ElseIf UCase(fitType) = "REGRESSION" Then
        x_slope = 0
        i_mean = (n + 1) / 2
        tmp_z = n * (n * n - 1) / 12
        For i = 1 To n
            x_slope = x_slope + (i - i_mean) * (x(i) - x_mean)
        Next i
        x_slope = x_slope / tmp_z
        x_intercept = x_mean - x_slope * i_mean
        
        x_err = 0
        For i = 1 To n
            tmp_x = i * x_slope + x_intercept
            x_err = x_err + (x(i) - tmp_x) ^ 2
        Next i
    Else
        MsgBox "calc_seg_err: " & fitType & " is not supported."
        End
    End If
    
    Select Case UCase(errType)
    Case "SSE"
        calc_seg_err = x_err
    Case "SSE_NORM"
        calc_seg_err = Sqr(x_err) / (x_max - x_min)
    Case Else
        MsgBox "calc_seg_err: " & errType & " is not supported."
        End
    End Select
    
End Function

