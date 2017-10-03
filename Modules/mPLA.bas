Attribute VB_Name = "mPLA"
Option Explicit
'*** Piecewise linear approximation of time series of regular interval
'*** "An Online Algorithm for Segmenting Time Series", Keogh
'*** https://pdfs.semanticscholar.org/14e8/6f39831e30b4037ab99b5de5e5d86608ea16.pdf
'*** Bottom-up approach is implemented here

Function Trend(x() As Double, Optional max_segment As Long = 50, Optional min_len As Long = 2) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_raw As Long, n_segment As Long
Dim segment_index() As Long
Dim y_seg() As Double, x_trend() As Double
    n_raw = UBound(x, 1)
    ReDim x_trend(1 To n_raw)
    Call BottomUp(x, n_segment, max_segment, min_len, segment_index)
    j = 1
    For n = 1 To n_segment
        k = 0
        ReDim y_seg(1 To n_raw)
        For i = j To n_raw
            If segment_index(i) = n Then
                k = k + 1
                y_seg(k) = x(i)
                If i = n_raw Then j = n_raw + 1
            ElseIf segment_index(i) > n Then
                j = i
                Exit For
            End If
        Next i
        ReDim Preserve y_seg(1 To k)
        y_seg = Linear_Trend(y_seg)
        For i = 1 To k
            x_trend(j - k - 1 + i) = y_seg(i)
        Next i
    Next n
    Trend = x_trend
    Erase y_seg, x_trend, segment_index
End Function


Sub BottomUp(x() As Double, n_segment As Long, Optional max_segment As Long = 50, Optional min_len As Long = 2, Optional segment_index As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_raw As Long
Dim min_cost As Double, INFINITY As Double, mse As Double
Dim segment() As Long
Dim sse() As Double, cost() As Double

    INFINITY = Exp(70)
    n_raw = UBound(x, 1)
    n_segment = n_raw - 1

    ReDim segment(1 To 2, 1 To n_segment)
    ReDim sse(1 To n_segment)
    ReDim cost(1 To n_segment - 1)
    For i = 1 To n_raw - 1
        segment(1, i) = i
        segment(2, i) = i + 1
    Next i
    For i = 1 To n_segment - 1
        cost(i) = Merge_Cost(x, sse(i), sse(i + 1), segment(1, i), segment(2, i), segment(2, i + 1))
    Next i
    
    'Find best segment to merge
    Do While n_segment > max_segment
        min_cost = INFINITY
        For i = 1 To n_segment - 1
            If cost(i) < min_cost Then
                min_cost = cost(i)
                n = i
            End If
        Next i
        Call Merge_segment(n, n_segment, x, segment, sse, cost)
'        mse = 0
'        For i = 1 To n_segment
'            mse = mse + sse(i)
'        Next i
'        mse = mse / n_raw
    Loop

    'Merge any remaining segments shorter than minimum allowable length
    Do
        m = 0
        For i = 1 To n_segment
            If (segment(2, i) - segment(1, i) + 1) < min_len Then
                m = m + 1
                n = i
                Exit For
            End If
        Next i
        If m = 0 Then Exit Do
        If n = n_segment Then n = n - 1
        Call Merge_segment(n, n_segment, x, segment, sse, cost)
    Loop
    
    If IsMissing(segment_index) = False Then
        ReDim segment_index(1 To n_raw)
        For n = 1 To n_segment
            For i = segment(1, n) To segment(2, n)
                segment_index(i) = n
            Next i
        Next n
    End If
    
    Erase sse, cost, segment
End Sub


'Merge segment n & n+1
Private Sub Merge_segment(n As Long, n_segment As Long, x() As Double, segment() As Long, sse() As Double, cost() As Double)
Dim i As Long
    segment(1, n) = segment(1, n)
    segment(2, n) = segment(2, n + 1)
    
    'Update sse of the merged segment
    sse(n) = cost(n) + sse(n) + sse(n + 1)
    
    'Update merge cost
    If n > 1 Then cost(n - 1) = Merge_Cost(x, sse(n - 1), sse(n), segment(1, n - 1), segment(2, n - 1), segment(2, n))
    If (n + 1) < n_segment Then cost(n) = Merge_Cost(x, sse(n), sse(n + 2), segment(1, n), segment(2, n), segment(2, n + 2))
    
    'Delete segment n+1 and reindex the array
    n_segment = n_segment - 1
    For i = n + 1 To n_segment
        sse(i) = sse(i + 1)
        If i < n_segment Then cost(i) = cost(i + 1)
        segment(1, i) = segment(1, i + 1)
        segment(2, i) = segment(2, i + 1)
    Next i
    ReDim Preserve sse(1 To n_segment)
    ReDim Preserve cost(1 To n_segment - 1)
    ReDim Preserve segment(1 To 2, 1 To n_segment)
End Sub


Private Function Merge_Cost(x() As Double, sse1 As Double, sse2 As Double, s1 As Long, s2 As Long, s3 As Long) As Double
    Merge_Cost = Segment_SSE(x, s1, s3) - sse1 - sse2
End Function

Private Function Segment_SSE(x() As Double, s As Long, t As Long) As Double
Dim i As Long, n As Long
Dim x_avg As Double, beta As Double, tmp_x As Double, tmp_y As Double
    n = t - s + 1
    For i = s To t
        x_avg = x_avg + x(i)
    Next i
    x_avg = x_avg / n
    
    tmp_y = s - 1 + (n + 1) * 0.5
    For i = s To t
        beta = beta + (x(i) - x_avg) * (i - tmp_y)
    Next i
    beta = ((beta / n) * 12) / (n * n - 1)
    
    tmp_x = 0
    For i = s To t
        tmp_x = tmp_x + ((x(i) - x_avg) - beta * (i - tmp_y)) ^ 2
    Next i
    Segment_SSE = tmp_x
End Function

Private Function Linear_Trend(x() As Double) As Double()
Dim i As Long, n As Long
Dim x_slope As Double, x_intercept As Double, x_avg As Double
Dim x_trend() As Double
    n = UBound(x)
    For i = 1 To n
        x_avg = x_avg + x(i)
    Next i
    x_avg = x_avg / n
    
    For i = 1 To n
        x_slope = x_slope + (x(i) - x_avg) * (i - (n + 1) * 0.5)
    Next i
    x_slope = ((x_slope / n) * 12) / (n * n - 1)
    
    ReDim x_trend(1 To n)
    For i = 1 To n
        x_trend(i) = x_slope * (i - (n + 1) * 0.5) + x_avg
    Next i
    Linear_Trend = x_trend
    Erase x_trend
End Function
