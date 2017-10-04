Attribute VB_Name = "mDTW"
Option Explicit


'=== DTW ==========================================
'Main algorithm of dynamic time warping =======
'Input: x() and y(), multi-dimensional time series (1:T, 1:dimension)
'Input: window, cab be an integer for fixed width band, or an array giving the lower and upper bound of time index of y for each time index of x
'Input: dist_type, 1=absolute, 2=Euclidean, 3=Euclidean-Squared
'Output: DTW itself is the distance between x & y, path() and d() are the path matrix and distance matrix
'==================================================
Function DTW(x() As Double, y() As Double, Optional w_window As Variant, _
        Optional dist_type As Long = 2, Optional path As Variant, Optional d As Variant) As Double
Dim i As Long, j As Long, k As Long, n As Long, m As Long, w As Long
Dim n_x As Long, n_y As Long, n_dimension As Long
Dim i_min As Long, j_lo As Long, j_hi As Long
Dim dist() As Double
Dim cost As Double

Dim INFINITY As Double

If dist_type < 1 Or dist_type > 3 Then
    Debug.Print "DTW: dist_type " & dist_type & " not implemented."
    Exit Function
End If

INFINITY = Exp(70)

n_x = UBound(x, 1)
n_y = UBound(y, 1)
n_dimension = UBound(x, 2)
ReDim dist(0 To n_x, 0 To n_y)

If IsMissing(w_window) = True Then

    For i = 1 To n_x
        dist(i, 0) = INFINITY
    Next i
    For i = 1 To n_y
        dist(0, i) = INFINITY
    Next i
    
    For i = 1 To n_x
        For j = 1 To n_y
            cost = cost_ij(x, y, i, j, dist_type)
            dist(i, j) = cost + min3(dist(i - 1, j), dist(i, j - 1), dist(i - 1, j - 1))
        Next j
    Next i
 
ElseIf IsArray(w_window) = False Then

    w = w_window
    If Abs(n_x - n_y) > w Then w = Abs(n_x - n_y)
    For i = 0 To n_x
        For j = 0 To n_y
            dist(i, j) = INFINITY
        Next j
    Next i
    dist(0, 0) = 0
    
    For i = 1 To n_x
        j_lo = max2(1, i - w)
        j_hi = min2(n_y, i + w)
        For j = j_lo To j_hi
            cost = cost_ij(x, y, i, j, dist_type)
            dist(i, j) = cost + min3(dist(i - 1, j), dist(i, j - 1), dist(i - 1, j - 1))
        Next j
    Next i

ElseIf IsArray(w_window) = True Then

    For i = 0 To n_x
        For j = 0 To n_y
            dist(i, j) = INFINITY
        Next j
    Next i
    dist(0, 0) = 0
    
    For i = 1 To n_x
        For j = w_window(i, 1) To w_window(i, 2)
            cost = cost_ij(x, y, i, j, dist_type)
            dist(i, j) = cost + min3(dist(i - 1, j), dist(i, j - 1), dist(i - 1, j - 1))
        Next j
    Next i
    
End If

DTW = dist(n_x, n_y)

'Run only if distance matrix to be output
If IsMissing(d) = False Then
    ReDim d(1 To n_x, 1 To n_y)
    For i = 1 To n_x
        For j = 1 To n_y
            d(i, j) = dist(i, j)
        Next j
    Next i
End If

'Run only if path needs to be output
If IsMissing(path) = False Then

    Dim path_tmp() As Long
    k = 1
    ReDim path_tmp(1 To 2, 1 To n_x + n_y - 1)
    path_tmp(1, 1) = n_x
    path_tmp(2, 1) = n_y
    i = n_x
    j = n_y
    Do While i > 1 Or j > 1
        If i = 1 Then
            j = j - 1
        ElseIf j = 1 Then
            i = i - 1
        Else
            i_min = min3_idx(dist(i - 1, j - 1), dist(i - 1, j), dist(i, j - 1))
            If i_min = 2 Then
                i = i - 1
            ElseIf i_min = 3 Then
                j = j - 1
            Else
                i = i - 1
                j = j - 1
            End If
        End If
        k = k + 1
        path_tmp(1, k) = i
        path_tmp(2, k) = j
    Loop
    ReDim Preserve path_tmp(1 To 2, 1 To k)

    ReDim path(1 To k, 1 To 2)
    For i = 1 To k
        path(i, 1) = path_tmp(1, k - i + 1)
        path(i, 2) = path_tmp(2, k - i + 1)
    Next i
    Erase path_tmp
End If
Erase dist
End Function


'Distance between x(i) and y(j), default is Euclidean (dist_type=2)
Private Function cost_ij(x() As Double, y() As Double, i As Long, j As Long, Optional dist_type As Long = 2) As Double
Dim d As Long, n_dimension As Long
    n_dimension = UBound(x, 2)
    cost_ij = 0
    If dist_type = 1 Then
        For d = 1 To n_dimension
            cost_ij = cost_ij + Abs(x(i, d) - y(j, d))
        Next d
    ElseIf dist_type = 2 Then
        For d = 1 To n_dimension
            cost_ij = cost_ij + (x(i, d) - y(j, d)) ^ 2
        Next d
        cost_ij = Sqr(cost_ij)
    ElseIf dist_type = 3 Then
        For d = 1 To n_dimension
            cost_ij = cost_ij + (x(i, d) - y(j, d)) ^ 2
        Next d
    End If
End Function

'Return the warped x() & y() side by side
Function xy_warped(x() As Double, y() As Double, path() As Long) As Double()
Dim i As Long, d As Long, n_dimension As Long
Dim x_out() As Double
    n_dimension = UBound(x, 2)
    ReDim x_out(1 To UBound(path, 1), 1 To 2 * n_dimension)
    For i = 1 To UBound(path, 1)
        For d = 1 To n_dimension
            x_out(i, d) = x(path(i, 1), d)
            x_out(i, n_dimension + d) = y(path(i, 2), d)
        Next d
    Next i
    xy_warped = x_out
End Function

'Print x() above y() and use lines to visualize mapping between the two series
Sub Print_Warped_Series(vRng As Range, path() As Long, x() As Double, y() As Double, Optional shift_y As Double = -3)
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
    n = UBound(path, 1)
    n_dimension = UBound(x, 2)
    With vRng
        k = 0
        For i = 1 To n
            .Offset(k, 0).Value = path(i, 1)
            .Offset(k + 1, 0).Value = path(i, 2)
            For j = 1 To n_dimension
                .Offset(k, j).Value = x(path(i, 1), j)
                .Offset(k + 1, j).Value = y(path(i, 2), j) + shift_y
            Next j
            k = k + 3
        Next i
    End With
End Sub

'Print M x N distance matrix, with step_size>1 to skip some points for lower resolution
Sub Print_Distance_Matrix(vRng As Range, dist() As Double, Optional step_size As Long = 1)
Dim i As Long, j As Long, k As Long
Dim xArr As Variant
Dim INFINITY As Double
    INFINITY = Exp(50)
    k = 0
    ReDim xArr(1 To 3, 1 To UBound(dist, 1) * UBound(dist, 2))
    For i = 1 To UBound(dist, 1) Step step_size
        For j = 1 To UBound(dist, 2) Step step_size
            If dist(i, j) < INFINITY Then
                k = k + 1
                xArr(1, k) = i
                xArr(2, k) = j
                xArr(3, k) = dist(i, j) ^ 2
            End If
        Next j
    Next i
    ReDim Preserve xArr(1 To 3, 1 To k)
    Range(vRng, vRng.Offset(k - 1, 2)).Value = modMath.wkshtTranspose(xArr)
End Sub


Private Function min2(x As Long, y As Long) As Long
    min2 = x
    If y < min2 Then min2 = y
End Function

Private Function max2(x As Long, y As Long) As Long
    max2 = x
    If y > max2 Then max2 = y
End Function

Private Function min3(x As Double, y As Double, z As Double) As Double
    min3 = x
    If y < min3 Then min3 = y
    If z < min3 Then min3 = z
End Function

Private Function min3_idx(x As Double, y As Double, z As Double) As Long
Dim min3 As Double
    min3 = x
    min3_idx = 1
    If y < min3 Then
        min3 = y
        min3_idx = 2
    End If
    If z < min3 Then
        min3 = z
        min3_idx = 3
    End If
End Function

Private Sub min3_val_idx(x As Double, y As Double, z As Double, min3 As Double, min3_idx As Long)
    min3 = x
    min3_idx = 1
    If y < min3 Then
        min3 = y
        min3_idx = 2
    End If
    If z < min3 Then
        min3 = z
        min3_idx = 3
    End If
End Sub


'Extract sequence s to t from x(1 to T, 1 to dimension)
Private Function sub_seq(x() As Double, s As Long, t As Long) As Double()
Dim i As Long, d As Long, n_dimension As Long
Dim xS() As Double
    n_dimension = UBound(x, 2)
    ReDim xS(1 To t - s + 1, 1 To n_dimension)
    For i = 1 To t - s + 1
        For d = 1 To n_dimension
            xS(i, d) = x(s + i - 1, d)
        Next d
    Next i
    sub_seq = xS
    Erase xS
End Function



'==================================================
'"FastDTW: Toward Accurate Dynamic Time Warping in Linear Time and Space", Stan Salvador
'==================================================
'Input: x(), y(), multivariate time series of size (1:T, 1:D)
'       radius, distance to search outside of the projected warp path
'               from the prev. resolution when refining the warp path
Function FastDTW(x() As Double, y() As Double, Optional radius As Long = 2, Optional dist_type As Long = 2, Optional path As Variant, Optional d As Variant) As Double
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim n_x As Long, n_y As Long, minTSsize As Long
Dim w() As Long
Dim tmp_x As Double
Dim x_shrunk() As Double, y_shrunk() As Double
Dim lowResPath() As Long
Dim lowResd() As Double
    n_x = UBound(x, 1)
    n_y = UBound(y, 1)
    minTSsize = radius + 2
    If n_x <= minTSsize Or n_y <= minTSsize Then
        If IsMissing(path) And IsMissing(d) Then
            FastDTW = DTW(x, y, , dist_type)
        ElseIf IsMissing(d) Then
            FastDTW = DTW(x, y, , dist_type, path:=path)
        Else
            FastDTW = DTW(x, y, , dist_type, path:=path, d:=d)
        End If
    Else
        x_shrunk = reduceByHalf(x)
        y_shrunk = reduceByHalf(y)
        tmp_x = FastDTW(x_shrunk, y_shrunk, radius, dist_type, lowResPath)
        w = ExpandedResWindow(lowResPath, x, y, radius)
        If IsMissing(path) And IsMissing(d) Then
            FastDTW = DTW(x, y, w, dist_type)
        ElseIf IsMissing(d) Then
            FastDTW = DTW(x, y, w, dist_type, path:=path)
        Else
            FastDTW = DTW(x, y, w, dist_type, path:=path, d:=d)
        End If
    End If
End Function


Private Function reduceByHalf(x() As Double) As Double()
Dim i As Long, j As Long, k As Long, d As Long
Dim n_raw As Long, n_new As Long, n_dimension As Long
Dim x_shrunk() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    k = n_raw Mod 2
    n_new = (n_raw + k) / 2
    ReDim x_shrunk(1 To n_new, 1 To n_dimension)
    j = 0
    For i = 1 To n_raw - 1 - k Step 2
        j = j + 1
        For d = 1 To n_dimension
            x_shrunk(j, d) = (x(i, d) + x(i + 1, d)) * 0.5
        Next d
    Next i
    If k = 1 Then
        For d = 1 To n_dimension
            x_shrunk(n_new, d) = x(n_raw, d)
        Next d
    End If
    reduceByHalf = x_shrunk
    Erase x_shrunk
End Function


Private Function ExpandedResWindow(lowResPath() As Long, x() As Double, y() As Double, radius As Long) As Long()
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim i_prev As Long, j_prev As Long
Dim i1 As Long, i2 As Long, j1 As Long, j2 As Long
Dim w() As Long
Dim n_x As Long, n_y As Long, n_raw As Long
Dim path2() As Long
Dim visited() As Long

n_x = UBound(x, 1)
n_y = UBound(y, 1)
n_raw = UBound(lowResPath, 1)
ReDim w(1 To n_x, 1 To 2)
ReDim visited(1 To n_x, 1 To n_y)

m = 0
ReDim path2(1 To 2, 1 To min2(n_x * n_y, n_raw * 4 * (1 + radius) ^ 2))
i_prev = 1
j_prev = 1
For k = 1 To n_raw
    i = lowResPath(k, 1)
    j = lowResPath(k, 2)

    If i = i_prev And j > j_prev Then
        i1 = max2(2 * i - 1 - radius, 1)
        i2 = min2(2 * i + radius, n_x)
        j1 = 2 * j - 1
        j2 = min2(2 * j + radius, n_y)
    ElseIf i > i_prev And j = j_prev Then
        i1 = 2 * i - 1
        i2 = min2(2 * i + radius, n_x)
        j1 = max2(2 * j - 1 - radius, 1)
        j2 = min2(2 * j + radius, n_y)
    Else
        i1 = max2(2 * i - 1 - radius, 1)
        i2 = min2(2 * i + radius, n_x)
        j1 = max2(2 * j - 1 - radius, 1)
        j2 = min2(2 * j + radius, n_y)
    End If
    i_prev = i
    j_prev = j
    
    For i = i1 To i2
        For j = j1 To j2
            If visited(i, j) = 0 Then
                visited(i, j) = 1
                m = m + 1
                path2(1, m) = i
                path2(2, m) = j
            End If
        Next j
    Next i
Next k
ReDim Preserve path2(1 To 2, 1 To m)

For k = 1 To m
    i = path2(1, k)
    j = path2(2, k)
    If j < w(i, 1) Or w(i, 1) = 0 Then w(i, 1) = j
    If j > w(i, 2) Then w(i, 2) = j
Next k

ExpandedResWindow = w
Erase visited, path2, w
End Function


'==========================================================================================
'"Exact indexing of dynamic time warping", Eamonn Keogh (2004)
'"Lower-Bounding of Dynamic Time Warping Distances for Multivariate Time Series", Toni M. rath
'==========================================================================================
'Return the starting index postion of a segment in y() that best matches x()
Function LB_Sequential_Scan(x() As Double, y() As Double, w As Long, Optional step_size As Long = 1, _
            Optional dist_type As Long = 2) As Long
Dim i As Long, j As Long, k As Long, d As Long, m As Long, n As Long
Dim n_x As Long, n_raw As Long, n_dimension As Long
Dim tmp_x As Double
Dim best_so_far As Double, LB_dist As Double, true_dist As Double
Dim yt() As Double, xU() As Double, xL() As Double
    n_x = UBound(x, 1)
    n_raw = UBound(y, 1)
    n_dimension = UBound(x, 2)
    ReDim yt(1 To n_x, 1 To n_dimension)
    
    'Upper and lower bound of x
    Call Calc_xUL(x, xU, xL, w)
    
    best_so_far = Exp(70)
    m = 0
    For n = 1 To n_raw - n_x + 1 Step step_size
    
        m = m + 1
        If m Mod 500 = 0 Then
            DoEvents
            Application.StatusBar = "mDTW:LB_Sequential_Scan: " & n & "/" & n_raw - n_x + 1
        End If
        
        'Extract segment from y() as yt()
        yt = sub_seq(y, n, n + n_x - 1)
        
        '===================
        'Normalize yt()
        '===================
'        tmp_x = yt(1, 1)
'        For i = 1 To n_x
'            yt(i, 1) = yt(i, 1) - tmp_x
'        Next i
    
        LB_dist = LB_Keogh(xU, xL, yt, w, dist_type)
        If LB_dist < best_so_far Then
            true_dist = DTW(x, yt, w, dist_type)
            If true_dist < best_so_far Then
                best_so_far = true_dist
                LB_Sequential_Scan = n
            End If
        End If
        
    Next n
    
    Application.StatusBar = False
End Function


Private Function LB_Keogh(xU() As Double, xL() As Double, y() As Double, r As Long, dist_type As Long) As Double
Dim i As Long, j As Long, k As Long, d As Long
Dim n_x As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double, INFINITY As Double
Dim LB As Double
    n_x = UBound(xU, 1)
    If UBound(y, 1) <> n_x Then
        Debug.Print "Both series need to be of same length with LB_Keogh."
        Exit Function
    End If
    n_dimension = UBound(xU, 2)
    LB_Keogh = 0
    If dist_type = 2 Then
        For i = 1 To n_x
            tmp_x = 0
            For d = 1 To n_dimension
                If y(i, d) > xU(i, d) Then
                    tmp_x = tmp_x + (y(i, d) - xU(i, d)) ^ 2
                ElseIf y(i, d) < xL(i, d) Then
                    tmp_x = tmp_x + (y(i, d) - xL(i, d)) ^ 2
                End If
            Next d
            LB_Keogh = LB_Keogh + Sqr(tmp_x)
        Next i
    ElseIf dist_type = 1 Then
        For i = 1 To n_x
            For d = 1 To n_dimension
                If y(i, d) > xU(i, d) Then
                    LB_Keogh = LB_Keogh + Abs(y(i, d) - xU(i, d))
                ElseIf y(i, d) < xL(i, d) Then
                    LB_Keogh = LB_Keogh + Abs(y(i, d) - xL(i, d))
                End If
            Next d
        Next i
    ElseIf dist_type = 3 Then
        For i = 1 To n_x
            For d = 1 To n_dimension
                If y(i, d) > xU(i, d) Then
                    LB_Keogh = LB_Keogh + (y(i, d) - xU(i, d)) ^ 2
                ElseIf y(i, d) < xL(i, d) Then
                    LB_Keogh = LB_Keogh + (y(i, d) - xL(i, d)) ^ 2
                End If
            Next d
        Next i
    End If
End Function

'calculate upper and lower binding function of x() to be used in LB_Keogh
Private Sub Calc_xUL(x() As Double, xU() As Double, xL() As Double, r As Long)
Dim i As Long, j As Long, k As Long, d As Long
Dim n_x As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double, INFINITY As Double
INFINITY = Exp(70)
n_x = UBound(x, 1)
n_dimension = UBound(x, 2)
ReDim xU(1 To n_x, 1 To n_dimension)
ReDim xL(1 To n_x, 1 To n_dimension)
For d = 1 To n_dimension
    For i = 1 To n_x
        xU(i, d) = -INFINITY
        xL(i, d) = INFINITY
        For j = -r To r
            k = i + j
            If k < 1 Then
                k = 1
            ElseIf k > n_x Then
                k = n_x
            End If
            tmp_x = x(k, d)
            If tmp_x > xU(i, d) Then xU(i, d) = tmp_x
            If tmp_x < xL(i, d) Then xL(i, d) = tmp_x
        Next j
    Next i
Next d
End Sub



'=== Return top N subsequences that match queries q()
'Input: query series q() with length n_q
'Input: streaming series x() with length n_x
'Input: n_match, number of best matches to return
'Input: block, range to block out from an identified minimum to avoid overlapping output
'Output: qx_idx(1 to n_match, 1 to 2), starting and ending index position of subsequences
'Output: qx_dist(1 to n_match), distance of q to subsequences
Sub SPRING(q() As Double, x() As Double, n_match As Long, qx_idx() As Long, qx_dist() As Double, Optional block As Long = 0)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, j_lo As Long, j_hi As Long, i_min As Long
Dim n_q As Long, n_x As Long
Dim tmp_min As Double, INFINITY As Double
Dim dist() As Double, delta() As Double
Dim sp() As Long

n_q = UBound(q, 1)
n_x = UBound(x, 1)
INFINITY = Exp(70)
ReDim qx_idx(1 To n_match, 1 To 2)
ReDim qx_dist(1 To n_match)

tmp_min = DTW_SPRING(q, x, dist, sp)

ReDim delta(1 To n_x)
For j = 1 To n_x
    delta(j) = dist(n_q, j)
Next j
Erase dist

For n = 1 To n_match
    
    'Find end point on y() that gives the shortest distance
    tmp_min = INFINITY
    For j = 1 To n_x
        If delta(j) < tmp_min Then
            tmp_min = delta(j)
            k = j
        End If
    Next j
    
    qx_idx(n, 1) = sp(k)
    qx_idx(n, 2) = k
    qx_dist(n) = Sqr(delta(k)) / n_q
    
    'Block out vicinity of current k
    j_hi = min2(n_x, k + block)
    Do While j_hi < n_x
        If delta(j_hi + 1) < delta(j_hi) Then Exit Do
        j_hi = j_hi + 1
    Loop
    j_lo = max2(1, k - block)
    Do While j_lo > 1
        If delta(j_lo - 1) < delta(j_lo) Then Exit Do
        j_lo = j_lo - 1
    Loop
    For j = j_lo To j_hi
        delta(j) = INFINITY
    Next j
    
Next n

End Sub


Function DTW_SPRING(q() As Double, x() As Double, dist() As Double, sp() As Long) As Double
Dim i As Long, j As Long, k As Long
Dim n_q As Long, n_x As Long, mn_idx As Long
Dim d() As Double
Dim cost As Double, INFINITY As Double, mn As Double
Dim s() As Long

INFINITY = Exp(70)
n_q = UBound(q, 1)
n_x = UBound(x, 1)
ReDim dist(1 To n_q, 1 To n_x)
ReDim s(0 To n_q, 1 To n_x)
ReDim d(0 To n_q, 0 To n_x)

For i = 1 To n_q
    d(i, 0) = INFINITY
Next i

For i = 1 To n_x
    s(0, i) = i
Next i

For i = 1 To n_q
    For j = 1 To n_x
        Call min3_val_idx(d(i - 1, j), d(i, j - 1), d(i - 1, j - 1), mn, mn_idx)
        d(i, j) = cost_ij(q, x, i, j) + mn
        If mn_idx = 1 Then
            s(i, j) = s(i - 1, j)
        ElseIf mn_idx = 2 Then
            s(i, j) = s(i, j - 1)
        ElseIf mn_idx = 3 Then
            s(i, j) = s(i - 1, j - 1)
        End If
    Next j
Next i

DTW_SPRING = Sqr(d(n_q, n_x)) / n_q

For i = 1 To n_q
    For j = 1 To n_x
        dist(i, j) = d(i, j)
    Next j
Next i

ReDim sp(1 To n_x)
For i = 1 To n_x
    sp(i) = s(n_q, i)
Next i

End Function


'=== Return top N subsequences that match queries q() and of the same length as q()
'Input: query series q() with length n_q
'Input: streaming series x() with length n_x
'Input: n_match, number of best matches to return
'Input w_window, global constraints on warp path
'Input: step_size to skip when scanning
'Output: qx_idx(), starting index position of subsequences
'Output: qx_dist(), distance of q to subsequences
Sub Sliding_Scan(q() As Double, x() As Double, n_match As Long, qx_idx() As Long, qx_dist() As Double, _
            Optional w_window As Variant, Optional step_size As Long = 1, Optional dist_type As Long = 2)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim n_q As Long, n_x As Long
Dim xt() As Double
Dim qx_idx_tmp() As Long

    n_q = UBound(q, 1)
    n_x = UBound(x, 1)
    ReDim qx_idx_tmp(1 To Int((n_x - n_q + 1) * 1# / step_size) + 1)
    ReDim qx_dist(1 To Int((n_x - n_q + 1) * 1# / step_size) + 1)
    
    n = 0
    For k = n_x - n_q + 1 To 1 Step -step_size
        xt = sub_seq(x, k, k + n_q - 1)
        n = n + 1
        qx_idx_tmp(n) = k
        If IsMissing(w_window) Then
            qx_dist(n) = DTW(q, xt, , dist_type) / (2 * n_q)
        Else
            qx_dist(n) = DTW(q, xt, w_window, dist_type) / (2 * n_q)
        End If
    Next k
    
    ReDim Preserve qx_idx_tmp(1 To n)
    ReDim Preserve qx_dist(1 To n)
    
    'Retrieve only the best n_match
    Call modMath.Sort_Quick_A(qx_dist, 1, n, qx_idx_tmp, 0)
    ReDim Preserve qx_dist(1 To n_match)
    ReDim Preserve qx_idx_tmp(1 To n_match)
    ReDim qx_idx(1 To n_match, 1 To 2)
    For i = 1 To n_match
        qx_idx(i, 1) = qx_idx_tmp(i)
        qx_idx(i, 2) = qx_idx_tmp(i) + n_q - 1
    Next i
End Sub


'=== Return top N subsequences that match queries q()
'=== A version of SPRING that allows contraints on warping path
'Input: query series q() with length n_q
'Input: streaming series x() with length n_x
'Input: n_match, number of best matches to return
'Input: block, range to block out from an identified minimum to avoid overlapping output
'Output: qx_idx(1 to n_match, 1 to 2), starting and ending index position of subsequences
'Output: qx_dist(1 to n_match), distance of q to subsequences
Sub ASM(q() As Double, x() As Double, n_match As Long, qx_idx() As Long, qx_dist() As Double, w_window As Long, Optional block As Long = 0)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, j_lo As Long, j_hi As Long, i_min As Long
Dim n_count As Long
Dim n_q As Long, n_x As Long
Dim tmp_x As Double, tmp_min As Double, INFINITY As Double
Dim dist() As Double, delta() As Double
Dim sp() As Long

n_q = UBound(q, 1)
n_x = UBound(x, 1)
INFINITY = Exp(70)
ReDim qx_idx(1 To n_match, 1 To 2)
ReDim qx_dist(1 To n_match)

tmp_x = DTW_ASM(q, x, w_window, dist, sp)

ReDim delta(1 To n_x)
For j = 1 To n_x
    delta(j) = dist(n_q, j)
Next j

For n = 1 To n_match
    
    'Find end point on y() that gives the shortest distance
    tmp_min = INFINITY
    For j = 1 To n_x
        If delta(j) < tmp_min Then
            tmp_min = delta(j)
            k = j
        End If
    Next j
    
    qx_idx(n, 1) = sp(k)
    qx_idx(n, 2) = k
    qx_dist(n) = Sqr(delta(k)) / n_q
    
    'Block out vicinity of current k
    j_hi = min2(n_x, k + block)
    Do While j_hi < n_x
        If delta(j_hi + 1) < delta(j_hi) Then Exit Do
        j_hi = j_hi + 1
    Loop
    j_lo = max2(1, k - block)
    Do While j_lo > 1
        If delta(j_lo - 1) < delta(j_lo) Then Exit Do
        j_lo = j_lo - 1
    Loop
    For j = j_lo To j_hi
        delta(j) = INFINITY
    Next j
    
Next n

End Sub


Function DTW_ASM(x() As Double, y() As Double, w As Long, dist() As Double, sp() As Long) As Double
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim n_x As Long, n_y As Long
Dim d() As Double
Dim cost As Double, d_best As Double, INFINITY As Double
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim i_min As Long, j_lo As Long, j_hi As Long
Dim MSM() As Long

INFINITY = Exp(20)

n_x = UBound(x, 1)
n_y = UBound(y, 1)

ReDim d(0 To n_x, 0 To n_y)
ReDim MSM(0 To n_x, 0 To n_y, 1 To 3)
For i = 1 To n_x
    d(i, 0) = INFINITY
Next i
For i = 1 To n_y
    MSM(0, i, 3) = i
Next i

For i = 1 To n_x
    For j = 1 To n_y
        
        If Abs(MSM(i - 1, j - 1, 1) - MSM(i - 1, j - 1, 2)) > w Then
            tmp_x = INFINITY
        Else
            tmp_x = d(i - 1, j - 1)
        End If
        
        If Abs(MSM(i, j - 1, 1) + 1 - MSM(i, j - 1, 2)) > w Then
            tmp_y = INFINITY
        Else
            tmp_y = d(i, j - 1)
        End If
        
        If Abs(MSM(i - 1, j, 1) - MSM(i - 1, j, 2) - 1) > w Then
            tmp_z = INFINITY
        Else
            tmp_z = d(i - 1, j)
        End If
        
        Call min3_val_idx(tmp_x, tmp_y, tmp_z, d_best, i_min)
        
        d(i, j) = cost_ij(x, y, i, j) + d_best
            
        If i_min = 1 Then
            MSM(i, j, 1) = MSM(i - 1, j - 1, 1) + 1
            MSM(i, j, 2) = MSM(i - 1, j - 1, 2) + 1
            MSM(i, j, 3) = MSM(i - 1, j - 1, 3)
        ElseIf i_min = 2 Then
            MSM(i, j, 1) = MSM(i, j - 1, 1) + 1
            MSM(i, j, 2) = MSM(i, j - 1, 2)
            MSM(i, j, 3) = MSM(i, j - 1, 3)
        ElseIf i_min = 3 Then
            MSM(i, j, 1) = MSM(i - 1, j, 1)
            MSM(i, j, 2) = MSM(i - 1, j, 2) + 1
            MSM(i, j, 3) = MSM(i - 1, j, 3)
        End If
        
    Next j
Next i

DTW_ASM = Sqr(d(n_x, n_y))

ReDim dist(1 To n_x, 1 To n_y)
For i = 1 To n_x
    For j = 1 To n_y
        dist(i, j) = d(i, j)
    Next j
Next i

ReDim sp(1 To n_y)
For i = 1 To n_y
    sp(i) = MSM(n_x, i, 3)
Next i

End Function
