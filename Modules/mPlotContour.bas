Attribute VB_Name = "mPlotContour"
Option Explicit

'Requires: cDelaunay, cQuadEdge, cQuadEdge_Quad

'==========================================================================================
'Use Marching Squares algorithm to create 2D contour plot in Excel
'Plot_Contour_2D() is used to plot a function with analytic form
'Plot_Contour_2DMatrix() is used to plot a 2D-Matrix of regularly spaced grid
'To use Plot_Contour_2D(), the private function g_func_xy() needs to be changed here manually.
'Arguments:
'min_tgt,   lowest level of contour line to show. In matrix mode, when set to NULL, it is automatically set
'           to a fraction away from the lowest value in the matrix.
'max_tgt,   similar to min_tgt
'n_tgt,     n_tgt gives the number of contour lines to be plot between min_tgt and max_tgt
'min_x/y,   range of x/y coordinates to show. In matrix mode, this specifies the coordinates that correspond
'           to the leftmost/bottommost elements.
'max_x/y,   similar to min_x/y
'n_x/y,     only in function mode, number of sampling points the specified bounds. Higehr value gives smoother
'           lines
'isRescale, only in matrix mode, used in conjunction with min_x/y and max_x/y.
'isClean,   default is FALSE and line segments are not joined together. When set to TRUE, line segments with
'           matching heads and tails are join together, eliminating empty spaces in the output array.
'isSeparate, default is FALSE and all contour lines are output in the same two columns to chart as one series
'            when set to TRUE, output is an array of 2xM columns where M is the number of levels. Each set
'            of two columns can be chart as separate series with headings indicating its level.
'==========================================================================================

'==========================================================================================
'Define g_func_xy()
'==========================================================================================
Private Function g_func_xy(x As Double, y As Double) As Double
    'Example 1
    g_func_xy = Exp(-x ^ 2 - y ^ 2 - x * y) _
                + Exp(-(x - 1) ^ 2 - (y - 2) ^ 2 + 0.7 * (x - 1) * (y - 2)) _
                + Exp(-(x + 2) ^ 2 - (y - 2) ^ 2) _
                + Exp(-(x - 2) ^ 2 - (y + 2) ^ 2 + 0.5 * (x - 2) * (y + 2)) _
                + Exp(-(x + 1) ^ 2 - (y + 1) ^ 2)
End Function


'==========================================================================================
'Contour plot of 2D function defined g_func_xy()
'==========================================================================================
Function Plot_Contour_2D(min_tgt As Double, max_tgt As Double, _
            Optional min_x As Double = -3, Optional max_x As Double = 3, _
            Optional min_y As Double = -3, Optional max_y As Double = 3, _
            Optional n_tgt As Long = 9, Optional n_x As Long = 50, Optional n_y As Long = 50, _
            Optional isClean As Boolean = False, Optional isSeparate As Boolean = False) As Variant
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim x_min() As Double, x_max() As Double, dx() As Double, x_bin() As Long
Dim tgt_value As Double, tgt_interval As Double
Dim vArr As Variant, uArr As Variant
    
    Plot_Contour_2D = VBA.CVErr(xlErrNA)
    
    ReDim x_min(1 To 2)
    ReDim x_max(1 To 2)
    ReDim x_bin(1 To 2)
    ReDim dx(1 To 2)
    x_min(1) = min_x: x_max(1) = max_x: x_bin(1) = n_x
    x_min(2) = min_y: x_max(2) = max_y: x_bin(2) = n_y

    dx(1) = (x_max(1) - x_min(1)) / x_bin(1)
    dx(2) = (x_max(2) - x_min(2)) / x_bin(2)
    
    If max_tgt < min_tgt Then
        Debug.Print "Plot_Contour_2D: min_tgt needs to be smaller than max_tgt. (" & min_tgt & ", " & max_tgt & ")"
        Exit Function
    End If
    
    If n_tgt > 1 Then
        tgt_interval = (max_tgt - min_tgt) / (n_tgt - 1)
    ElseIf n_tgt = 1 Then
        tgt_interval = 0
    End If
    
    If isSeparate = False Then
    
        ReDim vArr(1 To 2, 1 To 1)
        For i = 1 To n_tgt
            tgt_value = min_tgt + (i - 1) * tgt_interval
            Call Plot_Contour_2D_isoline(vArr, tgt_value, x_min, x_max, x_bin, dx)
        Next i
        If UBound(vArr, 2) > 1 Then
            If isClean = True Then
                Call Plot_Contour_2D_CleanUp(vArr)
            End If
            Plot_Contour_2D = Application.WorksheetFunction.Transpose(vArr)
        End If
        
    Else
    
        ReDim vArr(1 To 2 * n_tgt, 0 To 1)
        For i = 1 To n_tgt
            ReDim uArr(1 To 2, 1 To 1)
            tgt_value = min_tgt + (i - 1) * tgt_interval
            Call Plot_Contour_2D_isoline(uArr, tgt_value, x_min, x_max, x_bin, dx)
            m = UBound(uArr, 2)
            If m > 1 Then
                If isClean = True Then
                    Call Plot_Contour_2D_CleanUp(uArr)
                End If
                m = UBound(uArr, 2)
                If m > UBound(vArr, 2) Then ReDim Preserve vArr(1 To 2 * n_tgt, 0 To m)
                vArr(i * 2 - 1, 0) = "p=" & Format(tgt_value, "0.0000")
                For j = 1 To m
                    vArr(i * 2 - 1, j) = uArr(1, j)
                    vArr(i * 2, j) = uArr(2, j)
                Next j
            End If
            Erase uArr
        Next i
        If UBound(vArr, 2) > 1 Then
            Plot_Contour_2D = Application.WorksheetFunction.Transpose(vArr)
        End If
    
    End If
    Erase vArr, x_min, x_max, x_bin, dx
End Function


Private Sub Plot_Contour_2D_isoline(vArr As Variant, tgt_value As Double, _
                x_min() As Double, x_max() As Double, x_bin() As Long, dx() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, k_side As Long
Dim ii As Long, jj As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim bit4 As String
Dim f_corner() As Double, f_row() As Double, f_row_prev() As Double
Dim LookUp2Side() As Long
    
    'This value shows from which side does the previous contour line come in
    '1: top, 2: right, 3: bottom, 4: left
    'if it comes in from the left then current line can be joined togehter as
    'one continuous segment, otherwise start a new segment
    ReDim LookUp2Side(0 To 15)
    LookUp2Side(0) = -1: LookUp2Side(15) = -1
    LookUp2Side(1) = 1: LookUp2Side(14) = 1
    LookUp2Side(2) = 4: LookUp2Side(13) = 4
    LookUp2Side(3) = 4: LookUp2Side(12) = 4
    LookUp2Side(4) = 4: LookUp2Side(11) = 4
    LookUp2Side(5) = -1: LookUp2Side(10) = -1
    LookUp2Side(6) = 1: LookUp2Side(9) = 1
    LookUp2Side(7) = 3: LookUp2Side(8) = 3
    
    'Scan the grids from top to bottm, left to right
    'f_corner(1:4) cache the values of the 4 corners after 1 move.
    '1 is the upper-left, 2 is upper right, 3 is lower right, 4 is lower left
    'f_row() cache the values of the lower row after 1-pass of horizontal scan
    ReDim f_row(0 To x_bin(1))
    For j = x_bin(2) To 1 Step -1
        tmp_y = x_min(2) + (j - 1) * dx(2)
        k_side = -1
        For i = 1 To x_bin(1)
            tmp_x = x_min(1) + (i - 1) * dx(1)
            
            'Calculate the values at each of the 4 corners
            If i > 1 And j = x_bin(2) Then
                f_corner(1) = f_corner(2)
                f_corner(4) = f_corner(3)
                f_corner(2) = g_func_xy(tmp_x + dx(1), tmp_y + dx(2)) '*** Function
                f_corner(3) = g_func_xy(tmp_x + dx(1), tmp_y)         '*** Function
            ElseIf i > 1 And j < x_bin(2) Then
                f_corner(1) = f_corner(2)
                f_corner(4) = f_corner(3)
                f_corner(2) = f_row_prev(i)
                f_corner(3) = g_func_xy(tmp_x + dx(1), tmp_y)         '*** Function
            Else
                ReDim f_corner(1 To 4): k = 0
                For ii = 1 To 0 Step -1
                    For jj = 0 To 1
                        If ii = 1 Then
                            tmp_z = g_func_xy(tmp_x + jj * dx(1), tmp_y + ii * dx(2))        '*** Function
                        Else
                            tmp_z = g_func_xy(tmp_x + (1 - jj) * dx(1), tmp_y + ii * dx(2))  '*** Function
                        End If
                        k = k + 1: f_corner(k) = tmp_z
                    Next jj
                Next ii
            End If
            
            If i = 1 Then f_row(i - 1) = f_corner(4)
            f_row(i) = f_corner(3)
            
            'Determine which template to use and add the line segment
            k = 0
            For ii = 1 To 4
                If f_corner(ii) >= tgt_value Then k = k + 2 ^ (4 - ii)
            Next ii
            If k <> 0 And k <> 15 Then
                Call Add_Lines(vArr, k, tmp_x, tmp_y, dx, f_corner, tgt_value, k_side)
            End If
            k_side = LookUp2Side(k)
        Next i
        f_row_prev = f_row
    Next j
    Erase f_corner, f_row, f_row_prev, LookUp2Side
End Sub


'Conventions used here follow the ones shown on Wikipedia
Private Sub Add_Lines(vArr As Variant, i_lkup As Long, tmp_x As Double, tmp_y As Double, dx() As Double, _
        f_corner() As Double, tgt_value As Double, k_side As Long)
Dim n As Long, f_center As Double

    n = UBound(vArr, 2)
    Select Case i_lkup
        Case 1, 14
            If k_side = 4 Then
                ReDim Preserve vArr(1 To 2, 1 To n + 1)
                vArr(1, n - 1) = tmp_x + dx(1) * (tgt_value - f_corner(4)) / (f_corner(3) - f_corner(4))
                vArr(2, n - 1) = tmp_y
            Else
                ReDim Preserve vArr(1 To 2, 1 To n + 3)
                vArr(1, n) = tmp_x
                vArr(2, n) = tmp_y + dx(2) * (tgt_value - f_corner(4)) / (f_corner(1) - f_corner(4))
                vArr(1, n + 1) = tmp_x + dx(1) * (tgt_value - f_corner(4)) / (f_corner(3) - f_corner(4))
                vArr(2, n + 1) = tmp_y
            End If
        Case 2, 13
'            If k_side = 3 Then
'                ReDim Preserve vArr(1 To 2, 1 To n + 1)
'                vArr(1, n - 1) = tmp_x + dx(1)
'                vArr(2, n - 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))
'            Else
                ReDim Preserve vArr(1 To 2, 1 To n + 3)
                vArr(1, n) = tmp_x + dx(1) * (tgt_value - f_corner(4)) / (f_corner(3) - f_corner(4))
                vArr(2, n) = tmp_y
                vArr(1, n + 1) = tmp_x + dx(1)
                vArr(2, n + 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))
'            End If
        Case 3, 12
            If k_side = 4 Then
                ReDim Preserve vArr(1 To 2, 1 To n + 1)
                vArr(1, n - 1) = tmp_x + dx(1)
                vArr(2, n - 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))
            Else
                ReDim Preserve vArr(1 To 2, 1 To n + 3)
                vArr(1, n) = tmp_x
                vArr(2, n) = tmp_y + dx(2) * (tgt_value - f_corner(4)) / (f_corner(1) - f_corner(4))
                vArr(1, n + 1) = tmp_x + dx(1)
                vArr(2, n + 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))
            End If
        Case 4, 11
'            If k_side = 1 Then
'                ReDim Preserve vArr(1 To 2, 1 To n + 1)
'                vArr(1, n - 1) = tmp_x + dx(1)
'                vArr(2, n - 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))
'            Else
                ReDim Preserve vArr(1 To 2, 1 To n + 3)
                vArr(1, n) = tmp_x + dx(1) * (tgt_value - f_corner(1)) / (f_corner(2) - f_corner(1))
                vArr(2, n) = tmp_y + dx(2)
                vArr(1, n + 1) = tmp_x + dx(1)
                vArr(2, n + 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))
'            End If
        Case 5, 10
            f_center = (f_corner(1) + f_corner(2) + f_corner(3) + f_corner(4)) / 4
            If (f_center >= tgt_value And i_lkup = 5) Or (f_center < tgt_value And i_lkup = 10) Then
                ReDim Preserve vArr(1 To 2, 1 To n + 6)
                vArr(1, n) = tmp_x + dx(1) * (tgt_value - f_corner(4)) / (f_corner(3) - f_corner(4))
                vArr(2, n) = tmp_y
                vArr(1, n + 1) = tmp_x + dx(1)
                vArr(2, n + 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))

                vArr(1, n + 3) = tmp_x
                vArr(2, n + 3) = tmp_y + dx(2) * (tgt_value - f_corner(4)) / (f_corner(1) - f_corner(4))
                vArr(1, n + 4) = tmp_x + dx(1) * (tgt_value - f_corner(1)) / (f_corner(2) - f_corner(1))
                vArr(2, n + 4) = tmp_y + dx(2)
            ElseIf (f_center >= tgt_value And i_lkup = 10) Or (f_center < tgt_value And i_lkup = 5) Then
                ReDim Preserve vArr(1 To 2, 1 To n + 6)
                vArr(1, n) = tmp_x + dx(1) * (tgt_value - f_corner(1)) / (f_corner(2) - f_corner(1))
                vArr(2, n) = tmp_y + dx(2)
                vArr(1, n + 1) = tmp_x + dx(1)
                vArr(2, n + 1) = tmp_y + dx(2) * (tgt_value - f_corner(3)) / (f_corner(2) - f_corner(3))
                
                vArr(1, n + 3) = tmp_x
                vArr(2, n + 3) = tmp_y + dx(2) * (tgt_value - f_corner(4)) / (f_corner(1) - f_corner(4))
                vArr(1, n + 4) = tmp_x + dx(1) * (tgt_value - f_corner(4)) / (f_corner(3) - f_corner(4))
                vArr(2, n + 4) = tmp_y
            End If
        Case 6, 9
'            If k_side = 1 Then
'                ReDim Preserve vArr(1 To 2, 1 To n + 1)
'                vArr(1, n - 1) = tmp_x + dx(1) * (tgt_value - f_corner(4)) / (f_corner(3) - f_corner(4))
'                vArr(2, n - 1) = tmp_y
'            Else
                ReDim Preserve vArr(1 To 2, 1 To n + 3)
                vArr(1, n) = tmp_x + dx(1) * (tgt_value - f_corner(1)) / (f_corner(2) - f_corner(1))
                vArr(2, n) = tmp_y + dx(2)
                vArr(1, n + 1) = tmp_x + dx(1) * (tgt_value - f_corner(4)) / (f_corner(3) - f_corner(4))
                vArr(2, n + 1) = tmp_y
'            End If
        Case 7, 8
            If k_side = 4 Then
                ReDim Preserve vArr(1 To 2, 1 To n + 1)
                vArr(1, n - 1) = tmp_x + dx(1) * (tgt_value - f_corner(1)) / (f_corner(2) - f_corner(1))
                vArr(2, n - 1) = tmp_y + dx(2)
            Else
                ReDim Preserve vArr(1 To 2, 1 To n + 3)
                vArr(1, n) = tmp_x
                vArr(2, n) = tmp_y + dx(2) * (tgt_value - f_corner(4)) / (f_corner(1) - f_corner(4))
                vArr(1, n + 1) = tmp_x + dx(1) * (tgt_value - f_corner(1)) / (f_corner(2) - f_corner(1))
                vArr(2, n + 1) = tmp_y + dx(2)
            End If
    End Select
End Sub

Private Function Bin2Dec(BinaryString As String) As Long
Dim i As Long, k As Long
    Bin2Dec = 0
    k = Len(BinaryString)
    For i = 1 To k
        If Mid(BinaryString, k - i + 1, 1) = "1" Then
            Bin2Dec = Bin2Dec + 2 ^ (i - 1)
        End If
    Next i
End Function


Private Sub Plot_Contour_2D_CleanUp(vArr As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_seg As Long, iterate As Long, n_sub As Long
Dim ii As Long, jj As Long, kk As Long
Dim idx_end() As Long, idx_start() As Long, x_start() As Double, x_end() As Double
Dim uArr As Variant
Dim x_anchor() As Double, x_anchor_start() As Double, tol As Double
Dim seg_list As Collection
Dim tmp_x As Double, tmp_y As Double, i_mode As Long
    tol = 0.000000001 '1E-9
    n = UBound(vArr, 2)

    'Identify head and tail of each segment
    n_seg = 0
    i = 1
    ReDim idx_start(1 To n): ReDim x_start(1 To 2, 1 To n)
    ReDim idx_end(1 To n): ReDim x_end(1 To 2, 1 To n)
    Do While i < n
        For ii = i + 1 To n
            If VBA.IsEmpty(vArr(1, ii)) And Not VBA.IsEmpty(vArr(1, ii - 1)) Then
                n_seg = n_seg + 1
                idx_start(n_seg) = i
                idx_end(n_seg) = ii - 1
                x_start(1, n_seg) = vArr(1, i)
                x_start(2, n_seg) = vArr(2, i)
                x_end(1, n_seg) = vArr(1, ii - 1)
                x_end(2, n_seg) = vArr(2, ii - 1)
                i = ii + 1
                Exit For
            End If
        Next ii
        If ii >= n Then Exit Do
    Loop
    ReDim Preserve idx_start(1 To n_seg): ReDim Preserve x_start(1 To 2, 1 To n_seg)
    ReDim Preserve idx_end(1 To n_seg): ReDim Preserve x_end(1 To 2, 1 To n_seg)
    
    'Build list of segments
    Set seg_list = New Collection
    For i = 1 To n_seg
        seg_list.Add i
    Next i
    
    'Merge segments until the list is empty
    ReDim uArr(1 To 2, 1 To n)
    ReDim x_anchor(1 To 2)
    ReDim x_anchor_start(1 To 2)
    k = 0
    Do While seg_list.Count > 0
        DoEvents
        Application.StatusBar = "Consolidating segments: " & seg_list.Count
        
        'Start new segment
        i = seg_list(1): seg_list.Remove (1)
        kk = k + 1
        For ii = idx_start(i) To idx_end(i)
            k = k + 1
            uArr(1, k) = vArr(1, ii)
            uArr(2, k) = vArr(2, ii)
        Next ii
        x_anchor(1) = x_end(1, i)
        x_anchor(2) = x_end(2, i)
        x_anchor_start(1) = x_start(1, i)
        x_anchor_start(2) = x_start(2, i)
        
        'Scan remaining segments that match the new semgment
        For iterate = 1 To 80
            n = seg_list.Count
            m = 0
            Do While m < seg_list.Count
                m = m + 1
                j = seg_list(m)
                If Abs(x_anchor(1) - x_start(1, j)) < tol And Abs(x_anchor(2) - x_start(2, j)) < tol Then
                    For jj = idx_start(j) + 1 To idx_end(j)
                        k = k + 1
                        uArr(1, k) = vArr(1, jj)
                        uArr(2, k) = vArr(2, jj)
                    Next jj
                    x_anchor(1) = x_end(1, j)
                    x_anchor(2) = x_end(2, j)
                    seg_list.Remove (m)
                    m = m - 1
                ElseIf Abs(x_anchor(1) - x_end(1, j)) < tol And Abs(x_anchor(2) - x_end(2, j)) < tol Then
                    For jj = idx_end(j) - 1 To idx_start(j) Step -1
                        k = k + 1
                        uArr(1, k) = vArr(1, jj)
                        uArr(2, k) = vArr(2, jj)
                    Next jj
                    x_anchor(1) = x_start(1, j)
                    x_anchor(2) = x_start(2, j)
                    seg_list.Remove (m)
                    m = m - 1
                ElseIf Abs(x_anchor_start(1) - x_end(1, j)) < tol And Abs(x_anchor_start(2) - x_end(2, j)) < tol Then
                    n_sub = idx_end(j) - idx_start(j) + 1
                    For ii = k + n_sub - 1 To kk + n_sub - 1 Step -1
                        uArr(1, ii) = uArr(1, ii - n_sub + 1)
                        uArr(2, ii) = uArr(2, ii - n_sub + 1)
                    Next ii
                    k = k + n_sub - 1
                    ii = 0
                    For jj = idx_start(j) To idx_end(j) - 1
                        uArr(1, kk + ii) = vArr(1, jj)
                        uArr(2, kk + ii) = vArr(2, jj)
                        ii = ii + 1
                    Next jj
                    x_anchor_start(1) = x_start(1, j)
                    x_anchor_start(2) = x_start(2, j)
                    seg_list.Remove (m)
                    m = m - 1
                ElseIf Abs(x_anchor_start(1) - x_start(1, j)) < tol And Abs(x_anchor_start(2) - x_start(2, j)) < tol Then
                    n_sub = idx_end(j) - idx_start(j) + 1
                    For ii = k + n_sub - 1 To kk + n_sub - 1 Step -1
                        uArr(1, ii) = uArr(1, ii - n_sub + 1)
                        uArr(2, ii) = uArr(2, ii - n_sub + 1)
                    Next ii
                    k = k + n_sub - 1
                    ii = 0
                    For jj = idx_end(j) To idx_start(j) + 1 Step -1
                        uArr(1, kk + ii) = vArr(1, jj)
                        uArr(2, kk + ii) = vArr(2, jj)
                        ii = ii + 1
                    Next jj
                    x_anchor_start(1) = x_end(1, j)
                    x_anchor_start(2) = x_end(2, j)
                    seg_list.Remove (m)
                    m = m - 1
                End If
            Loop
            If seg_list.Count = n Then Exit For
        Next iterate
        k = k + 1
    Loop
    Set seg_list = Nothing
    vArr = uArr
    ReDim Preserve vArr(1 To 2, 1 To k - 1)
    Erase uArr
    Application.StatusBar = False
End Sub


'==========================================================================================
'Contour plot of 2D-matrix defined g_matrix(1:M,1:N)
'By default, element(1,1) is assumed to be at the upperleft-most corner, element(M,N) is at
'the bottomright-most corner. And coordinates are taken to be regular spaced from 0.5 to M/N-0.5.
'If "true" coordinates needs to be shown, then set isRescale to TRUE and specify min/max_x/y values
'that corresponds to center-points of each four corners
'==========================================================================================
Function Plot_Contour_2DMatrix(g_matrix As Variant, _
            Optional min_tgt As Variant = Null, Optional max_tgt As Variant = Null, _
            Optional n_tgt As Long = 9, _
            Optional isClean As Boolean = False, _
            Optional isRescale As Boolean = False, _
            Optional min_x As Variant = Null, Optional max_x As Variant = Null, _
            Optional min_y As Variant = Null, Optional max_y As Variant = Null, _
            Optional isSeparate As Boolean = False) As Variant
Dim i As Long, j As Long, k As Long, m As Long, n As Long, mm As Long
Dim x_min() As Double, x_max() As Double, dx() As Double, x_bin() As Long
Dim tgt_value As Double, tgt_interval As Double, z_max As Double, z_min As Double
Dim vArr As Variant, uArr As Variant
Dim tmp_x As Double, tmp_y As Double
    
    Plot_Contour_2DMatrix = VBA.CVErr(xlErrNA)
    
    m = UBound(g_matrix, 1)
    n = UBound(g_matrix, 2)

    ReDim x_min(1 To 2)
    ReDim x_max(1 To 2)
    ReDim x_bin(1 To 2)
    ReDim dx(1 To 2)
    x_min(1) = 0: x_max(1) = n: x_bin(1) = n
    x_min(2) = 0: x_max(2) = m: x_bin(2) = m
    dx(1) = 1: dx(2) = 1
    
    If VBA.IsNull(min_tgt) Or VBA.IsNull(max_tgt) Then
        z_min = Exp(70)
        z_max = -Exp(70)
        For i = 1 To m
            For j = 1 To n
                If g_matrix(i, j) < z_min Then z_min = g_matrix(i, j)
                If g_matrix(i, j) > z_max Then z_max = g_matrix(i, j)
            Next j
        Next i
    End If
    
    If VBA.IsNull(min_tgt) Then
        z_min = z_min + (z_max - z_min) / n_tgt
    Else
        z_min = min_tgt
    End If
    If VBA.IsNull(max_tgt) Then
        z_max = z_max - (z_max - z_min) / n_tgt
    Else
        z_max = max_tgt
    End If
    
    If z_max < z_min Then
        Debug.Print "Plot_Contour_2DMatrix: min_tgt needs to be smaller than max_tgt. (" & z_min & ", " & z_max & ")"
        Exit Function
    End If
    
    If n_tgt > 1 Then
        tgt_interval = (z_max - z_min) / (n_tgt - 1)
    ElseIf n_tgt = 1 Then
        tgt_interval = 0
    End If
    

    If isSeparate = False Then
    
        ReDim vArr(1 To 2, 1 To 1)
        For i = 1 To n_tgt
            tgt_value = z_min + (i - 1) * tgt_interval
            Call Plot_Contour_2DMatrix_isoline(g_matrix, vArr, tgt_value, x_min, x_max, x_bin, dx)
        Next i
            
        If UBound(vArr, 2) > 1 Then
            If isClean = True Then
                Call Plot_Contour_2D_CleanUp(vArr)
            End If
            If isRescale = True Then
                tmp_x = (max_x - min_x) / (n - 1)
                tmp_y = (max_y - min_y) / (m - 1)
                For i = 1 To UBound(vArr, 2)
                    If Not VBA.IsEmpty(vArr(1, i)) Then
                        vArr(1, i) = min_x + tmp_x * (vArr(1, i) - 0.5)
                        vArr(2, i) = min_y + tmp_y * (vArr(2, i) - 0.5)
                    End If
                Next i
            End If
            Plot_Contour_2DMatrix = Application.WorksheetFunction.Transpose(vArr)
        End If
            
    Else
    
        ReDim vArr(1 To 2 * n_tgt, 0 To 1)
        For i = 1 To n_tgt
            ReDim uArr(1 To 2, 1 To 1)
            tgt_value = z_min + (i - 1) * tgt_interval
            Call Plot_Contour_2DMatrix_isoline(g_matrix, uArr, tgt_value, x_min, x_max, x_bin, dx)
            mm = UBound(uArr, 2)
            If mm > 1 Then
                If isClean = True Then
                    Call Plot_Contour_2D_CleanUp(uArr)
                End If
                If isRescale = True Then
                    tmp_x = (max_x - min_x) / (n - 1)
                    tmp_y = (max_y - min_y) / (m - 1)
                    For j = 1 To UBound(uArr, 2)
                        If Not VBA.IsEmpty(uArr(1, j)) Then
                            uArr(1, j) = min_x + tmp_x * (uArr(1, j) - 0.5)
                            uArr(2, j) = min_y + tmp_y * (uArr(2, j) - 0.5)
                        End If
                    Next j
                End If
                mm = UBound(uArr, 2)
                If mm > UBound(vArr, 2) Then ReDim Preserve vArr(1 To 2 * n_tgt, 0 To mm)
                vArr(i * 2 - 1, 0) = "p=" & Format(tgt_value, "0.0000")
                For j = 1 To mm
                    vArr(i * 2 - 1, j) = uArr(1, j)
                    vArr(i * 2, j) = uArr(2, j)
                Next j
            End If
            Erase uArr
        Next i
        If UBound(vArr, 2) > 1 Then
            Plot_Contour_2DMatrix = Application.WorksheetFunction.Transpose(vArr)
        End If
    
    End If

    Erase vArr, x_min, x_max, x_bin, dx
End Function


Private Sub Plot_Contour_2DMatrix_isoline(g_matrix As Variant, vArr As Variant, tgt_value As Double, _
                x_min() As Double, x_max() As Double, x_bin() As Long, dx() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, k_side As Long
Dim ii As Long, jj As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim bit4 As String
Dim f_corner() As Double, f_row() As Double, f_row_prev() As Double
Dim LookUp2Side() As Long
    
    m = UBound(g_matrix, 1)
    n = UBound(g_matrix, 2)
    
    ReDim LookUp2Side(0 To 15)
    LookUp2Side(0) = -1: LookUp2Side(15) = -1
    LookUp2Side(1) = 1: LookUp2Side(14) = 1
    LookUp2Side(2) = 4: LookUp2Side(13) = 4
    LookUp2Side(3) = 4: LookUp2Side(12) = 4
    LookUp2Side(4) = 4: LookUp2Side(11) = 4
    LookUp2Side(5) = -1: LookUp2Side(10) = -1
    LookUp2Side(6) = 1: LookUp2Side(9) = 1
    LookUp2Side(7) = 3: LookUp2Side(8) = 3

    ReDim f_row(0 To n)
    For j = 1 To m - 1
        tmp_y = m - j - 0.5
        k_side = -1
        For i = 1 To n - 1
            tmp_x = i - 0.5
            
            If i > 1 And j = 1 Then
                f_corner(1) = f_corner(2)
                f_corner(4) = f_corner(3)
                f_corner(2) = g_matrix(j, i + 1) 'Function
                f_corner(3) = g_matrix(j + 1, i + 1)   'Function
            ElseIf i > 1 And j > 1 Then
                f_corner(1) = f_corner(2)
                f_corner(4) = f_corner(3)
                f_corner(2) = f_row_prev(i)
                f_corner(3) = g_matrix(j + 1, i + 1)    'Function
            Else
                ReDim f_corner(1 To 4): k = 0
                For jj = 0 To 1
                    For ii = 0 To 1
                        If jj = 0 Then
                            tmp_z = g_matrix(j + jj, i + ii)
                        Else
                            tmp_z = g_matrix(j + jj, i + (1 - ii))
                        End If
                        k = k + 1: f_corner(k) = tmp_z
                    Next ii
                Next jj
            End If

            If i = 1 Then f_row(i - 1) = f_corner(4)
            f_row(i) = f_corner(3)
            
            k = 0
            For ii = 1 To 4
                If f_corner(ii) >= tgt_value Then k = k + 2 ^ (4 - ii)
            Next ii
            If k <> 0 And k <> 15 Then
                Call Add_Lines(vArr, k, tmp_x, tmp_y, dx, f_corner, tgt_value, k_side)
            End If
            k_side = LookUp2Side(k)
        Next i
        f_row_prev = f_row
    Next j
End Sub


'==========================================================================================
'Contour plot of datasets given by xy(1:N,1:2) and z(1:N), where xy() is the cartesian coordinates
'of N datapoints and z() is the "height" at each of these points.
'==========================================================================================
Function Plot_Contour_xyz(xy As Variant, z As Variant, _
            Optional min_tgt As Variant = Null, Optional max_tgt As Variant = Null, _
            Optional n_tgt As Long = 9, _
            Optional isClean As Boolean = False, _
            Optional n_x As Long = 9, Optional n_y As Long = 9, _
            Optional min_x As Variant = Null, Optional max_x As Variant = Null, _
            Optional min_y As Variant = Null, Optional max_y As Variant = Null, _
            Optional isSeparate As Boolean = False) As Variant
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim vArr As Variant
Dim xy_d As cDelaunay
Dim tmp_x As Double, tmp_y As Double
Dim x_min As Double, x_max As Double, y_min As Double, y_max As Double
    n = UBound(xy, 1)
    
    'Set up grid that covers xy()
    For j = 1 To 2
        tmp_x = Exp(70): tmp_y = -tmp_x
        For i = 1 To n
            If xy(i, j) < tmp_x Then tmp_x = xy(i, j)
            If xy(i, j) > tmp_y Then tmp_y = xy(i, j)
        Next i
        If j = 1 Then
            If IsNull(min_x) Then
                x_min = tmp_x
            Else
                x_min = min_x
            End If
            If IsNull(max_x) Then
                x_max = tmp_y
            Else
                x_max = max_x
            End If
        ElseIf j = 2 Then
            If IsNull(min_y) Then
                y_min = tmp_x
            Else
                y_min = min_y
            End If
            If IsNull(max_y) Then
                y_max = tmp_y
            Else
                y_max = max_y
            End If
        End If
    Next j
    
    'Intrapolate value at each grid point
    Set xy_d = New cDelaunay
    With xy_d
        Call .Init(xy, z)
        Call .Intrapolate_Grid(vArr, n_x, n_y, x_min, x_max, y_min, y_max)
        Call .Reset
    End With
    Set xy_d = Nothing
    
    'Contour plot
    Plot_Contour_xyz = Plot_Contour_2DMatrix(vArr, min_tgt, max_tgt, n_tgt, isClean, True, _
                            x_min, x_max, y_min, y_max, isSeparate)
                            
    Erase vArr
End Function


Private Sub Test_mPlotContour()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim vArr As Variant
Dim tmp_x As Double
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With ActiveWorkbook.Sheets("Sheet1")
        .Range("B2:C1048576").Clear
        .Range("N1:IV1048576").Clear
        
        vArr = Plot_Contour_2D(0.1, 0.9, -4, 4, -4, 4, 9, 50, 50, True)
        .Range("B2").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr

        vArr = Plot_Contour_2D(0.1, 1.1, -4, 4, -4, 4, 5, 50, 50, True, True)
        .Range("N1").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
    End With
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Private Sub Test_mPlotContour_Matrix()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim vArr As Variant, A As Variant
Dim tmp_x As Double
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    With ActiveWorkbook.Sheets("Sheet2")
        A = .Range("B4:AD44").Value

        'Plot as one series
        vArr = Plot_Contour_2DMatrix(A, 0.2, 1, 7, True, True, -2, 2.2, -2, 2)
        .Range("AH2:AI1048576").Clear
        If Not VBA.IsError(vArr) Then .Range("AH2").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
        
        'Plot each level of contour line separately
        vArr = Plot_Contour_2DMatrix(A, 0.2, 1, 5, True, False, -2, 2.2, -2, 2, True)
        .Range("AK1:AT1048576").Clear
        If Not VBA.IsError(vArr) Then .Range("AK1").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
        
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Private Sub Test_mPlotContour_xyz()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim vArr As Variant, A As Variant, uArr As Variant
Dim tmp_x As Double
Dim xy As Variant, z As Variant
Dim x_d As cDelaunay

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    With ActiveWorkbook.Sheets("Sheet5_xyz")
        n = .Range("B1048576").End(xlUp).Row - 2
        xy = .Range("B3").Resize(n, 2).Value
        ReDim z(1 To n)
        For i = 1 To n
            z(i) = .Range("D" & 2 + i).Value
        Next i
        
        'Plot Delaunay and Voronoi diagrams
        Set x_d = New cDelaunay
        With x_d
            Call .Init(xy, z)
            Call .Plot(vArr, A)
            Call .Plot_Voronoi(uArr)
            Call .Reset
        End With
        Set x_d = Nothing
        .Range("F3:J1048576").Clear
        .Range("F3").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
        .Range("I3").Resize(UBound(uArr, 1), UBound(uArr, 2)).Value = uArr
        .Range("H3").Resize(UBound(A, 1), 1).Value = Application.WorksheetFunction.Transpose(A)
            
        'Plot as one series
        vArr = Plot_Contour_xyz(xy, z, , , , True)
        .Range("L3:M1048576").Clear
        If Not VBA.IsError(vArr) Then .Range("L3").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr

        'Plot each level of contour line separately
        vArr = Plot_Contour_xyz(xy, z, , , 3, True, , , , , , , True)
        .Range("O2:AZ1048576").Clear
        If Not VBA.IsError(vArr) Then .Range("O2").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
        
    End With

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
