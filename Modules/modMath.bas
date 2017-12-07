Attribute VB_Name = "modMath"
Option Explicit

'====================================================
'This module contains commonly used functions like sorting,
'matrix decompositions, random number generator etc. In most cases
'arrays are assumed to be 1-based, unless otherwise specified.
'=====================================================

'========================================
'Sorting & Searching Algorithms
'========================================

'Search for target integer in sorted array x(1 to N), returns positon index if found, -1 if not found
Function Binary_Search(x() As Long, tgt As Long) As Long
Dim i As Long, n As Long, m As Long
Dim iL As Long, iR As Long
    n = UBound(x)
    iL = 1
    iR = n
    Binary_Search = -1
    Do
        If iL > iR Then
            Binary_Search = -1 'target not found
            Exit Do
        ElseIf x(iL) = tgt Then
            Binary_Search = iL
            Exit Do
        ElseIf x(iR) = tgt Then
            Binary_Search = iR
            Exit Do
        Else
            m = Int((iL + iR) / 2)
            If x(m) = tgt Then
                Binary_Search = m
                Exit Do
            ElseIf x(m) < tgt Then
                iL = m + 1
            ElseIf x(m) > tgt Then
                iR = m - 1
            End If
        End If
    Loop
End Function


'Return the k-th smallest value and its respective postion
'Input:  x() and k
'Output: x_min and i_min, single value if output_list is FALSE.
'If output_list is set to TRUE, x_min and i_min return vector of values holding all k values.
Sub Smallest_k(x As Variant, k As Long, x_min As Variant, i_min As Variant, Optional output_list As Boolean = False)
Dim i As Long, j As Long, m As Long, n As Long, minIndex As Long
Dim minValue As Variant, swap As Variant, y As Variant, iArr() As Long
    n = UBound(x)
    y = x
    ReDim iArr(1 To n)
    For i = 1 To n
        iArr(i) = i
    Next i
    For i = 1 To k
        minIndex = i
        minValue = y(i)
        For j = i + 1 To n
            If y(j) < minValue Then
                minIndex = j
                minValue = y(j)
                swap = y(i)
                y(i) = y(minIndex)
                y(minIndex) = swap
                m = iArr(i)
                iArr(i) = iArr(minIndex)
                iArr(minIndex) = m
            End If
        Next j
    Next i
    ReDim Preserve y(1 To k)
    ReDim Preserve iArr(1 To k)
    If output_list = True Then
        x_min = y
        i_min = iArr
    Else
        x_min = y(k)
        i_min = iArr(k)
    End If
    Erase y, iArr
End Sub

Sub Sort_Bubble(x As Variant)
Dim i As Long, n As Long, swap As Long
Dim temp As Variant
    n = UBound(x)
    Do
        swap = 0
        For i = 1 To n - 1
            If x(i + 1) < x(i) Then
                swap = i
                temp = x(i)
                x(i) = x(i + 1)
                x(i + 1) = temp
            End If
        Next i
        n = swap
    Loop Until swap = 0
End Sub

Sub Sort_Bubble_A(x As Variant, sort_index() As Long, Optional first_run As Long = 1)
Dim i As Long, j As Long, n As Long, swap As Long
Dim temp As Variant
    n = UBound(x)
    If first_run = 1 Then
        ReDim sort_index(1 To n)
        For i = 1 To n
            sort_index(i) = i
        Next i
    End If
    Do
        swap = 0
        For i = 1 To n - 1
            If x(i + 1) < x(i) Then
                swap = i
                
                temp = x(i)
                x(i) = x(i + 1)
                x(i + 1) = temp
                
                j = sort_index(i)
                sort_index(i) = sort_index(i + 1)
                sort_index(i + 1) = j
            End If
        Next i
        n = swap
    Loop Until swap = 0
End Sub


Sub Sort_Quick(vArray As Variant, inLow As Long, inHi As Long)
Dim pivot As Double
Dim tmpSwap As Variant
Dim tmpLow As Long, tmpHi As Long
    tmpLow = inLow
    tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then Sort_Quick vArray, inLow, tmpHi
    If (tmpLow < inHi) Then Sort_Quick vArray, tmpLow, inHi
End Sub


Sub Sort_Quick_A(vArray As Variant, inLow As Long, inHi As Long, sort_index() As Long, Optional first_run As Long = 1)
Dim pivot   As Double
Dim tmpSwap As Variant
Dim tmpLow As Long, tmpHi As Long, i As Long
    If first_run = 1 Then
        ReDim sort_index(LBound(vArray) To UBound(vArray))
        For i = LBound(vArray) To UBound(vArray)
            sort_index(i) = i
        Next i
    End If
    tmpLow = inLow
    tmpHi = inHi
    pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            
            i = sort_index(tmpLow)
            sort_index(tmpLow) = sort_index(tmpHi)
            sort_index(tmpHi) = i
            
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then Sort_Quick_A vArray, inLow, tmpHi, sort_index, 0
    If (tmpLow < inHi) Then Sort_Quick_A vArray, tmpLow, inHi, sort_index, 0
End Sub


Sub Sort_Merge(x As Variant)
Dim x1 As Variant, x2 As Variant
Dim i As Long, j As Long, n_raw As Long, n_mid As Long
    n_raw = UBound(x, 1)
    If n_raw = 1 Then Exit Sub
    n_mid = Int(n_raw / 2)
    ReDim x1(1 To n_mid)
    ReDim x2(1 To n_raw - n_mid)
    For i = 1 To n_mid
        x1(i) = x(i)
    Next i
    For i = 1 To (n_raw - n_mid)
        x2(i) = x(n_mid + i)
    Next i
    Call Sort_Merge(x1)
    Call Sort_Merge(x2)
    Call Merge(x1, x2, x)
End Sub


Private Sub Merge(xA As Variant, xB As Variant, xc As Variant)
Dim i As Long, j As Long, k As Long
Dim nA As Long, nB As Long, nc As Long
Dim sA As Long, Sb As Long
    nA = UBound(xA)
    nB = UBound(xB)
    nc = nA + nB
    ReDim xc(1 To nc)
    sA = 1
    Sb = 1
    Do While sA <= nA And Sb <= nB
        If xA(sA) > xB(Sb) Then
            k = k + 1
            xc(k) = xB(Sb)
            Sb = Sb + 1
        Else
            k = k + 1
            xc(k) = xA(sA)
            sA = sA + 1
        End If
    Loop
    Do While sA <= nA
        k = k + 1
        xc(k) = xA(sA)
        sA = sA + 1
    Loop
    Do While Sb <= nB
        k = k + 1
        xc(k) = xB(Sb)
        Sb = Sb + 1
    Loop
End Sub


'Heap sort as implemented in "Numerical Recipes in FORTAN77"
Sub Sort_Heap(x As Variant)
Dim i As Long, j As Long, k As Long, n As Long, iR As Long
Dim tmp_x As Variant
    n = UBound(x, 1)
    If n < 2 Then Exit Sub
    If n Mod 2 = 0 Then
        k = n / 2 + 1
    Else
        k = (n - 1) / 2 + 1
    End If
    iR = n
    Do
        If k > 1 Then
            k = k - 1
            tmp_x = x(k)
        Else
            tmp_x = x(iR)
            x(iR) = x(1)
            iR = iR - 1
            If iR = 1 Then
                x(k) = tmp_x
                Exit Sub
            End If
        End If
        i = k
        j = k + k
        Do While j <= iR
            If j < iR Then
                If x(j) < x(j + 1) Then j = j + 1
            End If
            If tmp_x < x(j) Then
                x(i) = x(j)
                i = j
                j = j + j
            Else
                j = iR + 1
            End If
        Loop
        x(i) = tmp_x
    Loop
End Sub


Sub Sort_Heap_A(x As Variant, sort_idx() As Long, Optional init_sort_idx As Long = 1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iR As Long
Dim tmp_x As Variant
    n = UBound(x, 1)
    If init_sort_idx = 1 Then
        ReDim sort_idx(1 To n)
        For i = 1 To n
            sort_idx(i) = i
        Next i
    End If
    If n < 2 Then Exit Sub
    If n Mod 2 = 0 Then
        k = n / 2 + 1
    Else
        k = (n - 1) / 2 + 1
    End If
    iR = n
    Do
        If k > 1 Then
            k = k - 1
            tmp_x = x(k)
            m = k
        Else
            tmp_x = x(iR)
            x(iR) = x(1)
            m = sort_idx(iR)
            sort_idx(iR) = sort_idx(1)
            iR = iR - 1
            If iR = 1 Then
                x(k) = tmp_x
                sort_idx(k) = m
                Exit Sub
            End If
        End If
        i = k
        j = k + k
        Do While j <= iR
            If j < iR Then
                If x(j) < x(j + 1) Then j = j + 1
            End If
            If tmp_x < x(j) Then
                x(i) = x(j)
                sort_idx(i) = sort_idx(j)
                i = j
                j = j + j
            Else
                j = iR + 1
            End If
        Loop
        x(i) = tmp_x
        sort_idx(i) = m
    Loop
End Sub


'========================================
'Percentile Functions
'========================================

Function fmedian(ByVal x As Variant) As Double
    Dim n As Long, k As Long
    n = UBound(x, 1)
    Call Sort_Quick(x, 1, n)
    If n Mod 2 = 0 Then
        fmedian = (x(n / 2) + x(n / 2 + 1)) / 2
    Else
        fmedian = x((n + 1) / 2)
    End If
End Function


'Intput: vector x(1 to N)
'Output: vector PercentileScore(1 to N)
Function PercentileScore(ByVal x As Variant) As Double()
Dim i As Long, j As Long, k As Long, n As Long
Dim iArr() As Long
Dim p() As Double
    n = UBound(x)
    Call Sort_Quick_A(x, 1, n, iArr, 1)
    ReDim p(1 To n)
    p(iArr(1)) = 0
    k = 0
    For i = 2 To n
        j = iArr(i)
        If x(i) > x(i - 1) Then k = i - 1
        p(j) = k * 1# / (n - 1)
    Next i
    PercentileScore = p
    Erase p, iArr
End Function


'Intput: vector x(1 to N)
'Output: vector fQuartile(0 to 4)
Function fQuartile(ByVal x As Variant) As Double()
Dim i As Long, j As Long, k As Long, n As Long
Dim tmp_y As Double
Dim iArr() As Long
Dim p() As Double
    n = UBound(x)
    Call Sort_Quick_A(x, 1, n, iArr, 1)
    ReDim p(0 To 4)
    p(0) = x(1)
    p(4) = x(n)
    For k = 1 To 3
        tmp_y = (n - 1) * k * 0.25
        i = Int(tmp_y)
        p(k) = x(i + 1) + (x(i + 2) - x(i + 1)) * (tmp_y - i)
    Next k
    fQuartile = p
    Erase p, iArr
End Function



'========================================
' Array manipulation
'========================================

'Return subset of an array A()
'Input: A(), 1D or 2D array
'       idx1, index positions of the first dimension to return
'       idx2, index positions of the second dimension to return
'Output: B()
Sub Filter_Array(ByVal A As Variant, B As Variant, Optional idx1 As Variant, Optional idx2 As Variant)
Dim i As Long, j As Long, k As Long
    k = getDimension(A)
    If k = 1 Then
        ReDim B(1 To UBound(idx1))
        For i = 1 To UBound(idx1)
            B(i) = A(idx1(i))
        Next i
    ElseIf k = 2 Then
        If IsMissing(idx1) = False And IsMissing(idx2) = False Then
            If IsArray(idx1) = True And IsArray(idx2) = True Then
                ReDim B(1 To UBound(idx1), 1 To UBound(idx2))
                For i = 1 To UBound(idx1)
                    For j = 1 To UBound(idx2)
                        B(i, j) = A(idx1(i), idx2(j))
                    Next j
                Next i
            ElseIf IsArray(idx1) = True And IsArray(idx2) = False Then
                ReDim B(1 To UBound(idx1))
                For i = 1 To UBound(idx1)
                    B(i) = A(idx1(i), idx2)
                Next i
            ElseIf IsArray(idx1) = False And IsArray(idx2) = True Then
                ReDim B(1 To UBound(idx2))
                For i = 1 To UBound(idx2)
                    B(i) = A(idx1, idx2(i))
                Next i
            End If
        ElseIf IsMissing(idx1) = False And IsMissing(idx2) = True Then
            If IsArray(idx1) = True Then
                ReDim B(1 To UBound(idx1), 1 To UBound(A, 2))
                For i = 1 To UBound(idx1)
                    For j = 1 To UBound(A, 2)
                        B(i, j) = A(idx1(i), j)
                    Next j
                Next i
            Else
                ReDim B(1 To UBound(A, 2))
                For i = 1 To UBound(A, 2)
                    B(i) = A(idx1, i)
                Next i
            End If
        ElseIf IsMissing(idx1) = True And IsMissing(idx2) = False Then
            If IsArray(idx2) = True Then
                ReDim B(1 To UBound(A, 1), 1 To UBound(idx2))
                For i = 1 To UBound(A, 1)
                    For j = 1 To UBound(idx2)
                        B(i, j) = A(i, idx2(j))
                    Next j
                Next i
            Else
                ReDim B(1 To UBound(A, 1))
                For i = 1 To UBound(A, 1)
                    B(i) = A(i, idx2)
                Next i
            End If
        End If
    End If
End Sub


'return y(1:n-m+1) with the elements of x(m:n)
Sub MidArray(x As Variant, m As Long, n As Long, y As Variant)
Dim i As Long
    ReDim y(1 To n - m + 1)
    For i = m To n
        y(i - m + 1) = x(i)
    Next i
End Sub


'Append tgt to the end of a 1D vector x(0 to n)
Sub Append_1D(x As Variant, tgt As Variant)
Dim n As Long
    n = UBound(x) + 1
    ReDim Preserve x(LBound(x) To n)
    x(n) = tgt
End Sub


'Remove the i-th element from vector x(0 to n)
Sub Erase_1D(x As Variant, i As Long)
Dim j As Long, n As Long
    n = UBound(x)
    If i = 0 Then Debug.Print "Erase_1D: error: cannot remove 0-th element"
    If i = n Then
        ReDim Preserve x(LBound(x) To n - 1)
    ElseIf i < n Then
        For j = i To n - 1
            x(j) = x(j + 1)
        Next j
        ReDim Preserve x(LBound(x) To n - 1)
    End If
End Sub


'Input: x(1 to M, 1 to N), 2D Matrix
'Output: y(), vector from the k-th row/column of x()
Sub get_vector(x As Variant, k As Long, idim As Long, y As Variant)
Dim i As Long, n As Long
If idim = 1 Then
    n = UBound(x, 2)
    ReDim y(1 To n)
    For i = 1 To n
        y(i) = x(k, i)
    Next i
ElseIf idim = 2 Then
    n = UBound(x, 1)
    ReDim y(1 To n)
    For i = 1 To n
        y(i) = x(i, k)
    Next i
End If
End Sub


'Generate an integer array from m to n
Function index_array(m As Long, n As Long) As Long()
    Dim i As Long
    Dim intArray() As Long
    ReDim intArray(m To n)
    For i = m To n
        intArray(i) = i
    Next i
    index_array = intArray
End Function


'Transpose when printing array to Excel worksheet
Function wkshtTranspose(A As Variant) As Variant
    wkshtTranspose = Application.WorksheetFunction.Transpose(A)
End Function


'For the k-th step during a k-fold cross validation, return the index of training set and validation set
Sub CrossValidate_set(k As Long, k_fold As Long, iList() As Long, i_validate() As Long, i_train() As Long)
Dim i As Long, j As Long, m As Long, n As Long, n_train As Long, n_validate As Long
    n = UBound(iList, 1)
    n_validate = n \ k_fold
    n_train = n - n_validate
    ReDim i_validate(1 To n_validate)
    ReDim i_train(1 To n_train)
    
    'Validation set
    For i = 1 To n_validate
        i_validate(i) = iList((k - 1) * n_validate + i)
    Next i
    
    'Training set
    j = 0
    If k > 1 Then
        For i = 1 To (k - 1) * n_validate
            j = j + 1
            i_train(j) = iList(i)
        Next i
    End If
    For i = k * n_validate + 1 To n
        j = j + 1
        i_train(j) = iList(i)
    Next i

    'If there are unused data at the last step add them to the validation set
    If k = k_fold And (k * n_validate) < n Then
        m = n - (k * n_validate)
        ReDim Preserve i_validate(1 To n_validate + m)
        ReDim Preserve i_train(1 To n_train - m)
        For i = 1 To m
            i_validate(n_validate + i) = iList(k * n_validate + i)
        Next i
    End If
End Sub



'========================================
' Random Numbers
'========================================

'A Random integer between lower to upper, including end points
Function Random_Integer(lower As Long, upper As Long) As Long
    Random_Integer = Int(Rnd() * (upper - lower + 1)) + lower
End Function

'Pick a random integer between 1 to N with given probablity distribution
'Input: Prob(), real vector of size (1:N) holding the probability of each integer
'       isCumulative, set to TRUE if Prob() is already giving the cumulative prob.
Function Random_Integer_Prob(Prob() As Double, Optional isCumulative As Boolean = False) As Long
Dim i As Long, n_raw As Long
Dim tmp_x As Double
Dim prob_C() As Double
    n_raw = UBound(Prob)
    If isCumulative = False Then
        ReDim prob_C(1 To n_raw)
        prob_C(1) = Prob(1)
        For i = 2 To n_raw
            prob_C(i) = prob_C(i - 1) + Prob(i)
        Next i
    Else
        prob_C = Prob
    End If
    
    Randomize
    tmp_x = Rnd()
    If tmp_x <= prob_C(1) Then
        Random_Integer_Prob = 1
    Else
        For i = 2 To n_raw
            If tmp_x > prob_C(i - 1) And tmp_x <= prob_C(i) Then
                Random_Integer_Prob = i
                Exit For
            End If
        Next i
    End If
    Erase prob_C
End Function

'Randomly shuffle elements of x()
Sub Shuffle(x As Variant)
Dim i As Long, j As Long, n As Long
Dim k As Long
Dim vtmp As Variant
    n = UBound(x)
    Randomize
    For i = n To 2 Step -1
        j = Int(Rnd() * i) + 1  'Random_Integer(1, i)
        vtmp = x(j)
        x(j) = x(i)
        x(i) = vtmp
    Next i
End Sub

'Radomly pick k-elements from vector x(1 to n) and output as y(1 to k)
Sub Random_Pick(x As Variant, k As Long, y As Variant)
Dim i As Long, j As Long, n As Long
    n = UBound(x)
    ReDim y(1 To k)
    For i = 1 To k
        y(i) = x(i)
    Next i
    Randomize
    For i = k + 1 To n
        j = Random_Integer(1, i)
        If j <= k Then y(j) = x(i)
    Next i
End Sub

'Heap's algorithm
'Generate all possible permuations of N objects
'Input:     n, number of elements
'           A(1:n), vector of n elements
'Output:    pList(1:n, 1:number of permutations), matrix where each row is a permuation of A
Sub Permute(n As Long, A As Variant, pList As Variant, Optional first_run As Long = 0)
Dim i As Long, j As Long, swap As Variant

    DoEvents
    If first_run = 0 Then
        ReDim pList(1 To n, 0 To 0)
        first_run = 1
    End If
    
    If n = 1 Then
        
        If UBound(pList, 2) = 0 Then
            j = 1
            ReDim Preserve pList(1 To UBound(pList, 1), 1 To 1)
        Else
            j = UBound(pList, 2) + 1
            ReDim Preserve pList(1 To UBound(pList, 1), 1 To j)
        End If
        For i = 1 To UBound(pList, 1)
            pList(i, j) = A(i)
        Next i
        Exit Sub
        
    Else
    
        For i = 1 To n - 1
            Call Permute(n - 1, A, pList, first_run)
            If n Mod 2 = 0 Then
                swap = A(i)
                A(i) = A(n)
                A(n) = swap
            Else
                swap = A(1)
                A(1) = A(n)
                A(n) = swap
            End If
        Next i
        Call Permute(n - 1, A, pList, first_run)
    
    End If

End Sub

Function nCk(n As Long, k As Long) As Double
Dim i As Long
    nCk = 1
    For i = 1 To k
        nCk = (n + 1 - i) * nCk / i
    Next i
End Function

'Return all combinations k elements from n, i.e. nCk
Function Combinations(n As Long, k As Long, Output() As Long) As Long
Dim i As Long
Dim v() As Long
    ReDim v(1 To k)
    ReDim Output(1 To k, 1 To 1)
    i = 0
    Call Combinations_Recursive(v, 1, n, 1, k, Output, i)
    Combinations = i
    Erase v
End Function

Private Sub Combinations_Recursive(v() As Long, start As Long, n As Long, k As Long, maxk As Long, _
        Output() As Long, iterate As Long)
Dim i As Long
    If k > maxk Then
        iterate = iterate + 1
        ReDim Preserve Output(1 To maxk, 1 To iterate)
        For i = 1 To maxk
            Output(i, iterate) = v(i)
        Next i
        Exit Sub
    End If
    For i = start To n
        v(k) = i
        Call Combinations_Recursive(v, i + 1, n, k + 1, maxk, Output, iterate)
    Next i
End Sub


'Generate a single random number from gaussian distribution with mean x_mean and std dev x_sd
Function Gaussian_Rand(Optional x_mean As Double = 0, Optional x_sd As Double = 1) As Double
    Randomize
    Gaussian_Rand = Sqr(-2 * Log(Rnd())) * Cos(6.28318530717959 * Rnd())
    Gaussian_Rand = Gaussian_Rand * x_sd + x_mean
End Function

'Generate N gaussian noise with given mean and std dev
Function Gaussian_Noise(x_mean As Double, x_sd As Double, n As Long, Optional force_mean_sd As Boolean = False) As Double()
Dim i As Long, m As Long
Dim x1 As Double, x2 As Double, w As Double
Dim y() As Double
ReDim y(1 To n)
Randomize
m = 0
Do
    w = 0
    Do
        x1 = 2 * Rnd() - 1
        x2 = 2 * Rnd() - 1
        w = x1 * x1 + x2 * x2
    Loop While w >= 1 Or w = 0
    w = Sqr(-2 * Log(w) / w)
    m = m + 1
    y(m) = x1 * w * x_sd + x_mean
    If m = n Then Exit Do
    m = m + 1
    y(m) = x2 * w * x_sd + x_mean
    If m = n Then Exit Do
Loop While m <= n

If force_mean_sd = True Then
    x1 = 0
    x2 = 0
    For i = 1 To n
        x1 = x1 + y(i)
        x2 = x2 + y(i) ^ 2
    Next i
    x1 = x1 / n
    x2 = Sqr((x2 / n - x1 * x1) * n * 1# / (n - 1))
    For i = 1 To n
        y(i) = ((y(i) - x1) / x2) * x_sd + x_mean
    Next i
End If
Gaussian_Noise = y
End Function


Function Gaussian_Noise_2D(mean1 As Double, mean2 As Double, sd1 As Double, sd2 As Double, correl As Double, n As Long) As Double()
Dim i As Long
Dim z1() As Double, z2() As Double, x() As Double
    z1 = Gaussian_Noise(0, 1, n)
    z2 = Gaussian_Noise(0, 1, n)
    ReDim x(1 To n, 1 To 2)
    For i = 1 To n
        x(i, 1) = mean1 + sd1 * z1(i)
        x(i, 2) = mean2 + sd2 * (z1(i) * correl + z2(i) * Sqr(1 - correl ^ 2))
    Next i
    Gaussian_Noise_2D = x
End Function


'=== https://arxiv.org/pdf/1505.03512.pdf
Function Gaussian_Noise_MV(x_mean() As Double, x_covar() As Double, n As Long)
Dim i As Long, j As Long, k As Long, m As Long, n_dimension As Long
Dim tmp_x As Double
Dim z() As Double, x() As Double, c() As Double
Dim x2() As Double, x_shift() As Double, x_scale() As Double

n_dimension = UBound(x_mean, 1)

ReDim z(1 To n, 1 To n_dimension)
For k = 1 To n_dimension
    c = Gaussian_Noise(0, 1, n)
    For i = 1 To n
        z(i, k) = c(i)
    Next i
Next k

c = Cholesky(x_covar)

ReDim x(1 To n, 1 To n_dimension)
For i = 1 To n
    For j = 1 To n_dimension
        x(i, j) = x_mean(j)
        For k = 1 To n_dimension
            x(i, j) = x(i, j) + c(j, k) * z(i, k)
        Next k
    Next j
Next i
Erase c, z

Call Normalize_x(x, x_shift, x_scale)
For j = 1 To n_dimension
    tmp_x = Sqr(x_covar(j, j))
    For i = 1 To n
        x(i, j) = x(i, j) * tmp_x + x_mean(j)
    Next i
Next j
Erase x_shift, x_scale

Gaussian_Noise_MV = x
End Function


'===========================================================
'Random number generator with continuous distribution
'===========================================================

'Normal Distribution:
Function Rand_Gaussian(n As Long, Optional x_mean As Double = 0, Optional x_sd As Double = 1) As Double()
    Rand_Gaussian = Gaussian_Noise(x_mean, x_sd, n)
End Function

'Laplace Distribution: characterized by median and average absolute deviation
Function Rand_Laplace(n As Long, Optional x_median As Double = 0, Optional x_aad As Double = 1) As Double()
Dim i As Long, u As Double, x() As Double
    Randomize
    ReDim x(1 To n)
    For i = 1 To n
        u = -0.5 + Rnd()
        x(i) = x_median - x_aad * Sgn(u) * Log(1 - 2 * Abs(u))
    Next i
    Rand_Laplace = x
    Erase x
End Function

'Asymetric Laplace Distribution:
Function Rand_ALD(n As Long, x_loc As Double, x_scale As Double, x_shape As Double) As Double()
Dim i As Long
Dim tmp_x As Double
Dim x() As Double
    Randomize
    ReDim x(1 To n)
    For i = 1 To n
        tmp_x = -x_shape + Rnd() * (1 / x_shape + x_shape)
        If tmp_x > 0 Then
            x(i) = x_loc - Log(tmp_x * x_shape) / (x_scale * x_shape)
        Else
            x(i) = x_loc + Log(-tmp_x / x_shape) / (x_scale / x_shape)
        End If
    Next i
    Rand_ALD = x
    Erase x
End Function

'Gamma distribution: X ~ Gamma(x_shape, x_scale)
Function Rand_Gamma(n As Long, x_shape As Double, x_scale As Double) As Double()
Dim i As Long, j As Long, k As Long, shape_int As Long
Dim tmp_x As Double, u As Double, v As Double, w As Double, e1 As Double
Dim delta As Double, xi As Double, eta As Double
Dim x() As Double
    shape_int = Int(x_shape)
    delta = x_shape - shape_int
    ReDim x(1 To n)
    For i = 1 To n
        tmp_x = 0
        For k = 1 To shape_int
            tmp_x = tmp_x - Log(Rnd())
        Next k
        x(i) = tmp_x
    Next i
    
    If delta > 0 Then
        e1 = Exp(1)
        For i = 1 To n
            Do
                u = Rnd()
                v = Rnd()
                w = Rnd()
                If u <= (e1 / (e1 + delta)) Then
                    xi = v ^ (1 / delta)
                    eta = w * xi ^ (delta - 1)
                Else
                    xi = 1 - Log(v)
                    eta = w * Exp(-xi)
                End If
                If eta <= (xi ^ (delta - 1) * Exp(-xi)) Then Exit Do
            Loop
            x(i) = (xi + x(i))
        Next i
    End If
    
    For i = 1 To n
        x(i) = x(i) * x_scale
    Next i
    
    Rand_Gamma = x
    Erase x
End Function


'========================================
' Special Functions
'========================================

'Returns ln|gamma(x)| for x>0
'Lanczos approximation from Numerical Recipe FORTRAN77 Chapter 6.1
'x can either be a single or a vector of real positive numbers
Function gammalns(x As Variant) As Variant
Dim i As Long, j As Long, m As Long, n As Long
Dim ser As Double, stp As Double, tmp As Double, z As Double
Dim cof() As Double, y() As Double
    ReDim cof(1 To 6)
    cof(1) = 76.1800917294715       '76.18009172947146d0
    cof(2) = -86.5053203294168      '-86.50532032941677d0
    cof(3) = 24.0140982408309       '24.01409824083091d0
    cof(4) = -1.23173957245015      '-1.231739572450155d0
    cof(5) = 1.20865097386618E-03   '.1208650973866179d-2
    cof(6) = -5.395239384953E-06    '-.5395239384953d-5
    stp = 2.506628274631            '2.5066282746310005d0
    If IsArray(x) = False Then
        tmp = x + 5.5
        tmp = (x + 0.5) * Log(tmp) - tmp
        ser = 1.00000000019001          '1.000000000190015d0
        For j = 1 To 6
            ser = ser + cof(j) / (x + j)
        Next j
        gammalns = tmp + Log(stp * ser / x)
    Else
        m = LBound(x, 1)
        n = UBound(x, 1)
        ReDim y(m To n)
        For i = m To n
            z = x(i)
            tmp = z + 5.5
            tmp = (z + 0.5) * Log(tmp) - tmp
            ser = 1.00000000019001      '1.000000000190015d0
            For j = 1 To 6
                ser = ser + cof(j) / (z + j)
            Next j
            y(i) = tmp + Log(stp * ser / z)
        Next i
        gammalns = y
        Erase y
    End If
    Erase cof
End Function

''pdf of student-t distribution
''nu = degree of freedom
'Function student_pdf(x As Double, nu As Double, Optional mu As Double = 0, Optional var As Double = 1) As Double
'    student_pdf = Exp(gammaln((nu + 1) * 0.5) - gammaln(nu * 0.5)) / Sqr(nu * 3.14159265358979 * var)
'    student_pdf = student_pdf * ((1 + ((x - mu) ^ 2) / (nu * var)) ^ (-(nu + 1) * 0.5))
'End Function
'
''beta function
'Function beta_func(x As Double, y As Double) As Double
'    beta_func = Exp(gammaln(x) + gammaln(y) - gammaln(x + y))
'End Function


'========================================
' Linear Algebra
'========================================

'Multivariate linear regression of y(1 to n_raw,) vs x(1 to n_raw,, 1 to n_dimension)
'Input: y(1 to n_raw)
'Input: x(1 to n_raw, 1 to n_dimension)
'Output: y_slope(1 to n_dimension+1), where the last term is the intercept
Sub linear_regression(y As Variant, x As Variant, y_slope() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim y_avg As Double, x_avg As Double
Dim tmp_x As Double
Dim x2() As Double, xA() As Double

n_raw = UBound(y)
n_dimension = UBound(x, 2)
n = n_dimension + 1

ReDim y_slope(1 To n)

x2 = x
ReDim Preserve x2(1 To n_raw, 1 To n)
For i = 1 To n_raw
    x2(i, n) = 1
Next i

ReDim xA(1 To n, 1 To n)
For i = 1 To n
    tmp_x = 0
    For k = 1 To n_raw
        tmp_x = tmp_x + x2(k, i) ^ 2
    Next k
    xA(i, i) = tmp_x
    For j = i + 1 To n
        tmp_x = 0
        For k = 1 To n_raw
            tmp_x = tmp_x + x2(k, i) * x2(k, j)
        Next k
        xA(i, j) = tmp_x
        xA(j, i) = tmp_x
    Next j
Next i

xA = Matrix_Inverse_Cholesky(xA)
y_slope = M_Dot(M_Dot(xA, x2, , 1), y)
Erase xA, x2
End Sub


'Multivariate linear regression of y(1 to N) vs x(1 to N, 1 to D)
'using the method of gradient descent
'Returns coeff(0 to D) where the 0-th element is the intercept
Sub Linear_Regression_Gradient_Descent(y() As Double, x() As Double, Coeff() As Double, _
    Optional learn_rate As Double = 0.0001, Optional iter_max As Long = 1000, _
    Optional tolerance As Double = 0.0000001, Optional conv_chk As Long = 10)
Dim i As Long, j As Long, k As Long, n As Long, iterate As Long, conv_count As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double, mse As Double, mse_prev As Double
Dim grad() As Double, gain() As Double, coeff_chg() As Double
Dim x2() As Double, y2() As Double, x_avg() As Double, x_sd() As Double, y_avg As Double, y_sd As Double

    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    'Normalize variables to zero mean and unit variance
    's.t. initial guess can be resonably assumed to be in range [-1,1]
    ReDim y2(1 To n)
    ReDim x2(1 To n, 1 To n_dimension)
    ReDim x_avg(1 To n_dimension)
    ReDim x_sd(1 To n_dimension)
    For k = 1 To n_dimension
        For i = 1 To n
            x_avg(k) = x_avg(k) + x(i, k)
            x_sd(k) = x_sd(k) + x(i, k) ^ 2
        Next i
        x_avg(k) = x_avg(k) / n
        x_sd(k) = Sqr((x_sd(k) - (x_avg(k) ^ 2) * n) / (n - 1))
        For i = 1 To n
            x2(i, k) = (x(i, k) - x_avg(k)) / x_sd(k)
        Next i
    Next k
    For i = 1 To n
        y_avg = y_avg + y(i)
        y_sd = y_sd + y(i) ^ 2
    Next i
    y_avg = y_avg / n
    y_sd = Sqr((y_sd - (y_avg ^ 2) * n) / (n - 1))
    For i = 1 To n
        y2(i) = (y(i) - y_avg) / y_sd
    Next i
    
    'Random guess of regression coefficients
    Randomize
    ReDim Coeff(0 To n_dimension)
    For k = 1 To n_dimension
        Coeff(k) = -1 + 2 * Rnd()
    Next k
    
    'Begin gradient descent (adaptive)
    mse_prev = Exp(70)
    ReDim gain(0 To n_dimension)
    ReDim coeff_chg(0 To n_dimension)
    For k = 0 To n_dimension
        gain(k) = 1
    Next k
    For iterate = 1 To iter_max
        mse = 0
        ReDim grad(0 To n_dimension)
        For i = 1 To n
            tmp_x = Coeff(0)
            For k = 1 To n_dimension
                tmp_x = tmp_x + x2(i, k) * Coeff(k)
            Next k
            tmp_x = y2(i) - tmp_x
            mse = mse + tmp_x ^ 2
            grad(0) = grad(0) - tmp_x
            For k = 1 To n_dimension
                grad(k) = grad(k) - tmp_x * x2(i, k)
            Next k
        Next i
        mse = mse / n
        
        'Exit if cost function is better than tolerance
        'or if there's no significat improvement to cost function
        If mse < tolerance Then
            Exit For
        ElseIf (mse_prev - mse) < tolerance Then
            conv_count = conv_count + 1
            If conv_count = conv_chk Then Exit For
        Else
            conv_count = 0
        End If
        mse_prev = mse
        
        'Accelerate learning process if current and previous
        'gradient are in the same direction
        For k = 0 To n_dimension
            If Sgn(grad(k)) = Sgn(coeff_chg(k)) Then
                gain(k) = gain(k) * 0.8
            Else
                gain(k) = gain(k) + 0.2
            End If
            If gain(k) < 0.1 Then gain(k) = 0.1
        Next k
        
        For k = 0 To n_dimension
            coeff_chg(k) = -learn_rate * gain(k) * grad(k)
            Coeff(k) = Coeff(k) + coeff_chg(k)
        Next k
    Next iterate
    Erase x2, y2, grad, gain, coeff_chg
    
    'recover normalization costants
    For k = 1 To n_dimension
        Coeff(0) = Coeff(0) - Coeff(k) * x_avg(k) / x_sd(k)
        Coeff(k) = Coeff(k) * y_sd / x_sd(k)
    Next k
    Coeff(0) = Coeff(0) * y_sd + y_avg
    Erase x2, y2, x_avg, x_sd, grad, gain, coeff_chg
End Sub



'Linear Regression with of y() vs x() with an additional categorical variable class()
'Constraint is applied such that number weighted sum of cofficients of all classes is zero.
Sub Linear_Regression_Dummy(y() As Double, x() As Double, Class As Variant, y_slope() As Double, class_list As Variant, class_slope() As Double, Optional y_predicted As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long, n_class As Long
Dim tmp_x As Double
Dim class_index() As Long, class_size() As Long
Dim xArr() As Double
    msgbox "ABC"
    n_raw = UBound(y, 1)
    n_dimension = UBound(x, 2)
    
    Call Unique_Items(Class, class_index, class_list, n_class, class_size)

    xArr = x
    ReDim Preserve xArr(1 To n_raw, 1 To n_dimension + n_class - 1)
    For i = 1 To n_raw
        k = class_index(i)
        If k = n_class Then
            For j = 1 To n_class - 1
                xArr(i, n_dimension + j) = -1 * class_size(j) / class_size(n_class)
            Next j
        Else
            xArr(i, n_dimension + k) = 1
        End If
    Next i
    
    Call linear_regression(y, xArr, y_slope)
    
    If IsEmpty(y_predicted) = False Then
        ReDim y_predicted(1 To n_raw)
        For i = 1 To n_raw
            tmp_x = y_slope(n_dimension + n_class)
            For j = 1 To n_dimension + n_class - 1
                tmp_x = tmp_x + xArr(i, j) * y_slope(j)
            Next j
            y_predicted(i) = tmp_x
        Next i
    End If
    
    xArr = y_slope
    ReDim y_slope(1 To n_dimension + 1)
    For i = 1 To n_dimension
        y_slope(i) = xArr(i)
    Next i
    y_slope(n_dimension + 1) = xArr(UBound(xArr, 1))
    
    ReDim class_slope(1 To n_class)
    For i = 1 To n_class - 1
        class_slope(i) = xArr(n_dimension + i)
        class_slope(n_class) = class_slope(n_class) - xArr(n_dimension + i) * class_size(i) / class_size(n_class)
    Next i
    
    Erase xArr, class_size
End Sub

'Input: y(1 to n_raw)
'Input: x(1 to n_raw)
'Output: y_slope(1 to 2), first term is the slope, 2nd term is the intercept
Sub linear_regression_single(y As Variant, x As Variant, y_slope() As Double)
Dim i As Long, m As Long, n As Long, n_raw As Long
Dim y_avg As Double, x_avg As Double
Dim tmp_x As Double
    ReDim y_slope(1 To 2)
    n_raw = UBound(y)
    y_avg = 0
    x_avg = 0
    For i = 1 To n_raw
        y_avg = y_avg + y(i)
        x_avg = x_avg + x(i)
    Next i
    y_avg = y_avg / n_raw
    x_avg = x_avg / n_raw
    tmp_x = 0
    For i = 1 To n_raw
        y_slope(1) = y_slope(1) + (x(i) - x_avg) * (y(i) - y_avg)
        tmp_x = tmp_x + (x(i) - x_avg) ^ 2
    Next i
    y_slope(1) = y_slope(1) / tmp_x
    y_slope(2) = y_avg - y_slope(1) * x_avg
End Sub

'1-D polynomial curve fitting of y versus x up to a specified order
'Input: y(1 to n), dependent variable
'Input: x(1 to n), independent variable
'Output: coeff(0 to fit_order)
Sub Polynomial_Fit(y As Variant, x As Variant, Coeff() As Double, Optional fit_order As Long = 2)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_raw As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, x_mean As Double, x_sd As Double
Dim x_matrix() As Double, A() As Double, B() As Double, coeff_new() As Double
    n_raw = UBound(y, 1)

    'Normalize x for better numerical stability
    x_mean = 0
    x_sd = 0
    For i = 1 To n_raw
        x_mean = x_mean + x(i)
        x_sd = x_sd + x(i) ^ 2
    Next i
    x_mean = x_mean / n_raw
    x_sd = Sqr((x_sd - (x_mean ^ 2) * n_raw) / (n_raw - 1))

    ReDim x_matrix(1 To n_raw, 0 To fit_order)
    ReDim A(1 To fit_order + 1, 1 To fit_order + 1)
    ReDim B(1 To 2 * fit_order)
    For i = 1 To n_raw
        x_matrix(i, 0) = 1
        tmp_x = (x(i) - x_mean) / x_sd
        tmp_z = tmp_x ^ fit_order
        For k = 1 To fit_order
            tmp_y = tmp_x ^ k
            x_matrix(i, k) = tmp_y
            B(k) = B(k) + tmp_y
            B(fit_order + k) = B(fit_order + k) + tmp_y * tmp_z
        Next k
    Next i

    A(1, 1) = n_raw
    For k = 1 To fit_order
        A(k + 1, k + 1) = B(2 * k)
    Next k
    For i = 0 To fit_order - 1
        For j = i + 1 To fit_order
            A(i + 1, j + 1) = B(i + j)
            A(j + 1, i + 1) = B(i + j)
        Next j
    Next i
    A = Matrix_Inverse_Cholesky(A)

    ReDim B(0 To fit_order, 1 To n_raw)
    For k = 1 To n_raw
        For i = 0 To fit_order
            tmp_x = 0
            For j = 0 To fit_order
                tmp_x = tmp_x + A(i + 1, j + 1) * x_matrix(k, j)
            Next j
            B(i, k) = tmp_x
        Next i
    Next k
    
    Erase x_matrix, A
    
    ReDim Coeff(0 To fit_order)
    For k = 0 To fit_order
        tmp_x = 0
        For i = 1 To n_raw
            tmp_x = tmp_x + B(k, i) * y(i)
        Next i
        Coeff(k) = tmp_x
    Next k
    
    For k = 1 To fit_order
        Coeff(k) = Coeff(k) / (x_sd ^ k)
    Next k
    
    ReDim coeff_new(0 To fit_order)
    For k = 0 To fit_order
        B = Polynomial_Coeff(k, -x_mean)
        For i = 0 To k
            coeff_new(i) = coeff_new(i) + Coeff(k) * B(i)
        Next i
    Next k
    
    Coeff = coeff_new
    Erase B, coeff_new
End Sub


'Find the coefficients of x^p in the expansion of (x+s)^p
Function Polynomial_Coeff(p As Long, s As Double) As Double()
Dim i As Long, j As Long, k As Long
Dim Coeff() As Double, Coeff_A() As Double

    If p = 0 Then
        ReDim Coeff(0 To 0)
        Coeff(0) = 1
        Polynomial_Coeff = Coeff
        Exit Function
    End If
    
    ReDim Coeff(0 To 1)
    Coeff(0) = s
    Coeff(1) = 1
    For k = 1 To (p - 1)
        ReDim Coeff_A(0 To (k + 1))
        Coeff_A(0) = s * Coeff(0)
        Coeff_A(k + 1) = Coeff(k)
        For i = 1 To k
            Coeff_A(i) = Coeff(i - 1) + s * Coeff(i)
        Next i
        Coeff = Coeff_A
    Next k
    Polynomial_Coeff = Coeff
End Function


'Returns a curve that best fits y as a polynomial function of x
Function Polynomial_Curve(y() As Double, x() As Double, Optional fit_order As Long = 2) As Double()
Dim i As Long, k As Long
Dim tmp_x As Double
Dim Coeff() As Double, y_fit() As Double
    Call Polynomial_Fit(y, x, Coeff, fit_order)
    ReDim y_fit(1 To UBound(y))
    For i = 1 To UBound(y)
        tmp_x = Coeff(0)
        For k = 1 To fit_order
            tmp_x = tmp_x + Coeff(k) * (x(i) ^ k)
        Next k
        y_fit(i) = tmp_x
    Next i
    Polynomial_Curve = y_fit
    Erase y_fit
End Function


'=== Ridge Regression
'L = ||y - b x||^2+ (lambda/N)||b||^2
Function RidgeReg(ByVal y As Variant, ByVal x As Variant, Optional lambda As Double = 0.1) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim x_mean() As Double, x_sd() As Double
Dim y_mean As Double, y_sd As Double
Dim u() As Double, s() As Double, v() As Double
Dim beta() As Double
    n = UBound(y, 1)
    If getDimension(x) = 1 Then
        n_dimension = 1
        Call Promote_Vec(x, x)
    Else
        n_dimension = UBound(x, 2)
    End If
    
    'scale to zero mean and variance
    Call Normalize_x(x, x_mean, x_sd, "AVGSD")
    y_mean = 0
    y_sd = 0
    For i = 1 To n
        y_mean = y_mean + y(i)
        y_sd = y_sd + y(i) ^ 2
    Next i
    y_mean = y_mean / n
    y_sd = Sqr((y_sd - n * (y_mean ^ 2)) / (n - 1))
    For i = 1 To n
        y(i) = (y(i) - y_mean) / y_sd
    Next i
    
    'X=UDV'
    'beta = V (D/(D^2+lambda)) U' y
    Call Matrix_SVD(x, u, s, v)
    For i = 1 To UBound(s)
        s(i) = s(i) / (s(i) ^ 2 + n * lambda)
    Next i
    s = mDiag(s)
    beta = M_Dot(modMath.M_Dot(v, s), modMath.M_Dot(u, y, 1))
    
    'Restore to original scale
    ReDim Preserve beta(1 To n_dimension + 1)
    beta(n_dimension + 1) = y_mean
    For i = 1 To n_dimension
        beta(i) = beta(i) * y_sd / x_sd(i)
        beta(n_dimension + 1) = beta(n_dimension + 1) - beta(i) * x_mean(i)
    Next i

    RidgeReg = beta
    Erase beta, u, s, v, x_mean, x_sd
End Function


'Solve system of linear equations with Gaussian Eliminations
'Input: A(1 to N, 1 to N), square matrix with coefficients to independent variables
'Input: y(1 to N), vector of dependent variables
Function Solve_Linear_Equations(A As Variant, y As Variant) As Double()
Dim i As Long, n As Long
Dim A_Augmented() As Double

    n = UBound(A, 1)
    
    If UBound(A, 2) <> n Then
        Debug.Print "Solve_Linear_Equations: A is not a square matrix!"
        Exit Function
    ElseIf UBound(y, 1) <> n Then
        Debug.Print "Solve_Linear_Equations: dimension of A and y not consistent!"
        Exit Function
    End If
    
    A_Augmented = A
    ReDim Preserve A_Augmented(1 To n, 1 To n + 1)
    For i = 1 To n
        A_Augmented(i, n + 1) = y(i)
    Next i
    
    Call Gaussian_Elimination(A_Augmented)
    Solve_Linear_Equations = Back_Substitution(A_Augmented)
    Erase A_Augmented
End Function


Function Back_Substitution(A As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double
Dim x() As Double
    m = UBound(A, 1)
    n = UBound(A, 2) - 1
    ReDim x(1 To m)
    If A(m, n + 1) = 0 And A(m, m) = 0 Then
        x(m) = 1
    Else
        x(m) = A(m, n + 1) / A(m, m)
    End If
    For i = m - 1 To 1 Step -1
        tmp_x = 0
        For j = i + 1 To n
            tmp_x = tmp_x + A(i, j) * x(j)
        Next j
        x(i) = (A(i, n + 1) - tmp_x) / A(i, i)
    Next i
    Back_Substitution = x
End Function


Sub Gaussian_Elimination(A As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim i_max As Long
Dim tmp_x As Double

m = UBound(A, 1)
n = UBound(A, 2)

For k = 1 To m - 1
    tmp_x = -Exp(70)
    For i = k To m
        If Abs(A(i, k)) > tmp_x Then
            tmp_x = Abs(A(i, k))
            i_max = i
        End If
    Next i
    
    If A(i_max, k) = 0 Then
        Debug.Print "Gaussian_Elimation: Matrix is singular!"
        Exit Sub
    End If
    
    If i_max <> k Then
        For i = 1 To n
            tmp_x = A(k, i)
            A(k, i) = A(i_max, i)
            A(i_max, i) = tmp_x
        Next i
    End If
    
    For i = k + 1 To m
        tmp_x = A(i, k) / A(k, k)
        For j = k + 1 To n
            A(i, j) = A(i, j) - A(k, j) * tmp_x
        Next j
        A(i, k) = 0
    Next i

Next k

End Sub


'Solve system of linear equations Ax=B with Cholesky factorization
'Input: A(1 to N, 1 to N), square matrix with coefficients to independent variables
'Input: B(1 to N, 1 to M), vector of dependent variables
Function Solve_Linear_Cholesky(A As Variant, B As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, B_dim As Long
Dim tmp_x As Double, L() As Double, x() As Double, y() As Double

    n = UBound(A, 1)

    If UBound(A, 2) <> n Then
        Debug.Print "Solve_Linear_Cholesky: A is not a square matrix!"
        Exit Function
    ElseIf UBound(B, 1) <> n Then
        Debug.Print "Solve_Linear_Cholesky: dimension of A and y not consistent!"
        Exit Function
    End If
    
    L = Cholesky(A)
    
    B_dim = getDimension(B)
    If B_dim = 1 Then
        m = 1
        ReDim y(1 To n)
    Else
        m = UBound(B, 2)
        ReDim y(1 To n, 1 To m)
    End If
    
    For k = 1 To m
    
        ReDim x(1 To n)
        
        If B_dim = 1 Then
            For i = 1 To n
                x(i) = B(i)
            Next i
        Else
            For i = 1 To n
                x(i) = B(i, k)
            Next i
        End If
        
        'Forward Substitution
        For i = 1 To n
            For j = 1 To i - 1
                x(i) = x(i) - L(i, j) * x(j)
            Next j
            x(i) = x(i) / L(i, i)
        Next i
        
        'Backward Substitution
        For i = n To 1 Step -1
            tmp_x = 0
            For j = i + 1 To n
                tmp_x = tmp_x + L(j, i) * x(j)
            Next j
            x(i) = (x(i) - tmp_x) / L(i, i)
        Next i
        
        If B_dim = 1 Then
            y = x
        Else
            For i = 1 To n
                y(i, k) = x(i)
            Next i
        End If
        
    Next k
    Solve_Linear_Cholesky = y
    Erase L, x, y
End Function


'Solve system of linear equations Ax=B with LDL-factorization
'Input: A(1 to N, 1 to N), square matrix with coefficients to independent variables
'Input: B(1 to N, 1 to M), m vector of dependent variables
Function Solve_Linear_LDL(A As Variant, B As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim L() As Double, d() As Double, x() As Double, y() As Double
    n = UBound(A, 1)
    Call LDL_Decompose(A, L, d)
    
    If getDimension(B) = 1 Then
        ReDim x(1 To n)
        For i = 1 To n
            x(i) = B(i)
            For j = 1 To i - 1
                x(i) = x(i) - L(i, j) * x(j)
            Next j
        Next i
        For i = 1 To n
            x(i) = x(i) / d(i)
        Next i
        For i = n To 1 Step -1
            For j = i + 1 To n
                x(i) = x(i) - L(j, i) * x(j)
            Next j
        Next i
    Else
        m = UBound(B, 2)
        ReDim x(1 To n, 1 To m)
        For k = 1 To m
            ReDim y(1 To n)
            For i = 1 To n
                y(i) = B(i, k)
                For j = 1 To i - 1
                    y(i) = y(i) - L(i, j) * y(j)
                Next j
            Next i
            For i = 1 To n
                y(i) = y(i) / d(i)
            Next i
            For i = n To 1 Step -1
                For j = i + 1 To n
                    y(i) = y(i) - L(j, i) * y(j)
                Next j
                x(i, k) = y(i)
            Next i
        Next k
    End If
    
    Solve_Linear_LDL = x
    Erase x, y
End Function


'Solve system of linear equations Ax=B with LDL-factorization with pivoting
'Input: A(1 to N, 1 to N), square matrix with coefficients to independent variables
'Input: B(1 to N, 1 to M), m vector of dependent variables
Function Solve_Linear_LDLP(A As Variant, B As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, B_dim As Long, m_max As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim L() As Double, d() As Double, p() As Long, pivot() As Long
Dim x() As Double, y() As Double
    n = UBound(A, 1)
    Call LDLP_Decompose(A, L, d, p, pivot)
    B_dim = getDimension(B)
    If B_dim = 1 Then
        ReDim y(1 To n)
        m_max = 1
    Else
        ReDim y(1 To n, 1 To UBound(B, 2))
        m_max = UBound(B, 2)
    End If
    For m = 1 To m_max
    
        ReDim x(1 To n)
        
        If B_dim = 1 Then
            For i = 1 To n
                x(i) = B(p(i))
            Next i
        Else
            For i = 1 To n
                x(i) = B(p(i), m)
            Next i
        End If
        For i = 1 To n
            For j = 1 To i - 1
                x(i) = x(i) - L(i, j) * x(j)
            Next j
        Next i
        
        k = 1
        Do While k <= n
            If pivot(k) = 1 Then
                x(k) = x(k) / d(k, k)
                k = k + 1
            ElseIf pivot(k) = 2 Then
                tmp_x = d(k, k) * d(k + 1, k + 1) - d(k + 1, k) ^ 2
                tmp_y = (d(k + 1, k + 1) * x(k) - d(k + 1, k) * x(k + 1)) / tmp_x
                tmp_z = (-d(k + 1, k) * x(k) + d(k, k) * x(k + 1)) / tmp_x
                x(k) = tmp_y
                x(k + 1) = tmp_z
                k = k + 2
            End If
        Loop
        
        For i = n To 1 Step -1
            For j = i + 1 To n
                x(i) = x(i) - L(j, i) * x(j)
            Next j
        Next i
        
        
        If B_dim = 1 Then
            For i = 1 To n
                y(p(i)) = x(i)
            Next i
        Else
            For i = 1 To n
                y(p(i), m) = x(i)
            Next i
        End If
    
    Next m

Solve_Linear_LDLP = y
Erase x, y, L, d, p, pivot
End Function



'Solve system of linear equations Ax=B with LU-factorization
'Input: A(1 to N, 1 to N), square matrix with coefficients to independent variables
'Input: B(1 to N, 1 to m), vector of dependent variables
Function Solve_Linear_LU(A As Variant, B As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, B_dim As Long
Dim tmp_x As Double
Dim LU() As Double, p() As Long, x() As Double, y() As Double
    n = UBound(A, 1)
    If UBound(A, 2) <> n Then
        Debug.Print "Solve_Linear_LU: A is not a square matrix!"
        Exit Function
    ElseIf UBound(B, 1) <> n Then
        Debug.Print "Solve_Linear_LU: dimension of A and y not consistent!"
        Exit Function
    End If
    
    ReDim LU(1 To n, 1 To n)
    For i = 1 To n
        For j = i To n
            tmp_x = 0
            For k = 1 To i - 1
                tmp_x = tmp_x + LU(i, k) * LU(k, j)
            Next k
            LU(i, j) = A(i, j) - tmp_x
        Next j
        For j = i + 1 To n
            tmp_x = 0
            For k = 1 To i - 1
                tmp_x = tmp_x + LU(j, k) * LU(k, i)
            Next k
            LU(j, i) = (A(j, i) - tmp_x) / LU(i, i)
        Next j
    Next i
    
    B_dim = modMath.getDimension(B)
    If B_dim = 1 Then
        ReDim x(1 To n)
        ReDim y(1 To n)
        For i = 1 To n
            tmp_x = 0
            For k = 1 To i - 1
                tmp_x = tmp_x + LU(i, k) * y(k)
            Next k
            y(i) = B(i) - tmp_x
        Next i
        For i = n To 1 Step -1
            tmp_x = 0
            For k = i + 1 To n
                tmp_x = tmp_x + LU(i, k) * x(k)
            Next k
            x(i) = (y(i) - tmp_x) / LU(i, i)
        Next i
    ElseIf B_dim = 2 Then
        ReDim x(1 To n, 1 To UBound(B, 2))
        For m = 1 To UBound(B, 2)
            ReDim y(1 To n)
            For i = 1 To n
                tmp_x = 0
                For k = 1 To i - 1
                    tmp_x = tmp_x + LU(i, k) * y(k)
                Next k
                y(i) = B(i, m) - tmp_x
            Next i
            For i = n To 1 Step -1
                tmp_x = 0
                For k = i + 1 To n
                    tmp_x = tmp_x + LU(i, k) * x(k, m)
                Next k
                x(i, m) = (y(i) - tmp_x) / LU(i, i)
            Next i
        Next m
    End If
    Solve_Linear_LU = x
    Erase LU, x, y
End Function


'Solve system of linear equations Ax=B with LU-factorization with partial pivoting
'Input: A(1 to N, 1 to N), square matrix with coefficients to independent variables
'Input: B(1 to N, 1 to m), vector of dependent variables
Function Solve_Linear_LUP(A As Variant, B As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, B_dim As Long
Dim tmp_x As Double
Dim LU() As Double, p() As Long, x() As Double, y() As Double
    n = UBound(A, 1)
    If UBound(A, 2) <> n Then
        Debug.Print "Solve_Linear_LU: A is not a square matrix!"
        Exit Function
    ElseIf UBound(B, 1) <> n Then
        Debug.Print "Solve_Linear_LU: dimension of A and y not consistent!"
        Exit Function
    End If
    
    LU = A
    Call LUPDecompose(LU, p)
    
    B_dim = modMath.getDimension(B)
    If B_dim = 1 Then
        ReDim x(1 To n)
        For i = 1 To n
            x(i) = B(p(i))
            For j = 1 To i - 1
                x(i) = x(i) - LU(i, j) * x(j)
            Next j
        Next i
        For i = n To 1 Step -1
            For j = i + 1 To n
                x(i) = x(i) - LU(i, j) * x(j)
            Next j
            x(i) = x(i) / LU(i, i)
        Next i
    ElseIf B_dim = 2 Then
        ReDim x(1 To n, 1 To UBound(B, 2))
        For m = 1 To UBound(B, 2)
            ReDim y(1 To n)
            For i = 1 To n
                y(i) = B(p(i), m)
                For j = 1 To i - 1
                    y(i) = y(i) - LU(i, j) * y(j)
                Next j
            Next i
            For i = n To 1 Step -1
                For j = i + 1 To n
                    y(i) = y(i) - LU(i, j) * y(j)
                Next j
                y(i) = y(i) / LU(i, i)
                x(i, m) = y(i)
            Next i
        Next m
    End If
    Solve_Linear_LUP = x
    Erase LU, x, y, p
End Function


'Solve for x() s.t. Ax=y, where A() is a NxN tri-diagonal matrix, represented as A(1:N,1:3)
Function Solve_Tridiag(ByVal A As Variant, ByVal y As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim mm As Long, nn As Long
Dim x() As Double
    m = LBound(A, 1)
    n = UBound(A, 1)
    If UBound(A, 2) <> 3 Or A(m, 1) <> 0 Or A(n, 3) <> 0 Then
        Debug.Print "Solve_Tridiag: A() is not in tri-diagonal form. "
        Exit Function
    End If
    For i = m + 1 To n
        A(i, 2) = A(i, 2) * A(i - 1, 2) - A(i, 1) * A(i - 1, 3)
        A(i, 3) = A(i - 1, 2) * A(i, 3)
    Next i
    k = getDimension(y)
    If k = 1 Then
        For i = m + 1 To n
            y(i) = y(i) * A(i - 1, 2) - y(i - 1) * A(i, 1)
        Next i
        ReDim x(m To n)
        x(n) = y(n) / A(n, 2)
        For i = n - 1 To m Step -1
            x(i) = (y(i) - A(i, 3) * x(i + 1)) / A(i, 2)
        Next i
    ElseIf k = 2 Then
        mm = LBound(y, 2)
        nn = UBound(y, 2)
        ReDim x(m To n, mm To nn)
        For k = mm To nn
            For i = m + 1 To n
                y(i, k) = y(i, k) * A(i - 1, 2) - y(i - 1, k) * A(i, 1)
            Next i
            x(n, k) = y(n, k) / A(n, 2)
            For i = n - 1 To m Step -1
                x(i, k) = (y(i, k) - A(i, 3) * x(i + 1, k)) / A(i, 2)
            Next i
        Next k
    End If
    Solve_Tridiag = x
    Erase x, A, y
End Function



'======================
'Data preparation
'======================

'Promote vector x() into a matrix x2() by adding an empty dimension
Sub Promote_Vec(ByVal x As Variant, x2 As Variant)
Dim i As Long, m As Long, n As Long
    m = LBound(x, 1)
    n = UBound(x, 1)
    ReDim x2(m To n, 1 To 1)
    For i = m To n
        x2(i, 1) = x(i)
    Next i
End Sub


'Divide multi-dimensional data x(1:N,1:D) into segments of length n_T, stored in jagged array xS
'Input:  x(1:N, 1:D), N observations of of D-dimensional vector
'        n_T, length of each segment
'        t_start, starting point to create segments
'        step_size, interval between starting points of consecutive segment, nT=step_size means no overlap
'Output: xS(1:M), variant array, each element xS(i) is an array of size (1:n_T, 1:D)
Sub Sequence2Segments(x() As Double, xS As Variant, n_T As Long, Optional t_start As Long = 1, Optional step_size As Long = 1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim x_tmp() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    m = 0
    ReDim xS(1 To (n_raw - t_start - n_T + 1) / step_size + 1)
    For n = t_start To n_raw Step step_size
        If (n + n_T - 1) <= n_raw Then
            ReDim x_tmp(1 To n_T, 1 To n_dimension)
            For i = 1 To n_T
                For j = 1 To n_dimension
                    x_tmp(i, j) = x(n + i - 1, j)
                Next j
            Next i
            m = m + 1
            xS(m) = x_tmp
        End If
    Next n
    ReDim Preserve xS(1 To m)
    Erase x_tmp
End Sub


'Winsorize mutli-dimensional data x()
'data with percentile scores outside threshold_upper and threshold_lower are replaced
Sub Winsorize_x(x As Variant, Optional threshold_upper As Double = 0.95, Optional threshold_lower As Double = 0.05)
Dim i As Long, j As Long, k As Long, n As Long
Dim n_dimension As Long, i_min As Long, i_max As Long
Dim tmp_vec() As Double, sort_idx() As Long
Dim x_max As Double, x_min As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim tmp_vec(1 To n)
    For j = 1 To n_dimension
        For i = 1 To n
            tmp_vec(i) = x(i, j)
        Next i
        tmp_vec = PercentileScore(tmp_vec)
        Call Sort_Quick_A(tmp_vec, 1, n, sort_idx)
        For i = 1 To n
            If tmp_vec(i) >= threshold_lower Then
                i_min = i
                x_min = x(sort_idx(i), j)
                Exit For
            End If
        Next i
        For i = n To 1 Step -1
            If tmp_vec(i) <= threshold_upper Then
                i_max = i
                x_max = x(sort_idx(i), j)
                Exit For
            End If
        Next i
        For i = 1 To i_min
            x(sort_idx(i), j) = x_min
        Next i
        For i = n To i_max Step -1
            x(sort_idx(i), j) = x_max
        Next i
    Next j
    Erase tmp_vec, sort_idx
End Sub


'Normalize mutli-dimensional data x() by a chosen scheme
'Input: x(1 to n_raw,1 to dimension)
'Original data can be recovered by x*x_scale+x_shift
'if isKnown=TRUE then x() will be transfromed with supplied x_shift()
'and x_scale(), and stype is ignored.
Sub Normalize_x(x As Variant, x_shift() As Double, x_scale() As Double, _
        Optional stype As String = "AVGSD", Optional isKnown As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim n_raw As Long, n_dimension As Long
Dim tmp_vec() As Double, tmp_vec2() As Double
Dim tmp_x As Double, tmp_y As Double, tmp_med As Double, tmp_rng As Double
Dim tmp_min As Double, tmp_max As Double

n_raw = UBound(x, 1)
n_dimension = UBound(x, 2)

If isKnown = False Then

    ReDim x_shift(1 To n_dimension)
    ReDim x_scale(1 To n_dimension)
    
    If UCase(stype) = "AVGSD" Then
    
        For k = 1 To n_dimension
            tmp_x = 0
            tmp_y = 0
            For i = 1 To n_raw
                tmp_x = tmp_x + x(i, k)
                tmp_y = tmp_y + x(i, k) ^ 2
            Next i
            tmp_x = tmp_x / n_raw
            tmp_y = Sqr((tmp_y / n_raw - tmp_x ^ 2) * n_raw * 1# / (n_raw - 1))
            x_shift(k) = tmp_x
            x_scale(k) = tmp_y
        Next k
        
    ElseIf UCase(stype) = "DEMEAN" Then
    
        For k = 1 To n_dimension
            tmp_x = 0
            For i = 1 To n_raw
                tmp_x = tmp_x + x(i, k)
            Next i
            x_shift(k) = tmp_x / n_raw
            x_scale(k) = 1
        Next k
    
    ElseIf UCase(stype) = "MINMAX" Then
    
        For k = 1 To n_dimension
            tmp_min = x(1, k)
            tmp_max = x(1, k)
            For i = 2 To n_raw
                If x(i, k) > tmp_max Then tmp_max = x(i, k)
                If x(i, k) < tmp_min Then tmp_min = x(i, k)
            Next i
            x_shift(k) = tmp_min
            x_scale(k) = tmp_max - tmp_min
        Next k
    
    ElseIf UCase(stype) = "MINMAXAVG" Then
    
        For k = 1 To n_dimension
            tmp_x = x(1, k)
            tmp_min = x(1, k)
            tmp_max = x(1, k)
            For i = 2 To n_raw
                tmp_x = tmp_x + x(i, k)
                If x(i, k) > tmp_max Then tmp_max = x(i, k)
                If x(i, k) < tmp_min Then tmp_min = x(i, k)
            Next i
            tmp_x = tmp_x / n_raw
            x_shift(k) = tmp_x
            x_scale(k) = tmp_max - tmp_min
        Next k
    
    ElseIf UCase(stype) = "MEDRNG" Then
    
        For k = 1 To n_dimension
            Call get_vector(x, k, 2, tmp_vec)
            tmp_vec2 = fQuartile(tmp_vec)
            x_shift(k) = tmp_vec2(2)
            x_scale(k) = tmp_vec2(3) - tmp_vec2(1)
        Next k
    
    ElseIf UCase(stype) = "AVGSDEX" Then
    
        For k = 1 To n_dimension
            Call get_vector(x, k, 2, tmp_vec)
            tmp_vec2 = fQuartile(tmp_vec)
            tmp_med = tmp_vec2(2)
            tmp_rng = 3 * (tmp_vec2(3) - tmp_vec2(1))
            tmp_x = 0
            tmp_y = 0
            n = 0
            For i = 1 To n_raw
                If Abs(x(i, k) - tmp_med) < tmp_rng Then
                    n = n + 1
                    tmp_x = tmp_x + x(i, k)
                    tmp_y = tmp_y + x(i, k) ^ 2
                End If
            Next i
            tmp_x = tmp_x / n
            tmp_y = Sqr((tmp_y / n - tmp_x ^ 2) * n * 1# / (n - 1))
            x_shift(k) = tmp_x
            x_scale(k) = tmp_y
        Next k
    
    Else
        Debug.Print "Normalize_x: " & stype & " not a valid option."
    End If

End If

For k = 1 To n_dimension
    If x_scale(k) = 0 Then
        x_scale(k) = 1
        x_shift(k) = 0
    End If
    For i = 1 To n_raw
        x(i, k) = (x(i, k) - x_shift(k)) / x_scale(k)
    Next i
Next k
End Sub


Sub Recover_x(x() As Double, x_shift() As Double, x_scale() As Double)
Dim i As Long, k As Long
Dim n_raw As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    For k = 1 To n_dimension
        tmp_x = x_scale(k)
        tmp_y = x_shift(k)
        For i = 1 To n_raw
            x(i, k) = x(i, k) * tmp_x + tmp_y
        Next i
    Next k
End Sub


'Squash x to a range of [0,1] using 1/(1+exp(-x/x_scale))
'smaller x_scale gives more squeezing
'Input: x(1 to n_raw,1 to dimension)
Sub Squash_x(x() As Double, Optional x_scale As Double = 1)
Dim i As Long, k As Long
Dim n_raw As Long, n_dimension As Long
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    For k = 1 To n_dimension
        For i = 1 To n_raw
            x(i, k) = 1# / (1 + Exp(-x(i, k) / x_scale))
        Next i
    Next k
End Sub

Sub UnSquash_x(x() As Double, Optional x_scale As Double = 1)
Dim i As Long, k As Long
Dim n_raw As Long, n_dimension As Long
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    For k = 1 To n_dimension
        For i = 1 To n_raw
            x(i, k) = -x_scale * Log(1 / x(i, k) - 1)
        Next i
    Next k
End Sub


'=== Upsample a vector y(1:N) to y(1:((N-1)*mag+1)) by interpolation
'Input:  mag, number of points to interpolate
'        y(), vector of size 1:N
'        method, 1=linear, 3=cubic spline
'        x(), timestep, vector of size 1:N, if skipped data is assumed to be evenly spaced
'Output: y2(), upsampled vector of size 1:((N-1)*mag+1)
'        x2(), upsampled time step of size 1:((N-1)*mag+1)
Sub Interpol(mag As Long, y() As Double, y2() As Double, _
    Optional method As Long = 1, Optional x As Variant, Optional x2 As Variant)
    If method = 1 Then
        If IsMissing(x) = False Then
            If IsMissing(x2) Then
                Debug.Print "Interpol: x and x2 must both be supplied."
                Exit Sub
            End If
            Call Interpol_Linear(mag, y, y2, x, x2)
        Else
            Call Interpol_Linear_Regular(mag, y, y2)
        End If
    ElseIf method = 3 Then
        If IsMissing(x) = False Then
            If IsMissing(x2) Then
                Debug.Print "Interpol: x and x2 must both be supplied."
                Exit Sub
            End If
            Call Cubic_Spline(mag, y, y2, x, x2)
        Else
            Call Cubic_Spline_Regular(mag, y, y2)
        End If
    Else
        Debug.Print "Interpol: method must either be 1 (linear) or 3 (cubic)."
        Exit Sub
    End If
End Sub

Private Sub Interpol_Linear(mag As Long, y() As Double, y2() As Double, x As Variant, x2 As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim tmp_x As Double, tmp_y As Double
    n = UBound(y, 1)
    ReDim x2(1 To ((n - 1) * mag + 1))
    ReDim y2(1 To ((n - 1) * mag + 1))
    For i = 1 To n - 1
        ii = (i - 1) * mag + 1
        x2(ii) = x(i)
        y2(ii) = y(i)
        tmp_x = (x(i + 1) - x(i)) / mag
        tmp_y = (y(i + 1) - y(i)) / mag
        For j = 1 To mag - 1
            x2(ii + j) = x(i) + j * tmp_x
            y2(ii + j) = y(i) + j * tmp_y
        Next j
    Next i
    x2((n - 1) * mag + 1) = x(n)
    y2((n - 1) * mag + 1) = y(n)
End Sub

Private Sub Interpol_Linear_Regular(mag As Long, y() As Double, y2() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim tmp_x As Double, tmp_y As Double
    n = UBound(y, 1)
    ReDim y2(1 To ((n - 1) * mag + 1))
    For i = 1 To n - 1
        ii = (i - 1) * mag + 1
        y2(ii) = y(i)
        tmp_y = (y(i + 1) - y(i)) / mag
        For j = 1 To mag - 1
            y2(ii + j) = y(i) + j * tmp_y
        Next j
    Next i
    y2((n - 1) * mag + 1) = y(n)
End Sub

Private Sub Cubic_Spline(mag As Long, y() As Double, y2() As Double, x As Variant, x2 As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim A() As Double, u() As Double
Dim tmp_x As Double, tmp_y As Double, t As Double, tmp_v As Double, tmp_w As Double
    n = UBound(y, 1)
    ReDim A(1 To n, 1 To 3)
    ReDim u(1 To n)
    A(1, 2) = 2 / (x(2) - x(1))
    A(1, 3) = 1 / (x(2) - x(1))
    u(1) = 3 * (y(2) - y(1)) / ((x(2) - x(1)) ^ 2)
    A(n, 1) = 1 / (x(n) - x(n - 1))
    A(n, 2) = 2 / (x(n) - x(n - 1))
    u(n) = 3 * (y(n) - y(n - 1)) / ((x(n) - x(n - 1)) ^ 2)
    For i = 2 To n - 1
        tmp_x = x(i) - x(i - 1)
        tmp_y = x(i + 1) - x(i)
        A(i, 1) = 1 / tmp_x
        A(i, 2) = 2 / tmp_x + 2 / tmp_y
        A(i, 3) = 1 / tmp_y
        u(i) = 3 * ((y(i) - y(i - 1)) / (tmp_x ^ 2) + (y(i + 1) - y(i)) / (tmp_y ^ 2))
    Next i
    u = Solve_Tridiag(A, u)
    Erase A

    ReDim x2(1 To ((n - 1) * mag + 1))
    ReDim y2(1 To ((n - 1) * mag + 1))
    For i = 1 To n - 1
        ii = (i - 1) * mag + 1
        x2(ii) = x(i)
        y2(ii) = y(i)
        tmp_x = (x(i + 1) - x(i)) / mag
        tmp_v = u(i) * (x(i + 1) - x(i)) - (y(i + 1) - y(i))
        tmp_w = -u(i + 1) * (x(i + 1) - x(i)) + (y(i + 1) - y(i))
        For j = 1 To mag - 1
            t = j / mag
            x2(ii + j) = x(i) + j * tmp_x
            y2(ii + j) = (1 - t) * y(i) + t * y(i + 1) + t * (1 - t) * (tmp_v * (1 - t) + tmp_w * t)
        Next j
    Next i
    x2((n - 1) * mag + 1) = x(n)
    y2((n - 1) * mag + 1) = y(n)
    Erase u
End Sub

Private Sub Cubic_Spline_Regular(mag As Long, y() As Double, y2() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim A() As Double, u() As Double
Dim tmp_x As Double, tmp_y As Double, t As Double, tmp_v As Double, tmp_w As Double
    n = UBound(y, 1)
    ReDim A(1 To n, 1 To 3)
    ReDim u(1 To n)
    A(1, 2) = 2
    A(1, 3) = 1
    u(1) = 3 * (y(2) - y(1))
    A(n, 1) = 1
    A(n, 2) = 2
    u(n) = 3 * (y(n) - y(n - 1))
    For i = 2 To n - 1
        A(i, 1) = 1
        A(i, 2) = 4
        A(i, 3) = 1
        u(i) = 3 * (y(i + 1) - y(i - 1))
    Next i
    u = Solve_Tridiag(A, u)
    Erase A

    ReDim y2(1 To ((n - 1) * mag + 1))
    For i = 1 To n - 1
        ii = (i - 1) * mag + 1
        y2(ii) = y(i)
        tmp_x = 1 / mag
        tmp_v = u(i) - (y(i + 1) - y(i))
        tmp_w = -u(i + 1) + (y(i + 1) - y(i))
        For j = 1 To mag - 1
            t = j / mag
            y2(ii + j) = (1 - t) * y(i) + t * y(i + 1) + t * (1 - t) * (tmp_v * (1 - t) + tmp_w * t)
        Next j
    Next i
    y2((n - 1) * mag + 1) = y(n)
    Erase u
End Sub


'================================================
'Metrics
'================================================

'Input: x(1 to N, 1 to D), N x D-dimenisonal data
'Input: sqroot, true if square root needs to be applied
'Output: Dist(1 to N, 1 to N), N X N pairwise Euclidean distance
Function Calc_Euclidean_Dist(x As Variant, Optional sqroot As Boolean = False) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_raw As Long, n_dimension As Long
Dim dist() As Double, x1() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim dist(1 To n_raw, 1 To n_raw)
    ReDim x1(1 To n_dimension)
    For i = 1 To n_raw - 1
        For k = 1 To n_dimension
            x1(k) = x(i, k)
        Next k
        For j = i + 1 To n_raw
            tmp_x = 0
            For k = 1 To n_dimension
                tmp_x = tmp_x + (x1(k) - x(j, k)) ^ 2
            Next k
            dist(i, j) = tmp_x
        Next j
    Next i
    If sqroot = True Then
        For i = 1 To n_raw - 1
            For j = i + 1 To n_raw
                dist(i, j) = Sqr(dist(i, j))
            Next j
        Next i
    End If
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            dist(j, i) = dist(i, j)
        Next j
    Next i
    Calc_Euclidean_Dist = dist
    Erase x1, dist
End Function


'Input: x(1 to n_raw, 1 to Dimension), n_raw is the number of observations
'Output: Covar() is the Dimension X Dimension covariance matrix
Function Covariance_Matrix(x As Variant, Optional isSample As Long = 1) As Double()
Dim i As Long, j As Long, m As Long, n As Long, d As Long, nn As Long
Dim n_raw As Long, dimension As Long
Dim x_avg() As Double, covar() As Double
Dim tmp_x As Double, tmp_y As Double

    n_raw = UBound(x, 1)
    dimension = UBound(x, 2)
    If isSample = 1 Then
        nn = n_raw - 1
    Else
        nn = n_raw
    End If
    ReDim covar(1 To dimension, 1 To dimension)
    ReDim x_avg(1 To dimension)
    For d = 1 To dimension
        tmp_x = 0
        For i = 1 To n_raw
            tmp_x = tmp_x + x(i, d)
        Next i
        x_avg(d) = tmp_x / n_raw
    Next d
    
    For n = 1 To dimension
        tmp_x = 0
        tmp_y = x_avg(n)
        For i = 1 To n_raw
            tmp_x = tmp_x + (x(i, n) - tmp_y) ^ 2
        Next i
        covar(n, n) = tmp_x / nn
    Next n
    
    For m = 1 To dimension - 1
        For n = m + 1 To dimension
            tmp_x = 0
            For i = 1 To n_raw
                tmp_x = tmp_x + (x(i, m) - x_avg(m)) * (x(i, n) - x_avg(n))
            Next i
            covar(m, n) = tmp_x / nn
            covar(n, m) = covar(m, n)
        Next n
    Next m
    
    Covariance_Matrix = covar
    Erase x_avg, covar
End Function



'Input: x(1 to n_raw, 1 to Dimension), n_raw is the number of observations
'Output: Covar() is the Dimension X Dimension covariance matrix
Function Correl_Matrix(x As Variant) As Double()
Dim i As Long, j As Long, m As Long, n As Long, d As Long
Dim n_raw As Long, n_dimension As Long
Dim x_avg() As Double, x_sse() As Double
Dim y() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim x_avg(1 To n_dimension)
    ReDim x_sse(1 To n_dimension)
    ReDim y(1 To n_dimension, 1 To n_dimension)
    
    For d = 1 To n_dimension
        For i = 1 To n_raw
            x_avg(d) = x_avg(d) + x(i, d)
        Next i
        x_avg(d) = x_avg(d) / n_raw
        For i = 1 To n_raw
            x_sse(d) = x_sse(d) + (x(i, d) - x_avg(d)) ^ 2
        Next i
    Next d
    
    y(n_dimension, n_dimension) = 1
    For m = 1 To n_dimension - 1
        y(m, m) = 1
        For n = m + 1 To n_dimension
            If x_sse(m) > 0 And x_sse(n) > 0 Then
                For i = 1 To n_raw
                    y(m, n) = y(m, n) + (x(i, m) - x_avg(m)) * (x(i, n) - x_avg(n))
                Next i
                y(m, n) = y(m, n) / Sqr(x_sse(m) * x_sse(n))
                y(n, m) = y(m, n)
            End If
        Next n
    Next m
    
    Correl_Matrix = y
    Erase y, x_avg, x_sse
End Function


'Complex Correlation Matrix of x(1 to N, 1 to D)
'Input:     x(1 to N, 1 to D) is a matrix holding D price series of length N
'           note that x() used here is log(price) level, not return
'           min_scale, finest scale to look at
'Output:    r() is a DxD matrix holding the magnitudes
'           r_phase() is a DxD matrix of phase angles
Sub Correl_Matrix_Complex(x() As Double, r() As Double, r_phase() As Double, Optional min_scale As Long = 1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_raw As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double
Dim tmp_vec() As Double, tmp_vec2() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim r(1 To n_dimension, 1 To n_dimension)
    ReDim r_phase(1 To n_dimension, 1 To n_dimension)
    ReDim tmp_vec(1 To n_raw)
    ReDim tmp_vec2(1 To n_raw)
    For m = 1 To n_dimension
        For i = 1 To n_raw
            tmp_vec(i) = x(i, m)
        Next i
        r(m, m) = 1
        r_phase(m, m) = 0
        For n = m + 1 To n_dimension
            For i = 1 To n_raw
                tmp_vec2(i) = x(i, n)
            Next i
            Call Complex_Correl(tmp_vec, tmp_vec2, tmp_x, tmp_y, min_scale)
            r(m, n) = tmp_x
            r(n, m) = tmp_x
            r_phase(m, n) = tmp_y
            r_phase(n, m) = -tmp_y
        Next n
    Next m
    Erase tmp_vec, tmp_vec2
End Sub


'Return the complex correlation betweem x() and y(), which are 1-dimensional real series
'"Complex Correlation Approach for High Frequency Financial Data", Mateusz Willinski, Yiochi Ikeda (2017)
'note that x() and y() used here are log(price) level, not return
'positive r_phase means x leads y
Sub Complex_Correl(x() As Double, y() As Double, r As Double, r_phase As Double, Optional min_scale As Long = 1)
Dim k As Long, n As Long, n_k As Long
Dim axy_re As Double, axy_im As Double, axx As Double, ayy As Double
Dim r_re As Double, r_im As Double
Dim ax() As Double, bx() As Double, ay() As Double, by() As Double
    n_k = Int(UBound(x, 1) / (2 * min_scale))
    Call Fourier_Coeff(x, ax, bx, min_scale)
    Call Fourier_Coeff(y, ay, by, min_scale)
    axy_re = 0: axy_im = 0
    axx = 0: ayy = 0
    For k = 1 To n_k
        axy_re = axy_re + ax(k) * ay(k) + bx(k) * by(k)
        axy_im = axy_im + ax(k) * by(k) - ay(k) * bx(k)
        axx = axx + (ax(k) ^ 2) + (bx(k) ^ 2)
        ayy = ayy + (ay(k) ^ 2) + (by(k) ^ 2)
    Next k
    r_re = axy_re / Sqr(axx * ayy)
    r_im = axy_im / Sqr(axx * ayy)
    r = Sqr((r_re ^ 2) + (r_im ^ 2))
    r_phase = complex_phase(r_re, r_im)
    Erase ax, bx, ay, by
End Sub


'Return the phase of a complex number, range between [-pi, pi]
Function complex_phase(x_re As Double, x_im As Double) As Double
Dim pi As Double
    pi = 3.14159265358979       '4 * VBA.Atn(1)
    If x_re > 0 Then
        complex_phase = VBA.Atn(x_im / x_re)
    ElseIf x_re < 0 Then
        If x_im >= 0 Then
            complex_phase = VBA.Atn(x_im / x_re) + pi
        Else
            complex_phase = VBA.Atn(x_im / x_re) - pi
        End If
    ElseIf x_re = 0 Then
        If x_im > 0 Then
            complex_phase = pi / 2
        ElseIf x_im < 0 Then
            complex_phase = -pi / 2
        Else
            Debug.Print "complex_phase: inputs are zeroes."
        End If
    End If
End Function


'Fourier Transfrom used by Complex_Correl(), this is performed directly on log-level index, not difference
Private Sub Fourier_Coeff(x() As Double, A() As Double, B() As Double, Optional min_scale As Long = 1)
Dim i As Long, k As Long, n As Long, n_k As Long
Dim tmp_x As Double, tmp_y As Double, pi As Double
Dim t() As Double
    pi = 3.14159265358979
    n = UBound(x, 1)
    n_k = n \ (2 * min_scale)
    ReDim A(1 To n_k)
    ReDim B(1 To n_k)
    ReDim t(1 To n)
    For i = 1 To n
        t(i) = (i - 2) * 2 * pi / (n - 2)
    Next i
    For k = 1 To n_k
        tmp_x = 0
        tmp_y = 0
        For i = 2 To n - 1
            tmp_x = tmp_x + x(i) * (Cos(k * t(i + 1)) - Cos(k * t(i)))
            tmp_y = tmp_y + x(i) * (Sin(k * t(i + 1)) - Sin(k * t(i)))
        Next i
        A(k) = (x(n) - x(1) - tmp_x) / pi
        B(k) = -tmp_y / pi
    Next k
    Erase t
End Sub


'Calculate the autocorrelation of x() with a given lag, default lag is 1
'Input: x(0 to N)
Function autocorrelation(x() As Double, Optional lag As Long = 1) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, avg1 As Double, avg2 As Double, sse1 As Double, sse2 As Double
    n = UBound(x, 1) - LBound(x, 1) - lag + 1
    avg1 = 0
    avg2 = 0
    sse1 = 0
    sse2 = 0
    For i = LBound(x, 1) + lag To UBound(x, 1)
        avg1 = avg1 + x(i)
        avg2 = avg2 + x(i - lag)
    Next i
    avg1 = avg1 / n
    avg2 = avg2 / n
    
    tmp_x = 0
    For i = LBound(x, 1) + lag To UBound(x, 1)
        tmp_x = tmp_x + (x(i) - avg1) * (x(i - lag) - avg2)
        sse1 = sse1 + (x(i) - avg1) ^ 2
        sse2 = sse2 + (x(i - lag) - avg2) ^ 2
    Next i
    autocorrelation = tmp_x / Sqr(sse1 * sse2)
End Function


'=== Returns the cross-correlation (x*y)(t)= Integrate<x(i)*y(i+t)>
'=== positive lag_max means x leads y, negative lag_max means x lags y.
'=== Note that x & y needs to be 0-based
'Input:     x(0:N-1) and y(0:N-1) are real vectors
'Output:    CrossCorrel(0:N*-1,1:2), 1st dimension is the number of lags,
'               2nd dimension is the corresponding correlation
'           N* is the number of N+zero paddings s.t. N* is power of 2
'           lag_max, the number of lags that gives the maximum correlation
'           r_max, maximum correlation
Function CrossCorrel(x() As Double, y() As Double, Optional lag_max As Variant, Optional r_max As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim xDFT As cDFT, yDFT As cDFT
Dim xg() As Double, yg() As Double, xyg() As Double
Dim correl_out() As Double

    Set xDFT = New cDFT
    Set yDFT = New cDFT
    
    Call xDFT.FFT(x): xg = xDFT.g: Call xDFT.Reset
    Call yDFT.FFT(y): yg = yDFT.g: Call yDFT.Reset: Set yDFT = Nothing
    
    xyg = xg
    n = UBound(xg, 1) + 1
    For i = LBound(xg, 1) To UBound(xg, 1)
        xyg(i, 1) = (xg(i, 1) * yg(i, 1) + xg(i, 2) * yg(i, 2))
        xyg(i, 2) = (xg(i, 1) * yg(i, 2) - xg(i, 2) * yg(i, 1))
    Next i
    
    Call xDFT.FFT_Inverse(xyg, xg, yg)
    Call xDFT.Reset
    Set xDFT = Nothing
    
    n = UBound(xg) + 1
    ReDim correl_out(0 To n - 1, 1 To 2)
    For i = 0 To n - 1
        If i < n / 2 Then  'n must be power of 2 since FFT above has zero padded the series
            correl_out(i, 1) = i
        Else
            correl_out(i, 1) = -(n - i)
        End If
        correl_out(i, 2) = xg(i) / n
        If Abs(yg(i)) > 0.0000001 Then
            Debug.Print "CrossCorrel: Failed: non-zero imaginary part. (" & yg(i) & ")"
            Erase correl_out, xg, yg, xyg
            Exit Function
        End If
    Next i
    CrossCorrel = correl_out
    
    If IsMissing(lag_max) = False Then
        r_max = 0
        lag_max = 0
        For i = 0 To n - 1
            If Abs(xg(i)) > Abs(r_max) Then
                r_max = xg(i)
                If i < (n / 2) Then
                    lag_max = i
                Else
                    lag_max = -(n - i)
                End If
            End If
        Next i
        r_max = r_max / n
    End If
    
    Erase correl_out, xg, yg, xyg
End Function



'=== Returns the auto-correlation spectrum (x*x)(t)= <x(i)*x(i+t)>
'=== Note that x needs to be 0-based
'Input:     x(0:N-1), real vector
'Output:    CrossCorrel(0:N*-1,1:2), 1st dimension is the number of lags,
'               2nd dimension is the corresponding correlation
'           N* is the number of N+zero paddings s.t. N* is power of 2
Function CrossCorrel_Auto(x() As Double) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim xDFT As cDFT
Dim xg() As Double, yg() As Double, x2g() As Double
Dim correl_out() As Double

    Set xDFT = New cDFT
    Call xDFT.FFT(x)
    xg = xDFT.g
    Call xDFT.Reset
    n = UBound(xg, 1) + 1
    ReDim x2g(0 To n - 1, 1 To 2)
    For i = LBound(xg, 1) To UBound(xg, 1)
        x2g(i, 1) = xg(i, 1) ^ 2 + xg(i, 2) ^ 2
    Next i
    Call xDFT.FFT_Inverse(x2g, xg, yg)
    Erase x2g
    Call xDFT.Reset
    Set xDFT = Nothing
    n = UBound(xg) + 1
    ReDim correl_out(0 To n - 1, 1 To 2)
    For i = 0 To n - 1
        If i < n / 2 Then  'n must be power of 2 since FFT above has zero padded the series
            correl_out(i, 1) = i
        Else
            correl_out(i, 1) = -(n - i)
        End If
        correl_out(i, 2) = xg(i) / n
        If Abs(yg(i)) > 0.0000001 Then
            Debug.Print "CrossCorrel_Auto: Failed: non-zero imaginary part. (" & yg(i) & ")"
            Exit Function
        End If
    Next i
    CrossCorrel_Auto = correl_out
    Erase correl_out, xg, yg
End Function










'========================================
'Categorical data
'========================================

'Convert a categorical vector x(1 to N) into binary matrix
'Output: v(1 to N, 1 to n_class)
'Output: class_map(1 to n_class)
Sub Class2Vec(x As Variant, v As Variant, n_class As Long, class_map As Variant, _
    Optional mapIsKnown As Boolean = False)
Dim i As Long, j As Long, m As Long, n As Long
Dim record_count() As Long, x_idx() As Long
Dim tmp_x As Variant
    n = UBound(x)
    If mapIsKnown = False Then
        Call Unique_Items(x, x_idx, class_map, n_class, record_count, True)
        ReDim v(1 To n, 1 To n_class)
        For i = 1 To n
            v(i, x_idx(i)) = 1
        Next i
        Erase record_count, x_idx
    ElseIf mapIsKnown = True Then
        n_class = UBound(class_map, 1)
        ReDim v(1 To n, 1 To n_class)
        For i = 1 To n
            For j = 1 To n_class
                If x(i) = class_map(j) Then
                    v(i, j) = 1
                    Exit For
                End If
            Next j
        Next i
    End If
End Sub

'Convert binary matrix v() back to categorical vector x()
'Output: x(1 to N)
Sub Vec2Class(v As Variant, class_map As Variant, x As Variant)
Dim i As Long, j As Long, m As Long, n As Long, n_class As Long
Dim tmp_max As Double
    n = UBound(v, 1)
    n_class = UBound(v, 2)
    ReDim x(1 To n)
    For i = 1 To n
        tmp_max = -999999
        For j = 1 To n_class
            If v(i, j) > tmp_max Then
                tmp_max = v(i, j)
                x(i) = class_map(j)
            End If
        Next j
    Next i
End Sub


'Input: 1D array x()
'Output: x_list() containing x_count unique items from x()
'Output: x_i(), integer representation of x() where the integer point to the item in x_list()
Sub Unique_Items(x As Variant, x_i() As Long, x_list As Variant, x_count As Long, x_list_size() As Long, _
        Optional show_largest_first As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, isUnique As Long
    n = UBound(x, 1)
    ReDim x_i(1 To n)
    ReDim x_list(1 To n)
    ReDim x_list_size(1 To n)
    x_count = 1
    x_list(1) = x(1)
    x_i(1) = 1
    x_list_size(1) = 1
    For i = 2 To n
        isUnique = 1
        For j = 1 To x_count
            If x_list(j) = x(i) Then
                isUnique = 0
                x_i(i) = j
                x_list_size(j) = x_list_size(j) + 1
                Exit For
            End If
        Next j
        If isUnique = 1 Then
            x_count = x_count + 1
            x_list(x_count) = x(i)
            x_i(i) = x_count
            x_list_size(x_count) = 1
        End If
    Next i
    ReDim Preserve x_list(1 To x_count)
    ReDim Preserve x_list_size(1 To x_count)
    If show_largest_first = True And x_count > 1 Then
        Dim iArr() As Long, sort_idx() As Long
        Dim tmpList As Variant
        iArr = x_list_size
        For i = 1 To x_count
            iArr(i) = -iArr(i)
        Next i
        Call Sort_Quick_A(iArr, 1, x_count, sort_idx, 1)
        tmpList = x_list
        For j = 1 To x_count
            x_list(j) = tmpList(sort_idx(j))
            x_list_size(j) = -iArr(j)
        Next j
        ReDim iArr(1 To x_count)
        For j = 1 To x_count
            iArr(sort_idx(j)) = j
        Next j
        For i = 1 To n
            x_i(i) = iArr(x_i(i))
        Next i
        Erase sort_idx, iArr, tmpList
    End If
End Sub







'=== Brute Force Search for k-Nearest Neighbors
'Input: x(1 to N, 1 to D), N x D-dimensional data
Sub kNN_Quadratic(k_idx() As Long, k_dist() As Double, x() As Double, n_neighbors As Long, Optional kth_only As Long = 1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_raw As Long, n_dimension
Dim tmp_x As Double, tmp_y As Double
Dim d() As Double, neighbors() As Long
Dim dist() As Double
Dim Output() As Double

n_raw = UBound(x, 1)
n_dimension = UBound(x, 2)

ReDim k_idx(1 To n_raw, 1 To n_neighbors)
ReDim k_dist(1 To n_raw, 1 To n_neighbors)

ReDim dist(1 To n_raw, 1 To n_raw)
For i = 1 To n_raw - 1
    For j = i + 1 To n_raw
        tmp_x = 0
        For k = 1 To n_dimension
            tmp_x = tmp_x + (x(i, k) - x(j, k)) ^ 2
        Next k
        dist(i, j) = tmp_x
        dist(j, i) = tmp_x
    Next j
Next i

ReDim Output(1 To n_raw, 1 To 2)

ReDim d(1 To n_raw - 1)
ReDim neighbors(1 To n_raw - 1)
For i = 1 To n_raw

    If i Mod 100 = 0 Then
        DoEvents
        Application.StatusBar = "kNN_Quadratic: " & i & "/" & n_raw
    End If
    
    k = 0
    For j = 1 To i - 1
            k = k + 1
            neighbors(k) = j
            d(k) = dist(i, j)
    Next j
    For j = i + 1 To n_raw
            k = k + 1
            neighbors(k) = j
            d(k) = dist(i, j)
    Next j
    Call Sort_Quick_A(d, 1, n_raw - 1, neighbors, 0)
    
    For k = 1 To n_neighbors
        k_idx(i, k) = neighbors(k)
        k_dist(i, k) = Sqr(d(k))
    Next k
Next i

If kth_only = 1 Then
    ReDim d(1 To n_raw)
    ReDim neighbors(1 To n_raw)
    For i = 1 To n_raw
        neighbors(i) = k_idx(i, n_neighbors)
        d(i) = k_dist(i, n_neighbors)
    Next i
    k_idx = neighbors
    k_dist = d
End If
Application.StatusBar = False
End Sub



'=======================================================
'Curve Fitting
'=======================================================

'=== LOWESS (locally weighted scatterplot smoothing)
'Fortran implementaion from : http://www.netlib.org/go/lowess.f
'=======================================================
Function LOWESS(y_in() As Double, x_in() As Double, Optional smoothing As Double = 0.25, Optional nsteps As Long = 2, Optional delta As Double = 0) As Double()
Dim i As Long, j As Long, k As Long, iterate As Long, M1 As Long, M2 As Long
Dim n_raw As Long
Dim ys() As Double, res() As Double, rw() As Double
Dim nleft As Long, nright As Long, ns As Long
Dim d1 As Double, d2 As Double, denom As Double, cut As Double, alpha As Double
Dim c1 As Double, c9 As Double, cmad As Double, r As Double
Dim OK As Boolean, itertest As Boolean
Dim sort_index() As Long
Dim y() As Double, x() As Double

n_raw = UBound(y_in, 1)

If n_raw < 2 Then
    LOWESS = y_in
    Exit Function
End If

y = y_in
x = x_in
Call Sort_Quick_A(x, 1, n_raw, sort_index)
For i = 1 To n_raw
    y(i) = y_in(sort_index(i))
Next i

ReDim ys(1 To n_raw)
ReDim rw(1 To n_raw)
ReDim res(1 To n_raw)

For i = 2 To n_raw
    If x(i) < x(i - 1) Then
        msgbox "x() is not sorted in ascending order yet."
        Exit Function
    End If
Next i

ns = Int(smoothing * n_raw)
If n_raw < ns Then ns = n_raw
If 2 > ns Then ns = 2

For iterate = 1 To nsteps + 1
    nleft = 1
    nright = ns
    k = 0
    i = 1
    Do
        Do While nright < n_raw
            d1 = x(i) - x(nleft)
            d2 = x(nright + 1) - x(i)
            If d1 <= d2 Then Exit Do
            nleft = nleft + 1
            nright = nright + 1
        Loop
        
        If iterate > 1 Then
            itertest = True
        Else
            itertest = False
        End If
        
        Call LOWEST(x, y, x(i), ys(i), nleft, nright, res, itertest, rw, OK)
        
        If OK = False Then ys(i) = y(i)
        
        If k < i - 1 Then
            denom = x(i) - x(k)
            For j = k + 1 To i - 1
                alpha = (x(j) - x(k)) / denom
                ys(j) = alpha * ys(i) + (1 - alpha) * ys(k)
            Next j
        End If
        k = i
        cut = x(k) + delta
        For i = k + 1 To n_raw
            If x(i) > cut Then Exit For
            If x(i) = x(k) Then
                ys(i) = ys(k)
                k = i
            End If
        Next i
        If (k + 1) > (i - 1) Then
            i = k + 1
        Else
            i = (i - 1)
        End If
        'i = MAX2(k + 1, i - 1)
    Loop Until k >= n_raw
    
    For i = 1 To n_raw
        res(i) = y(i) - ys(i)
    Next i
    If iterate > nsteps Then Exit For
    For i = 1 To n_raw
        rw(i) = Abs(res(i))
    Next i
    Call Sort_Quick(rw, 1, n_raw)
    M1 = 1 + n_raw / 2
    M2 = n_raw - M1 + 1
    cmad = 3 * (rw(M1) + rw(M2))
    c9 = 0.999 * cmad
    c1 = 0.001 * cmad
    For i = 1 To n_raw
        r = Abs(res(i))
        If r <= c1 Then
            rw(i) = 1
        ElseIf r > c9 Then
            rw(i) = 0
        Else
            rw(i) = (1 - (r / cmad) ^ 2) ^ 2
        End If
    Next i
Next iterate

For i = 1 To n_raw
    y(sort_index(i)) = ys(i)
Next i

LOWESS = y
End Function


Private Sub LOWEST(x() As Double, y() As Double, xS As Double, ys As Double, _
            nleft As Long, nright As Long, w() As Double, userw As Boolean, rw() As Double, OK As Boolean)

Dim i As Long, j As Long
Dim n_raw As Long, nrt As Long
Dim x_rng As Double
Dim h As Double, h9 As Double, h1 As Double
Dim A As Double, B As Double, c As Double, r As Double

n_raw = UBound(y, 1)
x_rng = x(n_raw) - x(1)
h = xS - x(nleft)
If (x(nright) - xS) > h Then h = x(nright) - xS
h9 = 0.999 * h
h1 = 0.001 * h
A = 0
For j = nleft To n_raw
    w(j) = 0
    r = Abs(x(j) - xS)
    If r <= h9 Then
        If r > h1 Then
            w(j) = (1 - (r / h) ^ 3) ^ 3
        Else
            w(j) = 1
        End If
        If userw = True Then w(j) = rw(j) * w(j)
        A = A + w(j)
    ElseIf x(j) > xS Then
        Exit For
    End If
Next j
nrt = j - 1
If A <= 0 Then
    OK = False
Else
    OK = True
    For j = nleft To nrt
        w(j) = w(j) / A
    Next j
    If h > 0 Then
        A = 0
        For j = nleft To nrt
            A = A + w(j) * x(j)
        Next j
        B = xS - A
        c = 0
        For j = nleft To nrt
            c = c + w(j) * (x(j) - A) ^ 2
        Next j
        If Sqr(c) > 0.001 * x_rng Then
            B = B / c
            For j = nleft To nrt
                w(j) = w(j) * (1 + B * (x(j) - A))
            Next j
        End If
    End If
    
    ys = 0
    For j = nleft To nrt
        ys = ys + w(j) * y(j)
    Next j
End If

End Sub


Function ANN_Curve_Fit(y() As Double, x() As Double, _
        Optional num_neurons As Long = 13, _
        Optional learn_rate As Double = 0.5, _
        Optional momentum As Double = 0.9, _
        Optional mini_batch As Long = 100, _
        Optional iterate_max As Long = 1000, _
        Optional mse_min As Double = 0.001, _
        Optional L1 As Double = 0, Optional L2 As Double = 0, Optional LMAX As Double = 0) As Double()
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim n_raw As Long, n_dimension As Long
Dim y_shift() As Double, y_scale() As Double
Dim target() As Double, x_train() As Double
Dim ANN1 As cANN_Regression
    n_raw = UBound(y)
    If getDimension(x) = 1 Then
        n_dimension = 1
        Call Promote_Vec(x, x_train)
    Else
        n_dimension = UBound(x, 2)
        x_train = x
    End If
    Call Promote_Vec(y, target)
    Call Normalize_x(x_train, y_shift, y_scale, "MEDRNG")
    Call Normalize_x(target, y_shift, y_scale, "MEDRNG")
    Call Squash_x(target, 1)
    m = mini_batch
    If m = 0 Then m = n_raw
    Set ANN1 = New cANN_Regression
    With ANN1
        Call .Init(n_dimension, 1, num_neurons)
        Call .Trainer(x_train, target, learn_rate, momentum, m, iterate_max, mse_min, L1, L2, LMAX)
        Call .InOut(x_train, target)
        Call .Reset
    End With
    Set ANN1 = Nothing
    Call modMath.UnSquash_x(target, 1)
    Call Recover_x(target, y_shift, y_scale)
    Call get_vector(target, 1, 2, ANN_Curve_Fit)
    Erase target, x_train, y_shift, y_scale
End Function



'====================
'File access
'====================

'==== Wait until a file is unlock due to slow read/write to network drive
Sub Wait_for_File_Access(strFileName As String, Optional wait_count As Long = 60)
Dim k As Long
    k = 1
    Do Until k = wait_count Or FileIsLocked(strFileName) = False
        Application.Wait (Now + TimeValue("00:00:01"))
        k = k + 1
    Loop
End Sub

Function FileIsLocked(strFileName As String) As Boolean
    FileIsLocked = False
    On Error Resume Next
    Open strFileName For Binary Access Read Write Lock Read Write As #1
    Close #1

    ' If an error occurs, the document is currently open
    If Err.Number <> 0 Then
       FileIsLocked = True
       Err.Clear
    End If
End Function


'=== Read mutli-dimensional data from CSV file
'Input: strPath, path to csv file
'       strDelimiter, delimiter to use
'Output: y(1:N), vector of size N from the first column, N is number of lines except the first.
'        x(1:N,1:D), array of size NxD where D is the number of columns except the first.
'        x_labels(1:D), column label of x() if first line is header
Sub Read_csv(strPath As String, y As Variant, x As Variant, _
            Optional strDelimiter As String = ",", _
            Optional first_header As Boolean = True, Optional x_labels As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, vFF As Long
Dim n_raw As Long, n_dimension As Long
Dim strArr As Variant, xArr As Variant
Dim FileContents As String

    vFF = VBA.FreeFile
    Open strPath For Binary Access Read As #vFF
    FileContents = VBA.Space$(VBA.LOF(vFF))
    Get #vFF, , FileContents
    Close #vFF
    
    strArr = VBA.Split(FileContents, vbCrLf)    '0-base matrix
    xArr = VBA.Split(strArr(0), strDelimiter)   '0-base vector
    n_raw = UBound(strArr, 1)
    If VBA.Len(strArr(n_raw)) = 0 Then n_raw = n_raw - 1 'delete last row if empty
    If first_header = False Then n_raw = n_raw + 1
    n_dimension = UBound(xArr, 1)
    
    'Read column headings
    If IsMissing(x_labels) = False Then
        ReDim x_labels(1 To n_dimension)
        If first_header = True Then
            For j = 1 To n_dimension
                x_labels(j) = xArr(j)
            Next j
        Else
            For j = 1 To n_dimension
                x_labels(j) = VBA.Format(j, "000")
            Next j
        End If
    End If

    'Start reading data from second line if first
    'line is header, read from first line otherwise
    k = 0
    If first_header = False Then k = -1
    ReDim x(1 To n_raw, 1 To n_dimension)
    ReDim y(1 To n_raw)
    For i = 1 To n_raw
        If i Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "Reading csv: " & i & "/" & n_raw
        End If
        xArr = VBA.Split(strArr(i + k), strDelimiter)
        y(i) = xArr(0)
        For j = 1 To n_dimension
            x(i, j) = xArr(j)
        Next j
    Next i
    
    Erase strArr, xArr
    Application.StatusBar = False
End Sub


'=== Read selected range of lines text file
Sub Read_csv_by_Line(strPath As String, x As Variant, _
            Optional strDelimiter As String = ",", _
            Optional first_row_is_header As Boolean = True, Optional x_labels As Variant, _
            Optional first_col_is_label As Boolean = True, Optional y As Variant, _
            Optional first_row As Long = 1, Optional n_line As Long = -1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, vFF As Long
Dim n_raw As Long, n_dimension As Long, i_start As Long, j_start As Long
Dim strArr As Variant, xArr As Variant
Dim strLine As String

    vFF = VBA.FreeFile
    Open strPath For Input As #vFF
    i = 0
    n_raw = 0
    ReDim strArr(1 To 1)
    Do Until VBA.EOF(vFF)
        If i Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "Reading csv 1: " & i & "/" & n_raw
        End If
        Line Input #vFF, strLine
        i = i + 1
        If i >= first_row Then
            n_raw = n_raw + 1
            ReDim Preserve strArr(1 To n_raw)
            strArr(n_raw) = strLine
            If n_raw = n_line Then Exit Do
        End If
    Loop
    Close #vFF
    
    If VBA.Len(strArr(n_raw)) = 0 Then n_raw = n_raw - 1 'delete last line if empty
    If first_row_is_header = True Then
        n_raw = n_raw - 1
        i_start = 2
    Else
        i_start = 1
    End If
    
    m = UBound(VBA.Split(strArr(i_start), strDelimiter), 1) + 1 '0-base vector
    If first_col_is_label = True Then
        n_dimension = m - 1
        j_start = 1
    Else
        n_dimension = m
        j_start = 0
    End If
    
    'Read column headings
    If IsMissing(x_labels) = False Then
        ReDim x_labels(1 To n_dimension)
        If first_row_is_header = True Then
            xArr = VBA.Split(strArr(1), strDelimiter)   '0-base vector
            For i = 1 To n_dimension
                x_labels(i) = xArr(j_start + i - 1)
            Next i
        Else
            For i = 1 To n_dimension
                x_labels(i) = VBA.Format(i, "000")
            Next i
        End If
    End If
    
    'Reading data
    ReDim x(1 To n_raw, 1 To n_dimension)
    If IsMissing(y) = False Then ReDim y(1 To n_raw)
    For i = 1 To n_raw
        If i Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "Reading csv 2: " & i & "/" & n_raw
        End If
        xArr = VBA.Split(strArr(i_start + i - 1), strDelimiter)
        For j = 1 To n_dimension
            x(i, j) = xArr(j_start + j - 1)
        Next j
        If IsMissing(y) = False Then
            If first_col_is_label = True Then
                y(i) = xArr(0)
            Else
                y(i) = i
            End If
        End If
    Next i
    
    Erase strArr, xArr
    Application.StatusBar = False
End Sub

'========================================
'Probability Distribution
'========================================

'Function for charting contour plot of 2D-Gaussian
'Input: x_mean(1:2), vector of mean values
'       x_covar(1:2,1:2), covariance matrix
'       n_level, number of contours to displau
'       n_sample, number of sample points along each contour
'       show_p, if set to TRUE add a thrid column to the output showing probability density
'Output: variant array of size (1:n_level*(n_sample+1)-1, 1:2)
Function Gaussian_Contour(x_mean() As Double, x_covar() As Double, _
    Optional n_level As Long = 5, Optional n_sample As Long = 60, Optional show_p As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double
Dim rho As Double, x_sd As Double, y_sd As Double
Dim p_max As Double, norm As Double, p As Double
Dim cs As Double, ss As Double, theta As Double
Dim x_pos As Variant
    If show_p = False Then
        ReDim x_pos(1 To n_level * (n_sample + 1) - 1, 1 To 2)
    Else
        ReDim x_pos(1 To n_level * (n_sample + 1) - 1, 1 To 3)
    End If
    n_dimension = UBound(x_mean)
    If n_dimension <> 2 Then
        Debug.Print "Gaussian_Contour: Error: only works for 2D-plot."
        Exit Function
    End If
    x_sd = Sqr(x_covar(1, 1))
    y_sd = Sqr(x_covar(2, 2))
    rho = x_covar(1, 2) / (x_sd * y_sd)
    If Abs(rho) > 1 Then
        Debug.Print "Gaussian_Contour: Error: covariance matrix is not positive definite."
        Exit Function
    End If
    norm = 6.28318530717959 * x_sd * y_sd * Sqr(1 - rho ^ 2)
    p_max = 1 / norm
    m = 0
    For k = 1 To n_level
        p = p_max / (2 ^ k)
        tmp_y = -2 * (1 - rho ^ 2) * Log(norm * p)
        For i = 1 To n_sample
            theta = (i - 1) * 6.28318530717959 / (n_sample - 1)
            cs = VBA.Cos(theta)
            ss = VBA.Sin(theta)
            tmp_x = Sqr(tmp_y / (cs ^ 2 + ss ^ 2 - 2 * rho * cs * ss))
            x_pos(m + i, 1) = tmp_x * cs * x_sd + x_mean(1)
            x_pos(m + i, 2) = tmp_x * ss * y_sd + x_mean(2)
            If show_p = True Then x_pos(m + i, 3) = p
        Next i
        m = m + n_sample + 1
    Next k
    Gaussian_Contour = x_pos
    Erase x_pos
End Function


'Return the histogram for a 1-D array as probability density
'Input: x(1 to N), 1-D array
'Input: n_bin, desired number of bins
'Output: x_hist(1 to n_bin, 1 to 2), first number is bin position, second is probability density
'if output_fit is set to TRUE then a third column is included to show the maximum likelihood fit
'of the given fit_type
Function Histogram_Create(x As Variant, Optional n_bin As Long = 30, _
        Optional output_fit As Boolean = False, Optional fit_type As String = "GAUSSIAN") As Double()
Dim i As Long, k As Long, n As Long
Dim x_max As Double, x_min As Double, dx As Double
Dim x_hist() As Double
Dim x_loc As Double, x_scale As Double, x_asym As Double, likelihood As Double
    n = UBound(x, 1) - LBound(x, 1) + 1
    x_max = -Exp(70)
    x_min = Exp(70)
    For i = LBound(x, 1) To UBound(x, 1)
        If x(i) > x_max Then x_max = x(i)
        If x(i) < x_min Then x_min = x(i)
    Next i
    dx = 1.005 * (x_max - x_min) / n_bin
    ReDim x_hist(1 To n_bin, 1 To 2)
    For i = LBound(x, 1) To UBound(x, 1)
        k = 1 + Int((x(i) - x_min) / dx)
        x_hist(k, 2) = x_hist(k, 2) + 1
    Next i
    For k = 1 To n_bin
        x_hist(k, 1) = x_min + (k - 0.5) * dx
        x_hist(k, 2) = x_hist(k, 2) / (n * dx)
    Next k
    
    If output_fit = True Then
        likelihood = Prob_Fit(x, x_loc, x_scale, x_asym, fit_type)
        ReDim Preserve x_hist(1 To n_bin, 1 To 3)
        For i = 1 To n_bin
            x_hist(i, 3) = Prob_x(x_hist(i, 1), x_loc, x_scale, x_asym, fit_type)
        Next i
    End If
    
    Histogram_Create = x_hist
    Erase x_hist
End Function


'Return the log likelihood of a pdf fit to a 1-D real value array x()
'also return the fitting parameters (x_loc, x_scale, x_asym) which are
'GAUSSIAN:  (mean, variance, 0)
'LAPLACE:   (median, absolution deviation, 0)
'ALD:       (m,lambda,kappa)
'AGD:       (xi, omega, alpha)
'CAUCHY:    (x_loc, x_scale)
Function Prob_Fit(x As Variant, x_loc As Double, x_scale As Double, x_asym As Double, Optional fit_type As String = "GAUSSIAN") As Double
Dim i As Long, n As Long
Dim likelihood As Double
    x_loc = 0
    x_scale = 0
    x_asym = 0
    n = UBound(x, 1)
    If VBA.UCase(fit_type) = "GAUSSIAN" Then
    
        For i = 1 To n
            x_loc = x_loc + x(i)
            x_scale = x_scale + x(i) ^ 2
        Next i
        x_loc = x_loc / n
        x_scale = x_scale / n - (x_loc ^ 2)
        likelihood = -(1 + Log(6.28318530717959 * x_scale)) / 2
        
    ElseIf VBA.UCase(fit_type) = "LAPLACE" Then
    
        x_loc = fmedian(x)
        For i = 1 To n
            x_scale = x_scale + Abs(x(i) - x_loc)
        Next i
        x_scale = x_scale / n
        likelihood = -1 - Log(2 * x_scale)
        
    ElseIf VBA.UCase(fit_type) = "ALD" Then
    
        likelihood = ALD_Fit(x, x_loc, x_scale, x_asym)

    ElseIf VBA.UCase(fit_type) = "AGD" Then
    
        likelihood = AGD_Fit(x, x_loc, x_scale, x_asym)
    
    ElseIf VBA.UCase(fit_type) = "CAUCHY" Then
        
        likelihood = Cauchy_Fit(x, x_loc, x_scale)
    
    Else
        Debug.Print "Prob_Fit: " & fit_type & " is not a valid option."
    End If
    Prob_Fit = likelihood
End Function


Function Prob_x(x As Double, x_loc As Double, x_scale As Double, Optional x_asym As Double = 0, Optional prob_type As String = "GAUSSIAN") As Double
Dim p As Double
    If prob_type = "GAUSSIAN" Then
        p = Exp(-((x - x_loc) ^ 2) / (2 * x_scale)) / Sqr(6.28318530717959 * x_scale)
    ElseIf prob_type = "LAPLACE" Then
        p = Exp(-Abs(x - x_loc) / x_scale) / (2 * x_scale)
    ElseIf prob_type = "ALD" Then
        If x < x_loc Then
            p = (x_scale / (x_asym + 1 / x_asym)) * Exp((x_scale / x_asym) * (x - x_loc))
        ElseIf x >= x_loc Then
            p = (x_scale / (x_asym + 1 / x_asym)) * Exp(-(x_scale * x_asym) * (x - x_loc))
        End If
    ElseIf prob_type = "AGD" Then
        p = (x - x_loc) / x_scale
        p = Exp(-(p ^ 2) / 2) * (1 + Application.WorksheetFunction.Erf(x_asym * p / Sqr(2))) / (Sqr(6.28318530717959) * x_scale)
    ElseIf prob_type = "CAUCHY" Then
        p = 1 / ((1 + ((x - x_loc) / x_scale) ^ 2) * x_scale * 3.14159265358979)
    End If
    Prob_x = p
End Function


'Asymetric Laplace Distribution:
'returns the maximum likelihood estimates of (m,lambda, kappa)
'returns the corresponding log likelihood
Function ALD_Fit(x As Variant, m As Double, lambda As Double, kappa As Double, Optional tol As Double = 0.00001) As Double
Dim i As Long, j As Long, k As Long, iterate As Long, n As Long
Dim x_A As Double, x_b As Double, x_c As Double, x_d As Double, f_a As Double, f_b As Double
Dim tmp_x As Double, tmp_y As Double, z As Double, y_sum As Double
Dim y() As Double

    n = UBound(x, 1)
    y = x
    Call Sort_Quick(y, 1, n)
    y_sum = 0
    For i = 1 To n
        y_sum = y_sum + y(i)
    Next i
    
    j = 1
    Do While y(j) = y(j + 1) And j < n
        j = j + 1
    Loop
    k = n
    Do While y(k) = y(k - 1) And k > 1
        k = k - 1
    Loop
    
    'Golden line search to find maximum likelihood
    z = (Sqr(5) + 1) / 2
    x_A = y(j + 1)
    x_b = y(k - 1)
    x_c = x_b - (x_b - x_A) / z
    x_d = x_A + (x_b - x_A) / z
    For iterate = 1 To 5000
        If Abs(x_d - x_c) / (Abs(x_c) / 2) < tol Then Exit For
        tmp_x = ALD_likelihood_calc_m(y, x_c, kappa, lambda, y_sum)
        tmp_y = ALD_likelihood_calc_m(y, x_d, kappa, lambda, y_sum)
        
        If tmp_x > tmp_y Then
            x_b = x_d
        Else
            x_A = x_c
        End If
        x_c = x_b - (x_b - x_A) / z
        x_d = x_A + (x_b - x_A) / z
    Next iterate
        
    m = (x_A + x_b) / 2
    ALD_Fit = ALD_likelihood_calc_m(y, m, kappa, lambda, y_sum)
    Erase y
End Function

Private Function ALD_likelihood_calc_m(x As Variant, x_root As Double, _
    kappa As Double, lambda As Double, x_sum As Double) As Double
Dim i As Long, n As Long, N1 As Long, N2 As Long
Dim x_sum1 As Double, x_sum2 As Double
    n = UBound(x)
    N1 = 0
    x_sum1 = 0
    For i = 1 To n
        If x(i) < x_root Then
            x_sum1 = x_sum1 + x(i)
            N1 = N1 + 1
        Else
            Exit For
        End If
    Next i
    N2 = n - N1
    x_sum2 = x_sum - x_sum1
    x_sum1 = x_sum1 - N1 * x_root
    x_sum2 = x_sum2 - N2 * x_root
    kappa = Sqr(Sqr(-x_sum1 / x_sum2))
    lambda = n / (x_sum2 * kappa - x_sum1 / kappa)
    ALD_likelihood_calc_m = -1 + Log(lambda / (kappa + 1 / kappa))
'    m_grad = (N2 * kappa - N1 / kappa) * lambda / n
End Function


'Asymetric Laplace Distribution: return the probability density at x
Function ALD(x As Double, m As Double, lambda As Double, kappa As Double) As Double
    If x < m Then
        ALD = (lambda / (kappa + 1 / kappa)) * Exp((lambda / kappa) * (x - m))
    ElseIf x >= m Then
        ALD = (lambda / (kappa + 1 / kappa)) * Exp(-(lambda * kappa) * (x - m))
    End If
End Function


'Asymetric Gaussian Distribution:
'returns the maximum likelihood estimates of (xi, omega, alpha)
'returns the corresponding log likelihood
Function AGD_Fit(x As Variant, x_loc As Double, x_scale As Double, x_asym As Double, _
    Optional iter_max As Long = 1000) As Double
Dim i As Long, j As Long, k As Long, n As Long, iterate As Long, backtrack As Long
Dim conv_count As Long, div_count As Long, stall_count As Long, GD_Stall As Long
Dim x_mean As Double, x_var As Double, x_skew As Double
Dim tmp_x As Double, tmp_y As Double, tmp_y1 As Double, wolfe_test As Double
Dim likelihood As Double, likelihood_prev As Double, likelihood_init As Double
Dim pi As Double, tolerance As Double
Dim grad() As Double, Hessian() As Double, chg() As Double
Dim h As Double, z As Double, h_mean As Double, xh_mean As Double, zh_mean As Double
Dim g As Double, dh As Double, dh_mean As Double, g_mean As Double, zg_mean As Double
Dim x2dh_mean As Double, xg_mean As Double
Dim alpha As Double, x_loc_prev As Double, x_scale_prev As Double, x_asym_prev As Double
    tolerance = 0.000001
    pi = 3.14159265358979
    
    '=== Initial estimate of parameters
    n = UBound(x, 1)
    For i = 1 To n
        x_mean = x_mean + x(i)
        x_var = x_var + x(i) ^ 2
        x_skew = x_skew + x(i) ^ 3
    Next i
    x_mean = x_mean / n
    x_var = (x_var / n - x_mean ^ 2)
    x_skew = ((x_skew / n - x_mean ^ 3) - 3 * x_mean * x_var) / (x_var ^ 1.5)
    If Abs(x_skew) > 0.99 Then x_skew = 0
    tmp_x = (x_skew ^ 2) ^ (1 / 3)
    tmp_x = (tmp_x / (tmp_x + 0.568995659411925)) * pi / 2
    x_asym = Sgn(x_skew) * Sqr(tmp_x / (1 - tmp_x))
    x_scale = Sqr(x_var / (1 - (2 / pi) * tmp_x))
    x_loc = x_mean - x_scale * Sgn(x_skew) * tmp_x * Sqr(2 / pi)
    '==========================================
    
    conv_count = 0
    div_count = 0
    stall_count = 0
    GD_Stall = 0
    likelihood_prev = -Exp(70)
    ReDim grad(1 To 3)
    ReDim Hessian(1 To 3, 1 To 3)
    For iterate = 1 To iter_max
        
        x_mean = 0
        x_var = 0
        tmp_y1 = 0
        h_mean = 0
        zh_mean = 0
        xh_mean = 0
        dh_mean = 0
        g_mean = 0
        zg_mean = 0
        xg_mean = 0
        x2dh_mean = 0
        For i = 1 To n
            tmp_x = (x(i) - x_loc) / x_scale
            z = tmp_x * x_asym / Sqr(2)
            tmp_y = Application.WorksheetFunction.ErfC(-z)
            If tmp_y <> 0 Then
                tmp_y1 = tmp_y1 + Log(tmp_y)
                h = Exp(-(z ^ 2)) / tmp_y
            Else
                tmp_y1 = tmp_y1 - (z ^ 2) - Log(-z) - Log(Sqr(pi)) - 0.5 / (z ^ 2)
                h = -Sqr(pi) * z / (1 - 0.5 / (z ^ 2))
            End If
            dh = -2 * h * (z + h / Sqr(pi))
            g = h + z * dh
            x_mean = x_mean + tmp_x
            x_var = x_var + tmp_x ^ 2
            h_mean = h_mean + h
            zh_mean = zh_mean + h * z
            xh_mean = xh_mean + h * tmp_x
            dh_mean = dh_mean + dh
            g_mean = g_mean + g
            zg_mean = zg_mean + g * z
            xg_mean = xg_mean + g * tmp_x
            x2dh_mean = x2dh_mean + (tmp_x ^ 2) * dh
        Next i
        x_mean = x_mean / n
        x_var = x_var / n
        h_mean = h_mean / n
        zh_mean = zh_mean / n
        xh_mean = xh_mean / n
        dh_mean = dh_mean / n
        g_mean = g_mean / n
        zg_mean = zg_mean / n
        xg_mean = xg_mean / n
        x2dh_mean = x2dh_mean / n
        likelihood = -x_var / 2 + tmp_y1 / n - 0.5 * Log(6.28318530717959 * (x_scale ^ 2))
        grad(1) = (x_mean - x_asym * Sqr(2 / pi) * h_mean) / x_scale
        grad(2) = (x_var - 1 - (2 / Sqr(pi)) * zh_mean) / x_scale
        grad(3) = Sqr(2 / pi) * xh_mean
        Hessian(1, 1) = (((x_asym ^ 2) / Sqr(pi)) * dh_mean - 1) / (x_scale ^ 2)
        Hessian(1, 2) = (x_asym * Sqr(2 / pi) * g_mean - 2 * x_mean) / (x_scale ^ 2)
        Hessian(1, 3) = -Sqr(2 / pi) * g_mean / x_scale
        Hessian(2, 2) = (1 - 3 * x_var + (2 / Sqr(pi)) * (zh_mean + zg_mean)) / (x_scale ^ 2)
        Hessian(2, 3) = -Sqr(2 / pi) * xg_mean / x_scale
        Hessian(3, 3) = x2dh_mean / Sqr(pi)
        Hessian(2, 1) = Hessian(1, 2)
        Hessian(3, 1) = Hessian(1, 3)
        Hessian(3, 2) = Hessian(2, 3)
        
        If iterate = 1 Then likelihood_init = likelihood
        If (Abs(grad(1)) + Abs(grad(2)) + Abs(grad(3))) < tolerance Then Exit For
        
        If likelihood >= likelihood_prev Then
            conv_count = conv_count + 1
            If (likelihood - likelihood_prev) < tolerance Then
                stall_count = stall_count + 1
            Else
                stall_count = 0
            End If
        Else
            conv_count = 0
        End If
        
        If stall_count >= 100 Then Exit For

        Hessian = modMath.Matrix_Inverse(Hessian)
        ReDim chg(1 To 3)
        For i = 1 To 3
            For j = 1 To 3
                chg(i) = chg(i) + Hessian(i, j) * grad(j)
            Next j
        Next i
        x_loc_prev = x_loc
        x_scale_prev = x_scale
        x_asym_prev = x_asym
        wolfe_test = 0
        For i = 1 To 3
            wolfe_test = wolfe_test + chg(i) * grad(i)
        Next i
        wolfe_test = -0.5 * wolfe_test
        If wolfe_test < 0 Then
            alpha = 0.0001
            x_loc = x_loc + alpha * grad(1)
            x_scale = x_scale + alpha * grad(2)
            x_asym = x_asym + alpha * grad(3)
        Else
            alpha = 1
            For backtrack = 1 To 10
                x_loc = x_loc_prev - alpha * chg(1)
                x_scale = x_scale_prev - alpha * chg(2)
                x_asym = x_asym_prev - alpha * chg(3)
                tmp_x = AGD_Likelihood(x, x_loc, x_scale, x_asym)
                If tmp_x > (likelihood + alpha * wolfe_test) Then
                    Exit For
                Else
                    alpha = alpha / 2
                End If
            Next backtrack
            If backtrack >= 10 Then  'Perform simple gradient ascent if Newton-Raphson fails
                alpha = 0.0001
                x_loc = x_loc_prev + alpha * grad(1)
                x_scale = x_scale_prev + alpha * grad(2)
                x_asym = x_asym_prev + alpha * grad(3)
            End If
        End If
        
        likelihood_prev = likelihood
    Next iterate
    
    If likelihood < likelihood_init Then Debug.Print "AGD_Fit not converging."
    AGD_Fit = likelihood
End Function


Private Function AGD_Likelihood(x As Variant, x_loc As Double, x_scale As Double, x_asym As Double) As Double
Dim i As Long, j As Long, k As Long, n As Long
Dim x_mean As Double, x_var As Double, pi As Double
Dim tmp_x As Double, tmp_y As Double, z As Double, tmp_y1 As Double
    pi = 3.14159265358979
    n = UBound(x, 1)
    x_mean = 0
    x_var = 0
    tmp_y1 = 0
    For i = 1 To n
        tmp_x = (x(i) - x_loc) / x_scale
        z = tmp_x * x_asym / Sqr(2)
        tmp_y = Application.WorksheetFunction.ErfC(-z)
        If tmp_y > 0 Then
            tmp_y1 = tmp_y1 + Log(tmp_y)
        Else
            tmp_y1 = tmp_y1 - (z ^ 2) - Log(-z) - Log(Sqr(pi)) - 0.5 / (z ^ 2)
        End If
        x_var = x_var + tmp_x ^ 2
    Next i
    AGD_Likelihood = (-x_var / 2 + tmp_y1) / n - 0.5 * Log(6.28318530717959 * (x_scale ^ 2))
End Function

Private Sub Adapt_Learn_rate(w As Double, w_chg As Double, w_grad As Double, gain As Double, learn_rate As Double, momentum As Double)
    If Sgn(w_chg) = Sgn(w_grad) Then
        gain = gain * 1.1
    Else
        gain = gain * 0.9
    End If
    If gain < 0.01 Then gain = 0.01
    If gain > 1000 Then gain = 1000
    w_chg = momentum * w_chg + w_grad * gain * learn_rate
    w = w + w_chg
End Sub


'Cauchy Distribution:
'returns the maximum likelihood estimates of (x_loc , x_scale)
'returns the corresponding log likelihood
'Input: x(1 to N) , 1D data vector
Function Cauchy_Fit(x As Variant, x_loc As Double, x_scale As Double) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim stall_count As Long, backtrack As Long, iter_max As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim likelihood As Double, likelihood_prev As Double
Dim grad() As Double, HG() As Double, H11 As Double, H12 As Double, H22 As Double
Dim wolfe_test As Double, alpha As Double, x_loc_prev As Double, x_scale_prev As Double

iter_max = 1000
n = UBound(x, 1)
grad = modMath.fQuartile(x)
x_loc = grad(2) 'Application.WorksheetFunction.Median(x)
x_scale = (grad(3) - grad(1)) / 2 '(Application.WorksheetFunction.Quartile(x, 3) - Application.WorksheetFunction.Quartile(x, 1)) / 2

ReDim HG(1 To 2)
stall_count = 0
likelihood_prev = -Exp(70)
For iterate = 1 To iter_max
    likelihood = 0
    ReDim grad(1 To 2)
    H11 = 0: H22 = 0: H12 = 0
    For i = 1 To n
        tmp_x = (x(i) - x_loc) / x_scale
        tmp_y = 1 / (1 + tmp_x ^ 2)
        tmp_z = tmp_x * tmp_y
        grad(1) = grad(1) + tmp_z
        grad(2) = grad(2) + tmp_z * tmp_x
        H11 = H11 + 2 * (tmp_z) ^ 2 - tmp_y
        H22 = H22 + (tmp_x * tmp_z) * (tmp_x * tmp_z - 1)
        H12 = H12 + tmp_x * (tmp_z ^ 2)
        likelihood = likelihood + Log(1 + tmp_x ^ 2)
    Next i
    grad(1) = grad(1) * 2 / (n * x_scale)
    grad(2) = grad(2) * 2 / (n * x_scale) - 1 / x_scale
    likelihood = -likelihood / n - Log(x_scale * 3.14159265358979)
    
    If (Abs(grad(1)) + Abs(grad(2))) < 0.0000000001 Then Exit For

    If likelihood >= likelihood_prev Then
        If (likelihood - likelihood_prev) < 0.000001 Then
            stall_count = stall_count + 1
        Else
            stall_count = 0
        End If
    Else
        stall_count = stall_count + 1
    End If
    If stall_count >= 100 Then
        Debug.Print "Cauchy_Fit: terminate of maximum stall_count."
        Exit For
    End If
    
    tmp_x = n * (x_scale ^ 2)
    H11 = H11 * 2 / tmp_x
    H22 = H22 * 4 / tmp_x - grad(2) / x_scale
    H12 = H12 * 4 / tmp_x - grad(1) * 2 / x_scale
    
    'Inverse of Hessian
    tmp_x = H11 * H22 - H12 * H12
    tmp_y = H22
    H22 = H11 / tmp_x
    H11 = tmp_y / tmp_x
    H12 = -H12 / tmp_x
    
    HG(1) = H11 * grad(1) + H12 * grad(2)
    HG(2) = H12 * grad(1) + H22 * grad(2)
    wolfe_test = -0.5 * (HG(1) * grad(1) + HG(2) * grad(2))
    If wolfe_test < 0 Then  'Raphson not applicable, perform gradient ascend
        alpha = 0.001
        x_loc = x_loc + alpha * grad(1)
        x_scale = x_scale + alpha * grad(2)
    Else
        alpha = 1
        x_loc_prev = x_loc
        x_scale_prev = x_scale
        For backtrack = 1 To 15
            x_loc = x_loc_prev - alpha * HG(1)
            x_scale = x_scale_prev - alpha * HG(2)
            tmp_x = Cauchy_Likelihood(x, x_loc, x_scale)
            If (tmp_x - likelihood) > (alpha * wolfe_test) Then
                Exit For
            Else
                alpha = alpha / 2
            End If
        Next backtrack
        If backtrack >= 15 Then 'Raphson not applicable, perform gradient ascend
            alpha = 0.001
            x_loc = x_loc_prev + alpha * grad(1)
            x_scale = x_scale_prev + alpha * grad(2)
        End If
    End If

    likelihood_prev = likelihood
Next iterate
Cauchy_Fit = likelihood
Erase grad, HG
End Function


Private Function Cauchy_Likelihood(x As Variant, x_loc As Double, x_scale As Double) As Double
Dim i As Long, n As Long
Dim tmp_x As Double
    n = UBound(x, 1)
    tmp_x = 0
    For i = 1 To n
        tmp_x = tmp_x + Log(1 + ((x(i) - x_loc) / x_scale) ^ 2)
    Next i
    Cauchy_Likelihood = -tmp_x / n - Log(x_scale * 3.14159265358979)
End Function


'=== Find convex hull of 2D data using Graham Scan
'Input: x(1 to N,1 to 2) is the set of N points with coordinate (x,y) store in the first dimension
'Output: hull() is an integer index of elements that belong to the convex hull
Function ConvexHull(x As Variant, Optional x_hull As Variant) As Long()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_raw As Long
Dim tmp_x As Double, tmp_y As Double
Dim x_polar() As Double
Dim intArray() As Long, sort_index() As Long, hull() As Long

    n_raw = UBound(x, 1)
    ReDim hull(0 To 0)
    
    '=== Identify Lowest point
    tmp_x = Exp(70)
    tmp_y = Exp(70)
    For i = 1 To n_raw
        If x(i, 2) < tmp_y Or (x(i, 2) = tmp_y And x(i, 1) < tmp_x) Then
            tmp_y = x(i, 2)
            tmp_x = x(i, 1)
            j = i
        End If
    Next i
    Call Push(hull, j)
    '===========================
    
    '=== Sort remaining points by polar angle rel. to first point
    ReDim x_polar(1 To n_raw - 1)
    ReDim x_dist(1 To n_raw - 1)
    ReDim intArray(1 To n_raw - 1)
    k = 0
    tmp_x = x(j, 1)
    tmp_y = x(j, 2)
    For i = 1 To n_raw
        If i <> j Then
            k = k + 1
            intArray(k) = i
            x_dist(k) = Sqr((x(i, 1) - tmp_x) ^ 2 + (x(i, 2) - tmp_y) ^ 2)
            x_polar(k) = (tmp_x - x(i, 1)) / x_dist(k)
        End If
    Next i
    
    Call modMath.Sort_Quick_A(x_polar, 1, n_raw - 1, sort_index)
    Erase x_polar, x_dist
    '===========================
    
    '=== Add next two points to the hull
    Call Push(hull, intArray(sort_index(1)))
    Call Push(hull, intArray(sort_index(2)))
    '========================================
    
    k = UBound(hull)
    For i = 3 To n_raw
        
        If i = n_raw Then
            j = hull(1)
        Else
            j = intArray(sort_index(i))
        End If
        
        '=== Remove top stack if current segment is not counter-clockwise
        m = hull(k - 1) 'next to top of stack
        n = hull(k) 'top of stack
        Do While CCW(x(m, 1), x(m, 2), x(n, 1), x(n, 2), x(j, 1), x(j, 2)) <= 0
            Call Pop(hull)
            k = UBound(hull)
            m = hull(k - 1)
            n = hull(k)
        Loop
        '========================================
        
        Call Push(hull, j)
        k = UBound(hull)
        
    Next i
    
    ConvexHull = hull
    If IsMissing(x_hull) = False Then
        ReDim x_hull(1 To UBound(hull, 1), 1 To 2)
        For i = 1 To UBound(hull)
            x_hull(i, 1) = x(hull(i), 1)
            x_hull(i, 2) = x(hull(i), 2)
        Next i
    End If
    Erase sort_index, intArray
End Function


Private Sub Push(s() As Long, i As Long)
    Dim n As Long
    n = UBound(s, 1) + 1
    ReDim Preserve s(0 To n)
    s(n) = i
End Sub

Private Sub Pop(s() As Long, Optional i As Long)
    Dim n As Long
    n = UBound(s, 1)
    i = s(n)
    ReDim Preserve s(0 To n - 1)
End Sub

'CCW>0, counter-clockwise
'CCW<0,clockwise
'CCW=0, collinear
Private Function CCW(x1 As Variant, y1 As Variant, x2 As Variant, y2 As Variant, x3 As Variant, y3 As Variant) As Variant
CCW = (x2 - x1) * (y3 - y1) - (y2 - y1) * (x3 - x1)
End Function




'========================================
'Matrix mathematics
'========================================

'Dot Product of A & B, set AT/BT=1 if A/B needs to be transpose
'if one of them is a 1D-vector, only need to set the tranpose of the matrix
Function M_Dot(A As Variant, B As Variant, Optional AT As Long = 0, Optional BT As Long = 0) As Double()
Dim A_dim As Long, B_dim As Long
Dim c() As Double
    A_dim = getDimension(A)
    B_dim = getDimension(B)
    If A_dim = 2 And B_dim = 2 Then
        M_Dot = MM_dot(A, B, AT, BT)
    ElseIf A_dim = 2 And B_dim = 1 Then
        M_Dot = MV_dot(A, B, AT)
    ElseIf A_dim = 1 And B_dim = 2 Then
        M_Dot = VM_dot(A, B, BT)
    End If
End Function

'Compute the sum of two matrix/vector: A+B
'Set AT/BT=1 if A/B needs to be transpose
'if one of them is a 1D-vector, only need to set the tranpose of the matrix
Function M_Add(A As Variant, B As Variant, Optional AT As Long = 0, Optional BT As Long = 0) As Double()
Dim A_dim As Long, i As Long, j As Long
Dim c() As Double
    A_dim = getDimension(A)
    If A_dim = 1 Then
        ReDim c(LBound(A, 1) To UBound(A, 1))
        For i = LBound(A, 1) To UBound(A, 1)
            c(i) = A(i) + B(i)
        Next i
    ElseIf A_dim = 2 Then
        If AT = 0 And BT = 0 Then
            ReDim c(LBound(A, 1) To UBound(A, 1), LBound(A, 2) To UBound(A, 2))
            For i = LBound(A, 1) To UBound(A, 1)
                For j = LBound(A, 2) To UBound(A, 2)
                    c(i, j) = A(i, j) + B(i, j)
                Next j
            Next i
        ElseIf AT = 0 And BT = 1 Then
            ReDim c(LBound(A, 1) To UBound(A, 1), LBound(A, 2) To UBound(A, 2))
            For i = LBound(A, 1) To UBound(A, 1)
                For j = LBound(A, 2) To UBound(A, 2)
                    c(i, j) = A(i, j) + B(j, i)
                Next j
            Next i
        ElseIf AT = 1 And BT = 0 Then
            ReDim c(LBound(A, 2) To UBound(A, 2), LBound(A, 1) To UBound(A, 1))
            For i = LBound(A, 2) To UBound(A, 2)
                For j = LBound(A, 1) To UBound(A, 1)
                    c(i, j) = A(j, i) + B(i, j)
                Next j
            Next i
        ElseIf AT = 1 And BT = 1 Then
            ReDim c(LBound(A, 2) To UBound(A, 2), LBound(A, 1) To UBound(A, 1))
            For i = LBound(A, 2) To UBound(A, 2)
                For j = LBound(A, 1) To UBound(A, 1)
                    c(i, j) = A(j, i) + B(j, i)
                Next j
            Next i
        End If
    End If
    M_Add = c
    Erase c
End Function


'Element-wise multplication by scalar x
Function M_scalar_dot(A As Variant, x As Variant) As Double()
Dim i As Long, j As Long, A_dim As Long
Dim c() As Double
    A_dim = getDimension(A)
    If A_dim = 1 Then
        ReDim c(1 To UBound(A, 1))
        For i = 1 To UBound(A, 1)
            c(i) = A(i) * x
        Next i
    ElseIf A_dim = 2 Then
        ReDim c(1 To UBound(A, 1), 1 To UBound(A, 2))
        For i = 1 To UBound(A, 1)
            For j = 1 To UBound(A, 2)
                c(i, j) = A(i, j) * x
            Next j
        Next i
    End If
    M_scalar_dot = c
    Erase c
End Function


'Element-wise addition by scalar x
Function M_scalar_Add(A As Variant, x As Variant) As Double()
Dim i As Long, j As Long, A_dim As Long
Dim c() As Double
    A_dim = getDimension(A)
    If A_dim = 1 Then
        ReDim c(1 To UBound(A, 1))
        For i = 1 To UBound(A, 1)
            c(i) = A(i) + x
        Next i
    ElseIf A_dim = 2 Then
        ReDim c(1 To UBound(A, 1), 1 To UBound(A, 2))
        For i = 1 To UBound(A, 1)
            For j = 1 To UBound(A, 2)
                c(i, j) = A(i, j) + x
            Next j
        Next i
    End If
    M_scalar_Add = c
    Erase c
End Function

Private Function MM_dot(A As Variant, B As Variant, Optional AT As Long = 0, Optional BT As Long = 0) As Double()
Dim i As Long, j As Long, k As Long
Dim m As Long, n As Long, p As Long, q As Long
Dim tmp_x As Double
Dim c() As Double

m = UBound(A, 1)
n = UBound(A, 2)
p = UBound(B, 1)
q = UBound(B, 2)

If AT = 0 And BT = 0 Then

    ReDim c(1 To m, 1 To q)
    For i = 1 To m
        For j = 1 To q
            tmp_x = 0
            For k = 1 To n
                tmp_x = tmp_x + A(i, k) * B(k, j)
            Next k
            c(i, j) = tmp_x
        Next j
    Next i

ElseIf AT = 1 And BT = 0 Then

    ReDim c(1 To n, 1 To q)
    For i = 1 To n
        For j = 1 To q
            tmp_x = 0
            For k = 1 To m
                tmp_x = tmp_x + A(k, i) * B(k, j)
            Next k
            c(i, j) = tmp_x
        Next j
    Next i

ElseIf AT = 0 And BT = 1 Then

    ReDim c(1 To m, 1 To p)
    For i = 1 To m
        For j = 1 To p
            tmp_x = 0
            For k = 1 To n
                tmp_x = tmp_x + A(i, k) * B(j, k)
            Next k
            c(i, j) = tmp_x
        Next j
    Next i
    
ElseIf AT = 1 And BT = 1 Then

    ReDim c(1 To n, 1 To p)
    For i = 1 To n
        For j = 1 To p
            tmp_x = 0
            For k = 1 To m
                tmp_x = tmp_x + A(k, i) * B(j, k)
            Next k
            c(i, j) = tmp_x
        Next j
    Next i

End If

MM_dot = c
Erase c
End Function

Private Function MV_dot(A As Variant, B As Variant, Optional AT As Long = 0) As Double()
Dim i As Long, j As Long, m As Long, n As Long
Dim c() As Double
m = UBound(A, 1)
n = UBound(A, 2)
If AT = 0 Then
    ReDim c(1 To m)
    For i = 1 To m
        For j = 1 To n
            c(i) = c(i) + A(i, j) * B(j)
        Next j
    Next i
ElseIf AT = 1 Then
    ReDim c(1 To n)
    For i = 1 To n
        For j = 1 To m
            c(i) = c(i) + A(j, i) * B(j)
        Next j
    Next i
End If
MV_dot = c
Erase c
End Function

Private Function VM_dot(A As Variant, B As Variant, Optional BT As Long = 0) As Double()
If BT = 0 Then
    VM_dot = MV_dot(B, A, 1)
ElseIf BT = 1 Then
    VM_dot = MV_dot(B, A, 0)
End If
End Function

Function VV_dot(A As Variant, B As Variant) As Double
Dim i As Long
    VV_dot = 0
    For i = 1 To UBound(A, 1)
        VV_dot = VV_dot + A(i) * B(i)
    Next i
End Function

Function getDimension(A As Variant) As Long
    Dim i As Long, j As Long
    i = 0
    On Error GoTo getDimension_Err:
    Do While True:
        i = i + 1
        j = UBound(A, i)
    Loop
getDimension_Err:
    getDimension = i - 1
End Function


'=== Add x to every diagonal element of A()
Function Diag_Add(A() As Double, x As Double) As Double()
Dim i As Long
Dim B() As Double
    B = A
    For i = 1 To UBound(A, 1)
        B(i, i) = B(i, i) + x
    Next i
    Diag_Add = B
    Erase B
End Function


'=== If A() is a 1D-vector, promote A() to a diagonal matrix
'=== If A() is a 2D-matrix, output diagonal elements of A() as a vector
Function mDiag(A As Variant) As Double()
Dim i As Long, k As Long, n As Long, m As Long
Dim B() As Double
    k = getDimension(A)
    m = LBound(A, 1)
    n = UBound(A, 1)
    If k = 1 Then
        ReDim B(m To n, m To n)
        For i = m To n
            B(i, i) = A(i)
        Next i
    ElseIf k = 2 Then
        ReDim B(m To n)
        For i = m To n
            B(i) = A(i, i)
        Next i
    End If
    mDiag = B
    Erase B
End Function




'========================================
'Matrix Factorization
'========================================

'=== Cholesky Decomposition, returns lower triangular matrix
Function Cholesky(A As Variant) As Double()
Dim i As Long, j As Long, k As Long
Dim n As Long
Dim tmp_x As Double, tmp_y As Double
Dim L() As Double
    n = UBound(A, 1)
    ReDim L(1 To n, 1 To n)
    If n = 1 Then
        L(1, 1) = Sqr(A(1, 1))
    ElseIf n = 2 Then
        L(1, 1) = Sqr(A(1, 1))
        L(2, 1) = A(2, 1) / L(1, 1)
        L(2, 2) = Sqr(A(2, 2) - L(2, 1) * L(2, 1))
    Else
        For i = 1 To n
            For j = 1 To i
                tmp_x = 0
                For k = 1 To j - 1
                    tmp_x = tmp_x + L(i, k) * L(j, k)
                Next k
                If i = j Then
                    If (A(i, i) - tmp_x) < 0 Then
                        Debug.Print "Cholesky: Failed: Matrix is not positive definite."
                        Cholesky = L
                        Erase L
                        Exit Function
                    End If
                    L(i, j) = Sqr(A(i, i) - tmp_x)
                Else
                    L(i, j) = (A(i, j) - tmp_x) / L(j, j)
                End If
            Next j
        Next i
    End If
    Cholesky = L
    Erase L
End Function


'=== LDL-Factorization: A=LDL^T
'Input:     A(), symmetric matrix A()
'Output:    L() is a lower triangular matrix, d() is a vector of the diagonals of D.
Sub LDL_Decompose(A As Variant, L() As Double, d() As Double)
Dim i As Long, j As Long, k As Long, n As Long
    n = UBound(A, 1)
    ReDim d(1 To n)
    ReDim L(1 To n, 1 To n)
    For j = 1 To n
        L(j, j) = 1
        d(j) = A(j, j)
        For k = 1 To j - 1
            d(j) = d(j) - (L(j, k) ^ 2) * d(k)
        Next k
        For i = j + 1 To n
            L(i, j) = A(i, j)
            For k = 1 To j - 1
                L(i, j) = L(i, j) - L(i, k) * L(j, k) * d(k)
            Next k
            L(i, j) = L(i, j) / d(j)
        Next i
    Next j
End Sub


'=== LDL-Factorization with pivoting: PAP^T=LDL^T
'Input:  A(), symmetric matrix A()
'Output: L() is a lower triangular matrix
'        d() is is a block diagonal matrix
'        p() is a vector s.t. P(p(j),j)=1
'        pivot() is a vector s.t. pivot(k)=1 or 2 if block k is 1x1 or 2x2
Sub LDLP_Decompose(A As Variant, L() As Double, d() As Double, p() As Long, pivot() As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, alpha As Double
Dim maxA As Double, maxAd As Double, q As Long, r As Long, maxd As Long
Dim B As Variant, cVec() As Double, EVec() As Double, tmp_vec() As Double

    alpha = 0.6404
    n = UBound(A, 1)
    B = A
    ReDim d(1 To n, 1 To n)
    ReDim p(1 To n)
    ReDim pivot(1 To n)
    ReDim L(1 To n, 1 To n)
    For i = 1 To n
        L(i, i) = 1
        p(i) = i
    Next i

    k = 1
    Do While k <= n
        
        maxA = 0: q = k: r = k
        maxAd = 0: maxd = k
        For i = k To n
            If Abs(B(i, i)) > maxAd Then
                maxAd = Abs(B(i, i))
                maxd = i
            End If
            For j = k To i
                If Abs(B(i, j)) > maxA Then
                    maxA = Abs(B(i, j))
                    r = i
                    q = j
                End If
            Next j
        Next i
        
        If maxAd >= (alpha * maxA) And maxd <> k Then
            
            Call swap_i(p(k), p(maxd))
            Call swap_x(B(k, k), B(maxd, maxd))
            For i = maxd + 1 To n
                Call swap_x(B(i, k), B(i, maxd))
            Next i
            For i = k + 1 To maxd - 1
                Call swap_x(B(i, k), B(maxd, i))
            Next i
            For j = 1 To k - 1
                Call swap_x(L(k, j), L(maxd, j))
            Next j
            
            For i = k + 1 To n
                L(i, k) = B(i, k) / B(k, k)
                For j = k + 1 To n
                    B(i, j) = B(i, j) - B(i, k) * B(j, k) / B(k, k)
                Next j
            Next i
            d(k, k) = B(k, k)
            pivot(k) = 1
            k = k + 1
            
        ElseIf maxAd < (alpha * maxA) And (q <> k Or r <> k) And k < (n - 1) Then

            Call swap_i(p(k), p(q))
            Call swap_i(p(k + 1), p(r))
            Call swap_x(B(k, k), B(q, q))
            For i = q + 1 To n
                Call swap_x(B(i, k), B(i, q))
            Next i
            For i = k + 1 To q - 1
                Call swap_x(B(i, k), B(q, i))
            Next i

            Call swap_x(B(k + 1, k + 1), B(r, r))
            For i = r + 1 To n
                Call swap_x(B(i, k + 1), B(i, r))
            Next i
            Call swap_x(B(k + 1, k), B(r, k))
            For i = k + 2 To r - 1
                Call swap_x(B(i, k + 1), B(r, i))
            Next i

            For j = 1 To k - 1
                Call swap_x(L(k, j), L(q, j))
            Next j
            For j = 1 To k - 1
                Call swap_x(L(k + 1, j), L(r, j))
            Next j

            ReDim cVec(1 To n - k - 1, 1 To 2)
            ReDim EVec(1 To 2, 1 To 2)
            tmp_x = B(k, k) * B(k + 1, k + 1) - B(k + 1, k) ^ 2
            EVec(1, 1) = B(k + 1, k + 1) / tmp_x
            EVec(2, 2) = B(k, k) / tmp_x
            EVec(1, 2) = -B(k + 1, k) / tmp_x
            EVec(2, 1) = -B(k + 1, k) / tmp_x
            For i = k + 2 To n
                cVec(i - k - 1, 1) = B(i, k)
                cVec(i - k - 1, 2) = B(i, k + 1)
            Next i

            tmp_vec = modMath.M_Dot(cVec, EVec)
            For i = k + 2 To n
                L(i, k) = tmp_vec(i - k - 1, 1)
                L(i, k + 1) = tmp_vec(i - k - 1, 2)
            Next i
            tmp_vec = modMath.M_Dot(tmp_vec, cVec, 0, 1)
            For i = k + 2 To n
                For j = k + 2 To n
                    B(i, j) = B(i, j) - tmp_vec(i - k - 1, j - k - 1)
                Next j
            Next i
            
            d(k, k) = B(k, k)
            d(k + 1, k + 1) = B(k + 1, k + 1)
            d(k, k + 1) = B(k + 1, k)
            d(k + 1, k) = B(k + 1, k)
            pivot(k) = 2
            k = k + 2
            
        Else
            For i = k + 1 To n
                L(i, k) = B(i, k) / B(k, k)
                For j = k + 1 To n
                    B(i, j) = B(i, j) - B(i, k) * B(j, k) / B(k, k)
                Next j
            Next i
            d(k, k) = B(k, k)
            pivot(k) = 1
            k = k + 1
        End If
        
    Loop
    Erase B, EVec, cVec
End Sub

Private Sub swap_i(i As Long, j As Long)
Dim k As Long
    k = i
    i = j
    j = k
End Sub

Private Sub swap_x(x As Variant, y As Variant)
Dim z As Variant
    z = x
    x = y
    y = z
End Sub


'=== LU Decompostion with partial pivoting
'Input:  A(), NxN square matrix
'Output: A() is changed, it contains both matrices L-E and U as A=(L-E)+U such that P*A=L*U.
'        The permutation matrix is not stored as a matrix, but in an integer vector P of size N+1
'        containing column indexes where the permutation matrix has "1". The last element P[N]=S+N,
'        where S is the number of row exchanges needed for determinant computation, det(P)=(-1)^S
Sub LUPDecompose(A() As Double, p() As Long)
Dim i As Long, j As Long, k As Long, n As Long
Dim imax As Long
Dim maxA As Double, absA As Double, tmp_x As Double

n = UBound(A, 1)
ReDim p(1 To n + 1)

For i = 1 To n + 1
    p(i) = i
Next i

For i = 1 To n

    maxA = 0
    imax = i
    For k = i To n
        absA = Abs(A(k, i))
        If absA > maxA Then
            maxA = absA
            imax = k
        End If
    Next k
    If maxA < 0.0000000001 Then
        Debug.Print "LUPDecompose: Fail: Matrix is degenerate."
        Exit Sub
    End If
    If imax <> i Then
        j = p(i)
        p(i) = p(imax)
        p(imax) = j
        
        For j = 1 To n
            tmp_x = A(i, j)
            A(i, j) = A(imax, j)
            A(imax, j) = tmp_x
        Next j
        
        p(n + 1) = p(n + 1) + 1
    End If
    
    For j = i + 1 To n
        A(j, i) = A(j, i) / A(i, i)
        For k = i + 1 To n
            A(j, k) = A(j, k) - A(j, i) * A(i, k)
        Next k
    Next j
    
Next i

End Sub


'=== QR-Factorization
'Input: A() mxn matrix with m<=n
'Output: Q() mxm orthonormal matrix and R() mxn upper triangular matrix
Sub QR_Factor(A() As Double, q() As Double, r() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long
Dim normx As Double, s_tmp As Double, u1 As Double, tau As Double, tmp_x As Double
Dim w() As Double, r_tmp() As Double, q_tmp() As Double
    m = UBound(A, 1)
    n = UBound(A, 2)
    If m > n Then
        Debug.Print "QR_Factor: Failed: number of rows > number of columns."
        Exit Sub
    End If
    ReDim q(1 To m, 1 To m)
    For i = 1 To m
        q(i, i) = 1
    Next i
    r = A
    For j = 1 To n
        normx = 0
        ReDim w(1 To m - j + 1)
        For k = j To m
            w(k - j + 1) = r(k, j)
            normx = normx + r(k, j) ^ 2
        Next k
        normx = Sqr(normx)
        s_tmp = -Sgn(r(j, j))
        If s_tmp = 0 Then s_tmp = -1
        u1 = r(j, j) - s_tmp * normx
        For k = 1 To m - j + 1
            w(k) = w(k) / u1
        Next k
        w(1) = 1
        tau = -s_tmp * u1 / normx
        ReDim r_tmp(j To m, 1 To n)
        ReDim q_tmp(1 To m, j To m)
        For i = 1 To n
            tmp_x = 0
            For p = j To m
                tmp_x = tmp_x + w(p - j + 1) * r(p, i)
            Next p
            For k = j To m
                r_tmp(k, i) = tau * w(k - j + 1) * tmp_x
            Next k
        Next i
        For i = 1 To m
            tmp_x = 0
            For p = j To m
                tmp_x = tmp_x + q(i, p) * w(p - j + 1)
            Next p
            For k = j To m
                q_tmp(i, k) = tmp_x * tau * w(k - j + 1)
            Next k
        Next i
        For k = j To m
            For i = 1 To n
                r(k, i) = r(k, i) - r_tmp(k, i)
            Next i
            For i = 1 To m
                q(i, k) = q(i, k) - q_tmp(i, k)
            Next i
        Next k
    Next j
    Erase r_tmp, q_tmp, w
End Sub


'=== Economical SVD of matrix A()
'Input: A() mxn matrix with m<=n
'Output: unitary matrix u() and v() (not v^T), vector s() that holds the singluar lvalue
Sub Matrix_SVD(A As Variant, u() As Double, s() As Double, v() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_vec() As Double, max_err As Double
Dim sort_idx() As Long
    m = UBound(A, 1)
    n = UBound(A, 2)
    u = A
    Call SVDCMP(u, s, v)

    'Sort singular values s.t. large absolute value comes first
    Call Sort_Quick_A(s, 1, n, sort_idx)
    tmp_vec = s
    For i = 1 To n
        s(i) = tmp_vec(n - i + 1)
    Next i
    tmp_vec = u
    For i = 1 To n
        j = sort_idx(n - i + 1)
        For k = 1 To UBound(u, 1)
            u(k, i) = tmp_vec(k, j)
        Next k
    Next i
    tmp_vec = v
    For i = 1 To n
        j = sort_idx(n - i + 1)
        For k = 1 To n
            v(k, i) = tmp_vec(k, j)
        Next k
    Next i
    Erase tmp_vec, sort_idx
    
    'remove zeroes columns/row
    If m > n Then
        ReDim Preserve s(1 To n)
        ReDim Preserve u(1 To m, 1 To n)
    ElseIf m < n Then
        ReDim Preserve s(1 To m)
        ReDim Preserve u(1 To m, 1 To m)
        ReDim Preserve v(1 To n, 1 To m)
    End If
End Sub


'"Numerical Recipes in Fortran 77: The Art of Scientific Computing", Section 2.6, Page 59)
'http://www.aip.de/groups/soe/local/numres/bookfpdf/f2-6.pdf
'Given a matrix a(1:m,1:n), with physical dimensions mp by np, this routine computes its
'singular value decomposition, A = U · W · V T . The matrix U replaces a on output. The
'diagonal matrix of singular values W is output as a vector w(1:n). The matrix V (not the
'transpose V T ) is output as v(1:n,1:n).
Private Sub SVDCMP(A() As Double, w() As Double, v() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, mn_min As Long
Dim its As Long, jj As Long, L As Long, nm As Long, nmax As Long
Dim anorm As Double, c As Double, f As Double, g As Double, h As Double, s As Double, s_scale As Double
Dim x As Double, y As Double, z As Double, rv1() As Double

nmax = 500
ReDim rv1(1 To nmax)
m = UBound(A, 1)
n = UBound(A, 2)
ReDim w(1 To n)
ReDim v(1 To n, 1 To n)

'Householder reduction to bidiagnonal form
g = 0
s_scale = 0
anorm = 0
For i = 1 To n
    L = i + 1
    rv1(i) = s_scale * g
    g = 0
    s = 0
    s_scale = 0
    If (i <= m) Then
        For k = i To m
            s_scale = s_scale + Abs(A(k, i))
        Next k
        If (s_scale <> 0) Then
            For k = i To m
                A(k, i) = A(k, i) / s_scale
                s = s + A(k, i) * A(k, i)
            Next k
            f = A(i, i)
            If Sgn(f) = 0 Then
                g = Sqr(s)
            Else
                g = -Sqr(s) * Sgn(f)
            End If
            h = f * g - s
            A(i, i) = f - g
            For j = L To n
                s = 0
                For k = i To m
                    s = s + A(k, i) * A(k, j)
                Next k
                f = s / h
                For k = i To m
                    A(k, j) = A(k, j) + f * A(k, i)
                Next k
            Next j
            For k = i To m
                A(k, i) = s_scale * A(k, i)
            Next k
        End If
    End If
    w(i) = s_scale * g
    g = 0
    s = 0
    s_scale = 0
    If ((i <= m) And (i <> n)) Then
        For k = L To n
            s_scale = s_scale + Abs(A(i, k))
        Next k
        If (s_scale <> 0) Then
            For k = L To n
                A(i, k) = A(i, k) / s_scale
                s = s + A(i, k) * A(i, k)
            Next k
            f = A(i, L)
            If Sgn(f) = 0 Then
                g = Sqr(s)
            Else
                g = -Sqr(s) * Sgn(f)
            End If
            h = f * g - s
            A(i, L) = f - g
            For k = L To n
                rv1(k) = A(i, k) / h
            Next k
            For j = L To m
                s = 0
                For k = L To n
                    s = s + A(j, k) * A(i, k)
                Next k
                For k = L To n
                    A(j, k) = A(j, k) + s * rv1(k)
                Next k
            Next j
            For k = L To n
                A(i, k) = s_scale * A(i, k)
            Next k
        End If
    End If
    If (Abs(w(i)) + Abs(rv1(i))) > anorm Then anorm = (Abs(w(i)) + Abs(rv1(i)))
Next i

For i = n To 1 Step -1
    If i < n Then
        If g <> 0 Then
            For j = L To n
                v(j, i) = (A(i, j) / A(i, L)) / g
            Next j
            For j = L To n
                s = 0
                For k = L To n
                    s = s + A(i, k) * v(k, j)
                Next k
                For k = L To n
                    v(k, j) = v(k, j) + s * v(k, i)
                Next k
            Next j
        End If
        For j = L To n
            v(i, j) = 0
            v(j, i) = 0
        Next j
    End If
    v(i, i) = 1
    g = rv1(i)
    L = i
Next i

mn_min = m
If n < m Then mn_min = n
For i = mn_min To 1 Step -1
    L = i + 1
    g = w(i)
    For j = L To n
        A(i, j) = 0
    Next j
    If g <> 0 Then
        g = 1 / g
        For j = L To n
            s = 0
            For k = L To m
                s = s + A(k, i) * A(k, j)
            Next k
            f = (s / A(i, i)) * g
            For k = i To m
                A(k, j) = A(k, j) + f * A(k, i)
            Next k
        Next j
        For j = i To m
            A(j, i) = A(j, i) * g
        Next j
    Else
        For j = i To m
            A(j, i) = 0
        Next j
    End If
    A(i, i) = A(i, i) + 1
Next i

'Diagonalization of the bidiagonal form
For k = n To 1 Step -1
    For its = 1 To 100
        For L = k To 1 Step -1
            nm = L - 1
            If (Abs(rv1(L)) + anorm) = anorm Then GoTo svd2
            If (Abs(w(nm)) + anorm) = anorm Then GoTo svd1
        Next L
svd1:
        c = 0
        s = 1
        For i = L To k
            f = s * rv1(i)
            rv1(i) = c * rv1(i)
            If ((Abs(f) + anorm) = anorm) Then GoTo svd2
            g = w(i)
            h = pythag(f, g)
            w(i) = h
            h = 1# / h
            c = g * h
            s = -f * h
            For j = 1 To m
                y = A(j, nm)
                z = A(j, i)
                A(j, nm) = y * c + z * s
                A(j, i) = -y * s + z * c
            Next j
        Next i
svd2:
        z = w(k)
        If L = k Then
            If z < 0 Then
                w(k) = -z
                For j = 1 To n
                    v(j, k) = -v(j, k)
                Next j
            End If
            GoTo svd3
        End If
        If its = 100 Then
            Debug.Print "no convergence in svdcmp."
            Exit Sub
        End If
        x = w(L)
        nm = k - 1
        y = w(nm)
        g = rv1(nm)
        h = rv1(k)
        f = ((y - z) * (y + z) + (g - h) * (g + h)) / (2 * h * y)
        g = pythag(f, 1)
        f = ((x - z) * (x + z) + h * ((y / (f + g * Sgn(f))) - h)) / x
        c = 1
        s = 1
        For j = L To nm
            i = j + 1
            g = rv1(i)
            y = w(i)
            h = s * g
            g = c * g
            z = pythag(f, h)
            rv1(j) = z
            c = f / z
            s = h / z
            f = (x * c) + (g * s)
            g = -(x * s) + (g * c)
            h = y * s
            y = y * c
            For jj = 1 To n
                x = v(jj, j)
                z = v(jj, i)
                v(jj, j) = (x * c) + (z * s)
                v(jj, i) = -(x * s) + (z * c)
            Next jj
            z = pythag(f, h)
            w(j) = z
            If z <> 0 Then
                z = 1 / z
                c = f * z
                s = h * z
            End If
            f = (c * g) + (s * y)
            x = -(s * g) + (c * y)
            For jj = 1 To m
                y = A(jj, j)
                z = A(jj, i)
                A(jj, j) = (y * c) + (z * s)
                A(jj, i) = -(y * s) + (z * c)
            Next jj
        Next j
        rv1(L) = 0
        rv1(k) = f
        w(k) = x
    Next its
svd3:
Next k
End Sub


Private Function pythag(A As Double, B As Double) As Double
Dim absA As Double, absb As Double
    absA = Abs(A)
    absb = Abs(B)
    If absA > absb Then
        pythag = absA * Sqr(1 + (absb / absA) ^ 2)
    Else
        If absb = 0 Then
            pythag = 0
        Else
            pythag = absb * Sqr(1 + (absA / absb) ^ 2)
        End If
    End If
End Function



'==================================================
'Matrix Inversion
'==================================================

'For small matrix, calculate inverse directly
Private Function Matrix_Inverse_small(A As Variant) As Double()
Dim i As Long, j As Long, m As Long, n As Long
Dim d As Double
Dim AI() As Double
    n = UBound(A)
    ReDim AI(1 To n, 1 To n)
    If n = 1 Then
        AI(1, 1) = 1# / A(1, 1)
        Matrix_Inverse_small = AI
    ElseIf n = 2 Then
        d = A(1, 1) * A(2, 2) - A(1, 2) * A(2, 1)
        If d = 0 Then
            Debug.Print "Matrix_Inverse: Error: Matrix is Singular!"
            Exit Function
        Else
            AI(1, 1) = A(2, 2) / d
            AI(2, 2) = A(1, 1) / d
            AI(1, 2) = -A(1, 2) / d
            AI(2, 1) = -A(2, 1) / d
            Matrix_Inverse_small = AI
        End If
    Else
        Debug.Print "Matrix size too large for Matrix_Inverse_small."
    End If
    Erase AI
End Function

'Find inverse of A(nxn) by solving n linear equations with Gaussian elimination
Function Matrix_Inverse_Gauss(A() As Double) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_vec() As Double, tmp_vec2() As Double, AI() As Double
    n = UBound(A, 1)
    If (n = 1 Or n = 2) Then
        Matrix_Inverse_Gauss = Matrix_Inverse_small(A)
        Exit Function
    End If
    ReDim AI(1 To n, 1 To n)
    For j = 1 To n
        ReDim tmp_vec(1 To n)
        tmp_vec(j) = 1
        tmp_vec2 = Solve_Linear_Equations(A, tmp_vec)
        For i = 1 To n
            AI(i, j) = tmp_vec2(i)
        Next i
    Next j
    Matrix_Inverse_Gauss = AI
    Erase AI, tmp_vec, tmp_vec2
End Function


'=== "An efficient and simple algorithm for matrix inversion"
'=== Ahmad Farooq, Khan Hamid
'Has a problem when a diagonal element is zero, can be improved by marking the zero diagonal and re-process
'Input: NxN square matrix x()
'Output: NxN square matrix which is the inverse of x()
Function Matrix_Inverse(x() As Double) As Double()
Dim i As Long, j As Long, m As Long, n As Long, p As Long
Dim d As Double
Dim A() As Double, A_Old() As Double
Dim zero_count As Long, zero_iterate As Long, zero_list()

n = UBound(x)
If (n = 1 Or n = 2) Then
    Matrix_Inverse = Matrix_Inverse_small(x)
    Exit Function
End If

ReDim A(1 To n, 1 To n)
ReDim A_Old(1 To n, 1 To n)
ReDim zero_list(1 To n)

A = x
A_Old = A
d = 1
For p = 1 To n
    If A_Old(p, p) = 0 Then
        zero_count = zero_count + 1
        zero_list(zero_count) = p
    Else
        d = d * A_Old(p, p)
        For j = 1 To n
            If j <> p Then
                A(p, j) = A_Old(p, j) / A_Old(p, p)
                A(j, p) = -A_Old(j, p) / A_Old(p, p)
            End If
        Next j
        For i = 1 To n
            For j = 1 To n
                If i <> p And j <> p Then
                    A(i, j) = A_Old(i, j) + A_Old(p, j) * A(i, p)
                End If
            Next j
        Next i
        A(p, p) = 1# / A_Old(p, p)
        A_Old = A
    End If
Next p

'Re-process zero diagonals
If zero_count > 0 Then
    ReDim Preserve zero_list(1 To zero_count)
    zero_count = 0
    A_Old = A
    d = 1
    For zero_iterate = 1 To UBound(zero_list, 1)
        p = zero_list(zero_iterate)
        If A_Old(p, p) = 0 Then
            zero_count = zero_count + 1
            Debug.Print "Matrix_Inverse: Error: Matrix is Singular!"
            Exit Function
        Else
            d = d * A_Old(p, p)
            For j = 1 To n
                If j <> p Then
                    A(p, j) = A_Old(p, j) / A_Old(p, p)
                    A(j, p) = -A_Old(j, p) / A_Old(p, p)
                End If
            Next j
            For i = 1 To n
                For j = 1 To n
                    If i <> p And j <> p Then
                        A(i, j) = A_Old(i, j) + A_Old(p, j) * A(i, p)
                    End If
                Next j
            Next i
            A(p, p) = 1# / A_Old(p, p)
            A_Old = A
        End If
    Next zero_iterate
End If

Matrix_Inverse = A
Erase A, A_Old
End Function


'=== Find the inverse of a square matrix A() usign iterative method
'=== "A New High-Order Stable Numerical Method for Matrix Inversion"
'=== F. Khaksar Haghani and F. Soleymani, 2014
Function Matrix_Inverse_Iterative(A() As Double, Optional iter_max As Long = 1000, Optional tolerance As Double = 0.0000000001) As Double()
Dim i As Long, j As Long, k As Long, n As Long, iterate As Long
Dim v() As Double, best_v() As Double, w() As Double, c() As Double, kappa() As Double
Dim tmp_x As Double, best_err As Double

    n = UBound(A, 1)
    If (n = 1 Or n = 2) Then
        Matrix_Inverse_Iterative = Matrix_Inverse_small(A)
        Exit Function
    End If
    
    tmp_x = Matrix_Norm(A, 1) * Matrix_Norm(A, 0)
    
    ReDim v(1 To n, 1 To n)
    For i = 1 To n
        For j = 1 To n
            v(i, j) = A(j, i) / tmp_x
        Next j
    Next i
    
    best_err = Exp(70)
    For iterate = 1 To iter_max

        w = M_Dot(A, v)
        c = Diag_Add(M_Dot(w, Diag_Add(M_Dot(w, Diag_Add(M_Dot(w, Diag_Add(w, -8)), 22)), -28)), 17)
        kappa = M_Dot(w, c)
        v = M_Dot(M_Dot(v, c), Diag_Add(M_Dot(kappa, Diag_Add(kappa, -12)), 48))
    
        For i = 1 To n
            For j = 1 To n
                v(i, j) = v(i, j) / 64
            Next j
        Next i
    
        If iterate Mod 5 = 0 Then
            DoEvents
            i = Identity_Chk(M_Dot(A, v), tmp_x, tolerance)
            If i = 1 Then
                best_v = v
                best_err = tmp_x
                Exit For
            End If
            If tmp_x < best_err Then
                best_err = tmp_x
                best_v = v
            End If
        End If
    
    Next iterate
    If best_err > tolerance Or iterate >= iter_max Then Debug.Print "Matrix_Inverse_Iterative: Error=" & best_err
    Matrix_Inverse_Iterative = best_v
    Erase v, best_v, w, c, kappa
End Function


'Find inverse of matrix via Chiolesky factorization,
'corresponding Lower triangular matrix is saved to xL if supplied
Function Matrix_Inverse_Cholesky(A() As Double, Optional xL As Variant) As Double()
Dim i As Long, j As Long, k As Long, n As Long
Dim tmp_x As Double, L() As Double, x() As Double, y() As Double
    n = UBound(A, 1)
    If (n = 1 Or n = 2) Then
        Matrix_Inverse_Cholesky = Matrix_Inverse_small(A)
        Exit Function
    End If
    ReDim y(1 To n, 1 To n)
    L = Cholesky(A)
    For j = 1 To n
        ReDim x(1 To n)
        x(j) = 1
        For i = 1 To n
            tmp_x = 0
            For k = 1 To i - 1
                tmp_x = tmp_x + L(i, k) * x(k)
            Next k
            x(i) = (x(i) - tmp_x) / L(i, i)
        Next i
        
        For i = n To 1 Step -1
            tmp_x = 0
            For k = i + 1 To n
                tmp_x = tmp_x + L(k, i) * x(k)
            Next k
            x(i) = (x(i) - tmp_x) / L(i, i)
        Next i
        For i = 1 To n
            y(i, j) = x(i)
        Next i
    Next j
    Matrix_Inverse_Cholesky = y
    If IsMissing(xL) = False Then xL = L
    Erase x, L
End Function


Function Matrix_Inverse_LDL(A() As Double, Optional isPivot As Boolean = False) As Double()
Dim i As Long, j As Long, k As Long, n As Long
Dim x() As Double
    n = UBound(A, 1)
    If (n = 1 Or n = 2) Then
        Matrix_Inverse_LDL = Matrix_Inverse_small(A)
        Exit Function
    End If
    ReDim x(1 To n, 1 To n)
    For i = 1 To n
        x(i, i) = 1
    Next i
    If isPivot = False Then
        Matrix_Inverse_LDL = modMath.Solve_Linear_LDL(A, x)
    Else
        Matrix_Inverse_LDL = modMath.Solve_Linear_LDLP(A, x)
    End If
    Erase x
End Function


Function Matrix_Inverse_LU(A() As Double) As Double()
Dim i As Long, j As Long, k As Long, n As Long
Dim LU() As Double, p() As Long, IA() As Double
    n = UBound(A, 1)
    If (n = 1 Or n = 2) Then
        Matrix_Inverse_LU = Matrix_Inverse_small(A)
        Exit Function
    End If
    ReDim IA(1 To n, 1 To n)
    LU = A
    Call LUPDecompose(LU, p)
    For j = 1 To n
        For i = 1 To n
            If p(i) = j Then
                IA(i, j) = 1
            Else
                IA(i, j) = 0
            End If
            For k = 1 To i - 1
                IA(i, j) = IA(i, j) - LU(i, k) * IA(k, j)
            Next k
        Next i
        For i = n To 1 Step -1
            For k = i + 1 To n
                IA(i, j) = IA(i, j) - LU(i, k) * IA(k, j)
            Next k
            IA(i, j) = IA(i, j) / LU(i, i)
        Next i
    Next j
    Matrix_Inverse_LU = IA
    Erase LU, IA, p
End Function





'==================================================
'Eigen-Vectors
'==================================================

'=== Jacobi algorithm to find eigen vectors and eigen values
'Input:  S_input is a N X N symmetrix matrix
'Output: Eigen_Vec is a N X N matrix, Eigen_Vec(i,k) is the i-th element of the k-th eigen vector
'Output: Eigen_Val is a N X 1 array, Eigen_Val(k) is the k-th eigen value
Sub Eigen_Jacobi(S_input() As Double, eigen_vec() As Double, eigen_val() As Double, _
            Optional iter_max As Long = 5000)
Dim i As Long, j As Long, k As Long, L As Long, m As Long, n As Long, state As Long
Dim s As Double, c As Double, t As Double, p As Double, y As Double, d As Double, r As Double
Dim n_raw As Long, iterate As Long, swap As Long
Dim temp As Double, temp_i As Long
Dim max_ind() As Long, sort_index() As Long
Dim changed() As Boolean
Dim eigen_vec_old() As Double, eigen_val_old() As Double, s_tmp() As Double

n_raw = UBound(S_input, 1)
s_tmp = S_input
ReDim max_ind(1 To n_raw)
ReDim changed(1 To n_raw)
ReDim eigen_vec(1 To n_raw, 1 To n_raw)
ReDim eigen_val(1 To n_raw)

For i = 1 To n_raw - 1
    For j = i + 1 To n_raw
        If s_tmp(i, j) <> s_tmp(j, i) Then
            Debug.Print "Eigen_Jacobi: Input matrix is not symmetric"
            Exit Sub
        End If
    Next j
Next i

'=== Initialize
For i = 1 To n_raw
    eigen_vec(i, i) = 1
    eigen_val(i) = s_tmp(i, i)
    changed(i) = True
Next i
state = n_raw

'Find index of largest off-diagonal element in each row
ReDim max_ind(1 To n_raw)
For i = 1 To n_raw
    temp = -999999999
    temp_i = 0
    For j = i + 1 To n_raw
        If Abs(s_tmp(i, j)) > temp Then
            temp = Abs(s_tmp(i, j))
            temp_i = j
        End If
    Next j
    max_ind(i) = temp_i
Next i
'================================================

iterate = 0
Do While state <> 0 And iterate < iter_max

    iterate = iterate + 1
    If iterate Mod 500 = 0 Then
        DoEvents
        Application.StatusBar = "Eigen_Jacobi: " & iterate & "," & state
    End If

    m = 1
    For k = 2 To n_raw - 1
        If Abs(s_tmp(k, max_ind(k))) > Abs(s_tmp(m, max_ind(m))) Then m = k
    Next k
    
    k = m
    L = max_ind(m)
    p = s_tmp(k, L)
    y = (eigen_val(L) - eigen_val(k)) / 2
    d = Abs(y) + Sqr(p ^ 2 + y ^ 2)
    r = Sqr(p ^ 2 + d ^ 2)
    c = d / r
    s = p / r
    t = (p ^ 2) / d
    
    If y < 0 Then
        s = -s
        t = -t
    End If
    
    s_tmp(k, L) = 0
    
    y = eigen_val(k)
    eigen_val(k) = y - t
    If changed(k) = True And y = eigen_val(k) Then
        changed(k) = False
        state = state - 1
    ElseIf changed(k) = False And y <> eigen_val(k) Then
        changed(k) = True
        state = state + 1
    End If

    y = eigen_val(L)
    eigen_val(L) = y + t
    If changed(L) = True And y = eigen_val(L) Then
        changed(L) = False
        state = state - 1
    ElseIf changed(L) = False And y <> eigen_val(L) Then
        changed(L) = True
        state = state + 1
    End If

    For i = 1 To k - 1
        Call rotate(s_tmp, i, k, i, L, c, s)
    Next i

    For i = k + 1 To L - 1
        Call rotate(s_tmp, k, i, i, L, c, s)
    Next i
    
    For i = L + 1 To n_raw
        Call rotate(s_tmp, k, i, L, i, c, s)
    Next i

    For i = 1 To n_raw
        Call rotate(eigen_vec, k, i, L, i, c, s)
    Next i

    For i = 1 To n_raw
        temp = -999999999
        temp_i = 0
        For j = i + 1 To n_raw
            If Abs(s_tmp(i, j)) > temp Then
            temp = Abs(s_tmp(i, j))
            temp_i = j
            End If
        Next j
        max_ind(i) = temp_i
    Next i
    
Loop

If (state <> 0 And iterate >= iter_max) Then Debug.Print "Eigen_Jacobi: Max iteration reached."

'=== Sort eigenvalue in descending order
eigen_val_old = eigen_val
eigen_vec_old = eigen_vec
For i = 1 To n_raw
    eigen_val_old(i) = Abs(eigen_val_old(i))
Next i
Call Sort_Bubble_A(eigen_val_old, sort_index)
eigen_val_old = eigen_val
For i = 1 To n_raw
    k = n_raw - i + 1
    eigen_val(i) = eigen_val_old(sort_index(k))
    For j = 1 To n_raw
        eigen_vec(j, i) = eigen_vec_old(sort_index(k), j)
    Next j
Next i
'=================================================

Erase max_ind, changed, eigen_val_old, eigen_vec_old, sort_index, s_tmp
Application.StatusBar = False
End Sub


Private Sub rotate(S_input() As Double, k As Long, L As Long, i As Long, j As Long, c As Double, s As Double)
Dim temp_x As Double, temp_y As Double
    temp_x = S_input(k, L)
    temp_y = S_input(i, j)
    S_input(k, L) = c * temp_x - s * temp_y
    S_input(i, j) = s * temp_x + c * temp_y
End Sub



'=== Find the largest n_Eigen eigenvectors using power iteration
'=== Suitable for large matrix (dimension > 100)
'Input:  A() is a N X N symmetrix matrix
'Output: Eigen_Vec is a n_Eigen X N matrix, Eigen_Vec(i,k) is the i-th element of the k-th eigen vector
'Output: Eigen_Val is a 1D-array of length n_Eigen
Sub Eigen_Power(A() As Double, eigen_vec() As Double, eigen_val() As Double, Optional n_Eigen As Long = 1)
Dim i As Long, j As Long, k As Long, n_dimension As Long
Dim tmp_x As Double, tmp_vec() As Double, B() As Double
    n_dimension = UBound(A, 1)
    If n_Eigen > n_dimension Or n_Eigen = 0 Then
        Debug.Print "Eigen_Power: number of eigenvector (" & n_Eigen & _
            ") cannot be zero or larger than matrix dimension (" & n_dimension & ")."
        Exit Sub
    End If
    ReDim eigen_val(1 To n_Eigen)
    ReDim eigen_vec(1 To n_dimension, 1 To n_Eigen)
    
    tmp_x = Power_Iteration(A, tmp_vec)
    eigen_val(1) = tmp_x
    For i = 1 To n_dimension
        eigen_vec(i, 1) = tmp_vec(i)
    Next i
    If n_Eigen > 1 Then
        B = A
        For k = 2 To n_Eigen
            For i = 1 To n_dimension
                For j = 1 To n_dimension
                    B(i, j) = B(i, j) - eigen_val(k - 1) * eigen_vec(i, k - 1) * eigen_vec(j, k - 1)
                Next j
            Next i
            tmp_x = Power_Iteration(B, tmp_vec)
            eigen_val(k) = tmp_x
            For i = 1 To n_dimension
                eigen_vec(i, k) = tmp_vec(i)
            Next i
        Next k
    End If
    Erase B, tmp_vec
End Sub


'=== Find the largest eigenvector of A() using power iteration
'Input: Square matrix A()
'Output: Largest eigenvalue of A()
'Output: eigenvec(), normalized eigenvector
Function Power_Iteration(A As Variant, eigenvec() As Double, Optional isSparse As Long = 0, Optional n_A As Long = 0, _
            Optional iter_max As Long = 10000, Optional tolerance As Double = 0.000000000000001) As Double
Dim i As Long, j As Long, k As Long, n As Long, iterate As Long, retry As Long
Dim B() As Double, c() As Double
Dim lambda As Double, tmp_x As Double

If isSparse = 0 Then
    n = UBound(A, 1)
    If UBound(A, 2) <> n Then
        Debug.Print "Power_Iteration: Input Matrix is not square"
        Exit Function
    End If
Else
    If n_A = 0 Then
        Debug.Print "Power_Iteration: size of A() must be provided for sparse matrix."
        Exit Function
    End If
    n = n_A
End If

Randomize
tmp_x = 0
ReDim B(1 To n)
For i = 1 To n
    B(i) = Rnd()
    tmp_x = tmp_x + B(i) * B(i)
Next i
tmp_x = Sqr(tmp_x)
For i = 1 To n
    B(i) = B(i) / tmp_x
Next i

For iterate = 1 To iter_max
    lambda = 0: tmp_x = 0
    c = A_dot_B(A, B, isSparse)
    For i = 1 To n
        lambda = lambda + c(i) * c(i)
        tmp_x = tmp_x + B(i) * c(i)
    Next i
    lambda = Sgn(tmp_x) * Sqr(lambda)
    If lambda = 0 Then Exit For
    For i = 1 To n
        c(i) = c(i) / lambda
    Next i
    B = c
    tmp_x = Abs(tmp_x / lambda - 1)
    If tmp_x < tolerance Then Exit For
Next iterate
If iterate >= iter_max Then Debug.Print "Power_Iteration: Max iteration reached. Err=" & tmp_x
Power_Iteration = lambda
eigenvec = B
Erase B, c
End Function

Private Function A_dot_B(A As Variant, B As Variant, Optional A_isSparse As Long = 0) As Double()
Dim i As Long, j As Long, n As Long
Dim c() As Double
    If A_isSparse = 0 Then
        n = UBound(A, 2)
        ReDim c(1 To n)
        For i = 1 To n
            For j = 1 To n
                c(i) = c(i) + A(i, j) * B(j)
            Next j
        Next i
    Else
        ReDim c(1 To UBound(B, 1))
        For j = 1 To UBound(A, 1)
            c(A(j, 1)) = c(A(j, 1)) + A(j, 3) * B(A(j, 2))
        Next j
    End If
    A_dot_B = c
    Erase c
End Function



'================================================
'Other Matrix operations
'================================================

'=== Norm of Matrix A
'for p=infinity put in p=0, for Frobenius put in p=-1
Function Matrix_Norm(A() As Double, Optional p As Long = 1) As Double
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim tmp_x As Double, tmp_max As Double
Dim eigen_vec() As Double, eigen_val() As Double
    m = UBound(A, 1)
    n = UBound(A, 2)
    tmp_max = -Exp(70)
    If p = 1 Then
        For j = 1 To n
            tmp_x = 0
            For i = 1 To m
                tmp_x = tmp_x + Abs(A(i, j))
            Next i
            If tmp_x > tmp_max Then tmp_max = tmp_x
        Next j
    ElseIf p = 0 Then   'Infinity
        For i = 1 To m
            tmp_x = 0
            For j = 1 To n
                tmp_x = tmp_x + Abs(A(i, j))
            Next j
            If tmp_x > tmp_max Then tmp_max = tmp_x
        Next i
    ElseIf p = 2 Then
        Call Eigen_Power(M_Dot(A, A, 1, 0), eigen_vec, eigen_val)
        tmp_max = Sqr(eigen_val(1))
    ElseIf p = -1 Then 'Frobenius
        tmp_x = 0
        For i = 1 To m
            For j = 1 To n
                tmp_x = tmp_x + A(i, j) ^ 2
            Next j
        Next i
        tmp_max = Sqr(tmp_x)
    End If
    Matrix_Norm = tmp_max
End Function


'=== Check that A() is an identity matrix
Function Identity_Chk(A As Variant, Optional ierr As Double = 0, Optional tolerance As Double = 0.0000000001) As Long
Dim i As Long, j As Long, n As Long
Dim tmp_x As Double, err1 As Double, err2 As Double
    n = UBound(A, 1)
    'Verify that A is a square matrix
    If UBound(A, 2) <> n Then
        Debug.Print "Identity_Chk: Matrix is not square"
        Identity_Chk = 0
        Exit Function
    End If
    'Verify that diagonal elements equal to 1
    err1 = -999999
    For i = 1 To n
        tmp_x = Abs(A(i, i) - 1)
        If tmp_x > err1 Then err1 = tmp_x
    Next i
    'Verify that off-diagonals are zeroes
    err2 = -999999
    For i = 1 To n - 1
        For j = i + 1 To n
            If Abs(A(i, j)) > err2 Then err2 = Abs(A(i, j))
            If Abs(A(j, i)) > err2 Then err2 = Abs(A(j, i))
        Next j
    Next i
    If err1 < tolerance And err2 < tolerance Then
        Identity_Chk = 1
    Else
        Identity_Chk = 0
    End If
    ierr = err1 + err2
End Function


'=== Matrix Multiplication of C=AB
Function Mult_Matrix(A() As Double, B() As Double) As Double()
Dim i As Long, j As Long, k As Long
Dim c() As Double
    ReDim c(1 To UBound(A, 1), 1 To UBound(B, 2))
    For i = 1 To UBound(A, 1)
        For j = 1 To UBound(B, 2)
            For k = 1 To UBound(A, 2)
                c(i, j) = c(i, j) + A(i, k) * B(k, j)
            Next k
        Next j
    Next i
    Mult_Matrix = c
    Erase c
End Function


'=== Find determinant of matrix A()
'Input: A(), NxN square matrix
'Output: Determinant of A()
Function LUPDeterminant(A() As Double) As Double
Dim i As Long, n As Long
Dim A_tmp() As Double, p() As Long
Dim det As Double

    n = UBound(A, 1)
    If n = 1 Then
        LUPDeterminant = A(1, 1)
        Exit Function
    ElseIf n = 2 Then
        LUPDeterminant = A(1, 1) * A(2, 2) - A(1, 2) * A(2, 1)
        Exit Function
    End If
    A_tmp = A
    
    Call LUPDecompose(A_tmp, p)
    
    det = A_tmp(1, 1)
    For i = 2 To n
        det = det * A_tmp(i, i)
    Next i
    If (p(n + 1) - 1 - n) Mod 2 = 0 Then
        LUPDeterminant = det
    Else
        LUPDeterminant = -det
    End If
    
    Erase A_tmp, p
End Function
