VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAffinityPropagation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pExemplar() As Long     'List of exemplars
Private pExemplar_num As Long   'Number of exemplars
Private pExemplar_index() As Long   'The exemplar that each member maps to
Private pconvergence_chk() As Double
Private pNet_Similarity As Double
Private pExemplar_pref As Double
Private pdata_to_exemplar_similarities As Double



Sub Reset()
    Erase pExemplar, pExemplar_index, pconvergence_chk
End Sub

Public Property Get Exemplars() As Long()
Exemplars = pExemplar
End Property

Public Property Get Exemplar_num() As Long
Exemplar_num = pExemplar_num
End Property

Public Property Get Exemplar_index() As Long()
Exemplar_index = pExemplar_index
End Property

Public Property Get converge_chk() As Double()
converge_chk = pconvergence_chk
End Property

Public Property Get Net_Similarity() As Double
Net_Similarity = pNet_Similarity
End Property

Public Property Get Exemplar_pref() As Double
Exemplar_pref = pExemplar_pref
End Property

Public Property Get data_to_exemplar_similarities() As Double
data_to_exemplar_similarities = pdata_to_exemplar_similarities
End Property


'Input: S(), N x N similarity matrix, not necessarily symmetric,
'           diagonals should either be 0 or be weights for preferences
'Preference input can be: MAX, MIN, MEDIAN or INPUT
Sub Affinity_Propagation(S_in() As Double, _
    Optional damping As Double = 0.5, _
    Optional iterate_max As Long = 500, _
    Optional convit As Long = 30, _
    Optional input_pref As String = "MEDIAN")
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim iterate As Long, converge_count As Long
Dim n_raw As Long
Dim tmp_x As Double, tmp_y As Double
Dim temp As Double, temp_max As Double, temp_min As Double
Dim s() As Double
Dim S_Median As Double, s_min As Double, s_max As Double
Dim r() As Double, A() As Double, R_old() As Double, A_Old() As Double
Dim AnS() As Double, AnS_Max() As Double, AnS_Max_Index() As Long
Dim Rp() As Double, Rp_sum() As Double
Dim e() As Double
Dim Exemplar_index_old() As Long
Dim tmp_vec() As Double

s = S_in
input_pref = UCase(input_pref)
n_raw = UBound(s, 1)
ReDim pExemplar(1 To n_raw)
ReDim pExemplar_index(1 To n_raw)
ReDim Exemplar_index_old(1 To n_raw)
ReDim pconvergence_chk(1 To 3, 1 To iterate_max)

ReDim r(1 To n_raw, 1 To n_raw)
ReDim Rp(1 To n_raw, 1 To n_raw)
ReDim Rp_sum(1 To n_raw)
ReDim A(1 To n_raw, 1 To n_raw)
ReDim AnS(1 To n_raw, 1 To n_raw)
ReDim AnS_Max(1 To n_raw, 1 To 2)
ReDim AnS_Max_Index(1 To n_raw, 1 To 2)
ReDim e(1 To n_raw, 1 To n_raw)


'=== Assign Preference
'Find median, min, max of S(i,k) s.t. i<>k
Application.StatusBar = "Affinity: assigning preferences..."
n = 0
ReDim tmp_vec(1 To n_raw * (n_raw - 1))
For i = 1 To n_raw - 1
    For k = i + 1 To n_raw
        n = n + 1
        tmp_vec(n) = s(i, k)
        n = n + 1
        tmp_vec(n) = s(k, i)
    Next k
Next i
Call Find_Med_Min_Max(tmp_vec, S_Median, s_min, s_max)
If input_pref = "MEDIAN" Then
    For i = 1 To n_raw
        s(i, i) = S_Median
    Next i
ElseIf input_pref = "MAX" Then
    For i = 1 To n_raw
        s(i, i) = s_max
    Next i
ElseIf input_pref = "MIN" Then
    For i = 1 To n_raw
        s(i, i) = s_min
    Next i
ElseIf input_pref = "INPUT" Then
    ReDim tmp_vec(1 To n_raw)
    For i = 1 To n_raw
        tmp_vec(i) = s(i, i)
    Next i
    Call Find_Med_Min_Max(tmp_vec, temp, temp_min, temp_max)
    For i = 1 To n_raw
        s(i, i) = s_min + (s_max - s_min) * (s(i, i) - temp_min) / (temp_max - temp_min)
    Next i
Else
    msgbox input_pref & " is not a valid input"
End If
'====================================================

'=== Add Random Noise
For i = 1 To n_raw
    Randomize
    For k = 1 To n_raw
        s(i, k) = s(i, k) + 0.000000000001 * Rnd * (s_max - s_min)
    Next k
Next i
'====================================================


converge_count = 0
For iterate = 1 To iterate_max

    DoEvents
    If iterate Mod 10 = 0 Then Application.StatusBar = "Affinity: Iterate: " & iterate & "/" & iterate_max
    
    '=== Compute responsibilities
    R_old = r
    
    For i = 1 To n_raw
        For k = 1 To n_raw
            AnS(i, k) = A(i, k) + s(i, k)
        Next k
    Next i
    
    For i = 1 To n_raw
        'Largest element on the i-th row
        tmp_x = -999999999
        For k = 1 To n_raw
            If AnS(i, k) > AnS_Max(i, 1) Then
                tmp_x = AnS(i, k)
                j = k
            End If
            AnS_Max(i, 1) = tmp_x
            AnS_Max_Index(i, 1) = j
        Next k
        
        '2nd Largest element on the i-th row
        tmp_x = -999999999
        tmp_y = AnS_Max(i, 1)
        m = AnS_Max_Index(i, 1)
        For k = 1 To n_raw
            If AnS(i, k) > tmp_x And AnS(i, k) <= tmp_y And k <> m Then
                tmp_x = AnS(i, k)
                j = k
            End If
            AnS_Max(i, 2) = tmp_x
            AnS_Max_Index(i, 2) = j
        Next k
    Next i
    
    
    For i = 1 To n_raw
        For k = 1 To n_raw
            If AnS_Max_Index(i, 1) <> k Then
                r(i, k) = s(i, k) - AnS_Max(i, 1)
            ElseIf AnS_Max_Index(i, 1) = k Then
                r(i, k) = s(i, k) - AnS_Max(i, 2)
            End If
        Next k
    Next i
    
    For i = 1 To n_raw
        For k = 1 To n_raw
            r(i, k) = (1 - damping) * r(i, k) + damping * R_old(i, k)
        Next k
    Next i
    '======================================
    
    '=== Compute availabilities
    A_Old = A
    
    For i = 1 To n_raw
        For k = 1 To n_raw
            Rp(i, k) = 0
            If r(i, k) > 0 Then Rp(i, k) = r(i, k)
        Next k
    Next i
    
    For k = 1 To n_raw
        Rp_sum(k) = 0
        For i = 1 To n_raw
            If i <> k Then Rp_sum(k) = Rp_sum(k) + Rp(i, k)
        Next i
    Next k
    
    For i = 1 To n_raw
        For k = 1 To n_raw
            If i <> k Then
                A(i, k) = r(k, k) + Rp_sum(k) - Rp(i, k)
                If A(i, k) > 0 Then A(i, k) = 0
            End If
        Next k
    Next i
    
    
    For k = 1 To n_raw
        A(k, k) = Rp_sum(k)
    Next k
    
    For i = 1 To n_raw
        For j = 1 To n_raw
            A(i, j) = (1 - damping) * A(i, j) + damping * A_Old(i, j)
        Next j
    Next i
    '======================================
    
    '=== Exemplar in current iteration
    For i = 1 To n_raw
        For k = 1 To n_raw
            e(i, k) = r(i, k) + A(i, k)
        Next k
    Next i
    
    Exemplar_index_old = pExemplar_index
    
    ReDim pExemplar(1 To n_raw)
    ReDim pExemplar_index(1 To n_raw)
    
    For i = 1 To n_raw
        temp_max = -999999
        For k = 1 To n_raw
            If e(i, k) > temp_max Then
                temp_max = e(i, k)
                pExemplar_index(i) = k
            End If
        Next k
    Next i
    
    pExemplar_num = 0
    For i = 1 To n_raw
        If pExemplar_index(i) = i Then
            pExemplar_num = pExemplar_num + 1
            pExemplar(pExemplar_num) = i
        End If
    Next i
    
    If pExemplar_num > 0 Then ReDim Preserve pExemplar(1 To pExemplar_num)
    
    pNet_Similarity = 0
    For i = 1 To n_raw
        pNet_Similarity = pNet_Similarity + s(i, pExemplar_index(i))
    Next i
    
    pconvergence_chk(1, iterate) = iterate
    pconvergence_chk(2, iterate) = pExemplar_num
    pconvergence_chk(3, iterate) = pNet_Similarity
    
    '=== Check for convergence
    n = 0
    For i = 1 To n_raw
        If pExemplar_index(i) = Exemplar_index_old(i) Then n = n + 1
    Next i
    If n = n_raw Then
        converge_count = converge_count + 1
    Else
        converge_count = 0
    End If
    If converge_count = convit Then Exit For
    '==================================
    
Next iterate

ReDim Preserve pconvergence_chk(1 To 3, 1 To iterate)

Erase A_Old, R_old, Rp, Rp_sum, AnS, AnS_Max, AnS_Max_Index, Exemplar_index_old

For i = 1 To n_raw
    For k = 1 To n_raw
        e(i, k) = r(i, k) + A(i, k)
    Next k
Next i

ReDim pExemplar(1 To n_raw)
ReDim pExemplar_index(1 To n_raw)

For i = 1 To n_raw
    temp_max = -999999
    For k = 1 To n_raw
        If e(i, k) > temp_max Then
            temp_max = e(i, k)
            pExemplar_index(i) = k
        End If
    Next k
Next i

pExemplar_num = 0
For i = 1 To n_raw
    If pExemplar_index(i) = i Then
        pExemplar_num = pExemplar_num + 1
        pExemplar(pExemplar_num) = i
    End If
Next i

ReDim Preserve pExemplar(1 To pExemplar_num)

pNet_Similarity = 0
For i = 1 To n_raw
    pNet_Similarity = pNet_Similarity + s(i, pExemplar_index(i))
Next i

pExemplar_pref = 0
For k = 1 To pExemplar_num
    pExemplar_pref = pExemplar_pref + s(pExemplar(k), pExemplar(k))
Next k

pdata_to_exemplar_similarities = pNet_Similarity - pExemplar_pref

Application.StatusBar = False

End Sub

Private Sub Find_Med_Min_Max(x() As Double, x_med As Double, x_min As Double, x_max As Double)
Dim n As Long
    n = UBound(x)
    Call modMath.Sort_Quick(x, 1, n)
    x_min = x(1)
    x_max = x(n)
    If n Mod 2 = 1 Then
        x_med = x((n + 1) / 2)
    ElseIf n Mod 2 = 0 Then
        x_med = (x(n / 2) + x(n / 2 + 1)) / 2
    End If
End Sub
