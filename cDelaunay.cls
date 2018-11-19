VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDelaunay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pn As Long, pn_dimension As Long, pn_edge As Long
Private px() As Double
Private px_sort_idx() As Long
Private pEdge_L As cQuadEdge, pEdge_R As cQuadEdge
Private Const pz_tol = 0.000000000000001 '1E-15

Sub Reset()
Dim i As Long
    pn = 0
    pn_dimension = 0
    pn_edge = 0
    Erase px, px_sort_idx
    Set pEdge_L = Nothing
    Set pEdge_R = Nothing
End Sub

Public Property Get Size() As Long
    Size = pn
End Property

Public Property Get n_edge() As Long
    n_edge = pn_edge
End Property

Function Plot(Optional node_list As Variant) As Variant
Dim i As Long, j As Long, k As Long, m As Long
Dim vArr As Variant
Dim EdgeVisited As Collection
    Call Traverse_All_Edges(EdgeVisited)
    Call EdgeList_2_ExcelChart(vArr, EdgeVisited, node_list)
    Plot = vArr
    Set EdgeVisited = Nothing
End Function

Function Plot_DFS(Optional node_list As Variant) As Variant
Dim i As Long, j As Long, k As Long, m As Long
Dim vArr As Variant
Dim EdgeVisited As Collection
    Call Traverse_All_Nodes(EdgeVisited)
    Call EdgeList_2_ExcelChart(vArr, EdgeVisited, node_list)
    Plot_DFS = vArr
    Set EdgeVisited = Nothing
End Function

Private Sub EdgeList_2_ExcelChart(vArr As Variant, EdgeVisited As Collection, node_list As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim u As Long, v As Long, w As Long, u_prev As Long, v_prev As Long
Dim A As cQuadEdge, uArr As Variant
    
    m = EdgeVisited.Count
    k = 0: u_prev = -1: v_prev = -1
    For Each A In EdgeVisited
        k = k + 1
        If k Mod 500 = 0 Then
            DoEvents
            Application.StatusBar = "Plotting Edges: " & k & "/" & m
        End If
        u = A.Org
        v = A.Dest
        If u = v_prev Then
            m = UBound(vArr, 2)
            ReDim Preserve vArr(1 To 2, 1 To m + 1)
            ReDim Preserve node_list(1 To m + 1)
            vArr(1, m + 1) = px(v, 1)
            vArr(2, m + 1) = px(v, 2)
            node_list(m + 1) = v
        Else
            If k = 1 Then
                m = -1
                ReDim vArr(1 To 2, 1 To 2)
                ReDim node_list(1 To 2)
            Else
                m = UBound(vArr, 2)
                ReDim Preserve vArr(1 To 2, 1 To m + 3)
                ReDim Preserve node_list(1 To m + 3)
            End If
            vArr(1, m + 2) = px(u, 1)
            vArr(2, m + 2) = px(u, 2)
            vArr(1, m + 3) = px(v, 1)
            vArr(2, m + 3) = px(v, 2)
            node_list(m + 2) = u
            node_list(m + 3) = v
        End If
        u_prev = u: v_prev = v
    Next A
    
    n = UBound(vArr, 2)
    uArr = vArr
    ReDim vArr(1 To n, 1 To 2)
    For j = 1 To 2
        For i = 1 To n
            vArr(i, j) = uArr(j, i)
        Next i
    Next j
    Application.StatusBar = False
End Sub


Private Sub Traverse_All_Edges(EdgeVisited As Collection, Optional reset_visit As Boolean = True)
Dim i As Long, j As Long, k As Long, m As Long
Dim u As Long, v As Long, w As Long
Dim A As cQuadEdge, B As cQuadEdge
Dim EdgeStack As Collection
    
    Set EdgeVisited = New Collection    'edges in visited order
    Set EdgeStack = New Collection      'Stack of edges to be visited next
    Set A = pEdge_L                     'First edge pointer
    Set B = A
    Do
        Set B = B.Onext
        EdgeStack.Add B
        If B Is A Then Exit Do
    Loop
    k = 0: j = 0
    Do While EdgeStack.Count > 0
    
        k = k + 1
        If k Mod 2000 = 0 Then
            DoEvents
            Application.StatusBar = "Traversing Edges: " & EdgeStack.Count
        End If
        
        m = EdgeStack.Count
        Set A = EdgeStack(m)
        EdgeStack.Remove (m)

        If Not A.isVisited Then
            A.isVisited = True
            EdgeVisited.Add A
            Set B = A.Lnext
            Do
                If Not B.isVisited Then EdgeStack.Add B
                Set B = B.Oprev
                If B Is A.Sym Then Exit Do
            Loop
        End If
    Loop
    Set EdgeStack = Nothing
    If reset_visit = True Then
        For Each A In EdgeVisited
            A.isVisited = False
        Next A
    End If
    Application.StatusBar = False
End Sub


Private Sub Traverse_All_Nodes(EdgeVisited As Collection, Optional reset_visit As Boolean = True)
Dim i As Long, j As Long, k As Long, m As Long
Dim u As Long, v As Long, w As Long
Dim A As cQuadEdge, B As cQuadEdge
Dim EdgeStack As Collection
Dim isVisited() As Long

    ReDim isVisited(1 To pn)
    Set EdgeVisited = New Collection    'edges in visited order
    Set EdgeStack = New Collection      'Stack of edges to be visited next
    Set A = pEdge_L                     'First edge pointer
    isVisited(A.Org) = 1
    Set B = A
    Do
        Set B = B.Onext
        EdgeStack.Add B
        If B Is A Then Exit Do
    Loop
    k = 0: j = 0
    Do While EdgeStack.Count > 0
        k = k + 1
        If k Mod 2000 = 0 Then
            DoEvents
            Application.StatusBar = "Traversing Node: " & EdgeStack.Count & "/" & pn
        End If
        
        m = EdgeStack.Count
        Set A = EdgeStack(m)
        u = A.Org: v = A.Dest
        EdgeStack.Remove (m)
        
        If isVisited(v) = 0 Then
            isVisited(v) = 1
            EdgeVisited.Add A
            Set B = A.Lnext
            Do
                EdgeStack.Add B
                Set B = B.Oprev
                If B Is A.Sym Then Exit Do
            Loop
        End If
    Loop
    Set EdgeStack = Nothing
    If reset_visit = True Then
        For Each A In EdgeVisited
            A.isVisited = False
        Next A
    End If
    Application.StatusBar = False
End Sub


Sub Init(x As Variant)
Dim i As Long, j As Long, k As Long, m As Long
Dim A As cQuadEdge, B As cQuadEdge, C As cQuadEdge
Dim tmp_vec() As Double, iArr() As Long, xArr() As Double, tmp_x As Double
    pn = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    ReDim px(1 To pn, 1 To pn_dimension)
    For j = 1 To pn_dimension
        For i = 1 To pn
            px(i, j) = x(i, j)
        Next i
    Next j
    
    ReDim tmp_vec(1 To pn)
    For i = 1 To pn
        tmp_vec(i) = px(i, 1)
    Next i
    Call Sort_Quick_A(tmp_vec, 1, pn, px_sort_idx, 1)

    'If ties exist in first dimension then sort by the second dimension
    i = 2
    Do While i <= pn
        tmp_x = tmp_vec(i)
        If tmp_vec(i) = tmp_vec(i - 1) Then
            m = i
            For j = i + 1 To pn
                If Abs(tmp_vec(j) - tmp_vec(i)) > pz_tol Then
                    m = j - 1
                    Exit For
                End If
            Next j
            ReDim iArr(1 To m - i + 2)
            ReDim xArr(1 To m - i + 2)
            For k = 1 To m - i + 2
                iArr(k) = px_sort_idx(i + k - 2)
                xArr(k) = px(px_sort_idx(i + k - 2), 2)
            Next k
            Call Sort_Quick_A(xArr, 1, m - i + 2, iArr, 0)
            For k = 1 To m - i + 2
                px_sort_idx(i + k - 2) = iArr(k)
            Next k
            i = m + 2
        Else
            i = i + 1
        End If
    Loop
    
    With ActiveWorkbook.Sheets("Sheet4")
        For i = 1 To pn
            .Range("B" & 2 + i).Value = px_sort_idx(i)
        Next i
    End With
    
    Call Delaunay_Step(px_sort_idx, pEdge_L, pEdge_R)
    Application.StatusBar = False
End Sub


Private Sub Delaunay_Step(x_ptr() As Long, Edge_L As cQuadEdge, Edge_R As cQuadEdge)
Dim i As Long, j As Long, k As Long, n As Long, n_L As Long, n_R As Long
Dim A As cQuadEdge, B As cQuadEdge, C As cQuadEdge
Dim ldo As cQuadEdge, ldi As cQuadEdge, lcand As cQuadEdge
Dim rdo As cQuadEdge, rdi As cQuadEdge, rcand As cQuadEdge
Dim basel As cQuadEdge, t As cQuadEdge
Dim x_ptr_L() As Long, x_ptr_R() As Long
Dim isValid_L As Boolean, isValid_R As Boolean
Dim vArr As Variant, tmp_x As Double, tmp_y As Double

    n = UBound(x_ptr)
    
    If n = 2 Then
    
        Set A = New cQuadEdge: Call A.Init
        A.Org = x_ptr(1)
        A.Dest = x_ptr(2)
        Set Edge_L = A
        Set Edge_R = A.Sym

    ElseIf n = 3 Then

        Set A = New cQuadEdge: Call A.Init
        Set B = New cQuadEdge: Call B.Init
        Call Splice(A.Sym, B)
        A.Org = x_ptr(1)
        A.Dest = x_ptr(2)
        B.Org = x_ptr(2)
        B.Dest = x_ptr(3)
        If isCCW(x_ptr(1), x_ptr(2), x_ptr(3)) Then
            Set C = ConnectEdge(B, A)
            Set Edge_L = A
            Set Edge_R = B.Sym
        ElseIf isCCW(x_ptr(1), x_ptr(3), x_ptr(2)) Then
            Set C = ConnectEdge(B, A)
            Set Edge_L = C.Sym
            Set Edge_R = C
        Else
            Set Edge_L = A
            Set Edge_R = B.Sym
        End If

    ElseIf n >= 4 Then

        n_L = n \ 2
        n_R = n - n_L
        ReDim x_ptr_L(1 To n_L)
        ReDim x_ptr_R(1 To n_R)
        For i = 1 To n_L
            x_ptr_L(i) = x_ptr(i)
        Next i
        For i = 1 To n_R
            x_ptr_R(i) = x_ptr(n_L + i)
        Next i

        Call Delaunay_Step(x_ptr_L, ldo, ldi)
        Call Delaunay_Step(x_ptr_R, rdi, rdo)
        
        If n > (pn / 4) Then
            DoEvents
            Application.StatusBar = "Delaunay Triangulation: " & n & ": " & n_L & "-" & n_R
        End If
        
        Do
            If isLeftOf(rdi.Org, ldi) Then
                Set ldi = ldi.Lnext
            ElseIf isRightOf(ldi.Org, rdi) Then
                Set rdi = rdi.Rprev
            Else
                Exit Do
            End If
        Loop

        Set basel = ConnectEdge(rdi.Sym, ldi)
        If ldi.Org = ldo.Org Then Set ldo = basel.Sym
        If rdi.Org = rdo.Org Then Set rdo = basel
        
'        If basel.Dest = 108 And basel.Dest = 108 Then ' Or basel.Org = 109 Or basel.Org = 108 Then
'            Set pEdge_L = ldo
'            Set pEdge_R = rdo
'            vArr = Me.Plot
'            With ActiveWorkbook.Sheets("Sheet1")
'                .Range("E3").Resize(1048500, 3).Clear
'                .Range("E3").Resize(UBound(vArr, 1), UBound(vArr, 2)).Value = vArr
'            End With
'            End
'        End If
        
        Do

            Set lcand = basel.Sym.Onext
            If isValid(lcand, basel) Then
                Do While inCircle(basel.Dest, basel.Org, lcand.Dest, lcand.Onext.Dest)
                    Set t = lcand.Onext
                    Call DeleteEdge(lcand)
                    Set lcand = t
                Loop
            End If
            
            Set rcand = basel.Oprev
            If isValid(rcand, basel) Then
                Do While inCircle(basel.Dest, basel.Org, rcand.Dest, rcand.Oprev.Dest)
                    Set t = rcand.Oprev
                    Call DeleteEdge(rcand)
                    Set rcand = t
                Loop
            End If
            
            isValid_L = isValid(lcand, basel)
            isValid_R = isValid(rcand, basel)
            
            If Not isValid_L And Not isValid_R Then Exit Do
            
            If Not isValid_L Or _
                    (isValid_R And inCircle(lcand.Dest, lcand.Org, rcand.Org, rcand.Dest)) Then
                Set basel = ConnectEdge(rcand, basel.Sym)
            Else
                Set basel = ConnectEdge(basel.Sym, lcand.Sym)
            End If

        Loop

        Set Edge_L = ldo
        Set Edge_R = rdo
        
    End If

End Sub

Private Sub Splice(A As cQuadEdge, B As cQuadEdge)
Dim t1 As cQuadEdge, t2 As cQuadEdge, t3 As cQuadEdge, t4 As cQuadEdge
Dim alpha As cQuadEdge, beta As cQuadEdge

    Set alpha = A.Onext.Rot
    Set beta = B.Onext.Rot
    
    Set t1 = B.Onext
    Set t2 = A.Onext
    Set t3 = beta.Onext
    Set t4 = alpha.Onext
    
    Set A.Nxt = t1
    Set B.Nxt = t2
    Set alpha.Nxt = t3
    Set beta.Nxt = t4
    
End Sub

Private Function ConnectEdge(A As cQuadEdge, B As cQuadEdge) As cQuadEdge
Dim C As cQuadEdge
    Set C = New cQuadEdge
    Call C.Init
    C.Org = A.Dest
    C.Dest = B.Org
    Call Splice(C, A.Lnext)
    Call Splice(C.Sym, B)
    Set ConnectEdge = C
End Function

Private Sub DeleteEdge(A As cQuadEdge)
    Call Splice(A, A.Oprev)
    Call Splice(A.Sym, A.Sym.Oprev)
End Sub

Private Sub SwapEdge(C As cQuadEdge)
Dim A As cQuadEdge, B As cQuadEdge
    Set A = C.Oprev
    Set B = C.Sym.Oprev
    Call Splice(C, A)
    Call Splice(C.Sym, B)
    Call Splice(C, A.Lnext)
    Call Splice(C.Sym, B.Lnext)
    C.Org = A.Dest
    C.Dest = B.Dest
End Sub

Private Function inCircle(u As Long, v As Long, w As Long, p As Long) As Boolean
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim vArr As Variant
    If px(u, 1) = px(v, 1) And px(u, 1) = px(w, 1) Then
        inCircle = True
        Exit Function
    ElseIf px(u, 2) = px(v, 2) And px(u, 2) = px(w, 2) Then
        inCircle = True
        Exit Function
    End If
    ReDim vArr(1 To 4, 1 To 4)
    vArr(1, 1) = px(u, 1): vArr(1, 2) = px(u, 2): vArr(1, 3) = px(u, 1) ^ 2 + px(u, 2) ^ 2: vArr(1, 4) = 1
    vArr(2, 1) = px(v, 1): vArr(2, 2) = px(v, 2): vArr(2, 3) = px(v, 1) ^ 2 + px(v, 2) ^ 2: vArr(2, 4) = 1
    vArr(3, 1) = px(w, 1): vArr(3, 2) = px(w, 2): vArr(3, 3) = px(w, 1) ^ 2 + px(w, 2) ^ 2: vArr(3, 4) = 1
    vArr(4, 1) = px(p, 1): vArr(4, 2) = px(p, 2): vArr(4, 3) = px(p, 1) ^ 2 + px(p, 2) ^ 2: vArr(4, 4) = 1
    tmp_x = Application.WorksheetFunction.MDeterm(vArr)
    If tmp_x > 0 Then
        inCircle = True
    Else
        inCircle = False
    End If
End Function

Private Function isCCW(u As Long, v As Long, w As Long) As Boolean
Dim tmp_x As Double
    If px(u, 1) = px(v, 1) And px(u, 1) = px(w, 1) Then
        isCCW = False
    ElseIf px(u, 2) = px(v, 2) And px(u, 2) = px(w, 2) Then
        isCCW = False
    Else
        tmp_x = px(v, 1) * px(w, 2) - px(v, 2) * px(w, 1) _
                - (px(u, 1) * px(w, 2) - px(u, 2) * px(w, 1)) _
                + (px(u, 1) * px(v, 2) - px(u, 2) * px(v, 1))
        If tmp_x > 0 Then
            isCCW = True
        Else
            isCCW = False
        End If
    End If
End Function

Private Function isCollinear(u As Long, v As Long, w As Long) As Boolean
Dim tmp_x As Double
    If px(u, 1) = px(v, 1) And px(u, 1) = px(w, 1) Then
        isCollinear = True
    ElseIf px(u, 2) = px(v, 2) And px(u, 2) = px(w, 2) Then
        isCollinear = True
    Else
        tmp_x = px(v, 1) * px(w, 2) - px(v, 2) * px(w, 1)
        tmp_x = tmp_x - (px(u, 1) * px(w, 2) - px(u, 2) * px(w, 1))
        tmp_x = tmp_x + (px(u, 1) * px(v, 2) - px(u, 2) * px(v, 1))
        If Abs(tmp_x) <= pz_tol Then
            isCollinear = True
        Else
            isCollinear = False
        End If
    End If
End Function

Private Function isRightOf(u As Long, A As cQuadEdge) As Boolean
    isRightOf = isCCW(u, A.Dest, A.Org)
End Function

Private Function isLeftOf(u As Long, A As cQuadEdge) As Boolean
    isLeftOf = isCCW(u, A.Org, A.Dest)
End Function

Private Function isValid(A As cQuadEdge, basel As cQuadEdge)
    isValid = isRightOf(A.Dest, basel)
End Function

Private Sub Calc_Center(x As Double, y As Double, u As Long, v As Long, w As Long)
Dim tmp_det As Double, d_uv As Double, d_uw As Double
    tmp_det = px(v, 1) * px(w, 2) - px(v, 2) * px(w, 1)
    tmp_det = tmp_det - (px(u, 1) * px(w, 2) - px(u, 2) * px(w, 1))
    tmp_det = tmp_det + (px(u, 1) * px(v, 2) - px(u, 2) * px(v, 1))
    
    d_uv = px(u, 1) ^ 2 - px(v, 1) ^ 2 + px(u, 2) ^ 2 - px(v, 2) ^ 2
    d_uw = px(u, 1) ^ 2 - px(w, 1) ^ 2 + px(u, 2) ^ 2 - px(w, 2) ^ 2
    
    x = (d_uv * (px(u, 2) - px(w, 2)) - d_uw * (px(u, 2) - px(v, 2))) / (2 * tmp_det)
    y = -(d_uv * (px(u, 1) - px(w, 1)) - d_uw * (px(u, 1) - px(v, 1))) / (2 * tmp_det)
End Sub

Private Sub Calc_Dual_Coordinates()
Dim i As Long, j As Long, k As Long, m As Long, n_edge As Long
Dim u As Long, v As Long, w As Long, u_prev As Long, v_prev As Long, u_org As Long
Dim vArr As Variant, n_visited As Long
Dim A As cQuadEdge, B As cQuadEdge, C As cQuadEdge, B_prv As cQuadEdge
Dim EdgeStack As Collection, EdgeVisited As Collection
Dim node_visited() As Long
Dim tmp_x As Double, tmp_y As Double
    For Each A In pEdgeList
        Set B = A.Lnext
        u = A.Org
        v = A.Dest
        w = B.Dest
        If isCCW(u, v, w) Then
            If B.Lnext.Dest = u Then
                Call Calc_Center(tmp_x, tmp_y, u, v, w)
            End If
        End If
    
    Next A
End Sub


Private Sub Sort_Quick_A(vArray As Variant, inLow As Long, inHi As Long, sort_index() As Long, Optional first_run As Long = 1)
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



Sub Density_Est(x_density() As Double)
Dim i As Long, j As Long, k As Long
Dim A As cQuadEdge
Dim EdgeVisited As Collection
    Call Traverse_All_Nodes(EdgeVisited)
    ReDim x_density(1 To pn)
    Set A = EdgeVisited(1)
    x_density(A.Org) = 1# / Area_around_Vertex(A, "ORG")
    Do While EdgeVisited.Count > 0
        Set A = EdgeVisited(1)
        EdgeVisited.Remove (1)
        x_density(A.Dest) = 1# / Area_around_Vertex(A, "DEST")
    Loop
    Set EdgeVisited = Nothing
End Sub

    Private Function Area_around_Vertex(A As cQuadEdge, strType As String) As Double
    Dim i As Long, j As Long, k As Long
    Dim u As Long, v As Long, w As Long
    Dim tmp_x As Double
    Dim A_copy As cQuadEdge, B As cQuadEdge, C As cQuadEdge
        tmp_x = 0
        Set A_copy = A
        If strType = "ORG" Then
            u = A_copy.Org
        ElseIf strType = "DEST" Then
            u = A_copy.Dest
            Set A_copy = A_copy.Sym
        End If
        Set C = A_copy
        Do
            v = C.Dest
            k = 1: Set B = C
            Do
                k = k + 1
                If k > 3 Then Exit Do
                Set B = B.Lnext
                If B.Dest = u Then Exit Do
                w = B.Dest
            Loop
            If k = 3 Then
                tmp_x = tmp_x + Calc_Area_Triangle(u, v, w)
            End If
            Set C = C.Onext
            If C Is A_copy Then Exit Do
        Loop
        Area_around_Vertex = tmp_x
    End Function

    Private Function Calc_Area_Triangle(u As Long, v As Long, w As Long) As Double
    Calc_Area_Triangle = 0.5 * Abs(px(u, 1) * px(v, 2) + px(v, 1) * px(w, 2) + px(w, 1) * px(u, 2) _
                - px(v, 1) * px(u, 2) - px(w, 1) * px(v, 2) - px(u, 1) * px(w, 2))
    End Function


'Input: x(1:N,1:3), datapoints at x-y positions with value z
'Output:    x_lines(1:M,1:3), lines to chart dataset on a Delaunay triangulation
Sub Plot_xyz(x As Variant, x_lines As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim vArr As Variant, uArr As Variant
Dim x_pos() As Double
Dim x_avg() As Double, x_sd() As Double
Dim tmp_x As Double, tmp_y As Double

    n = UBound(x, 1)
    ReDim x_avg(1 To 3)
    ReDim x_sd(1 To 3)
    For j = 1 To 3
        tmp_x = 0: tmp_y = 0
        For i = 1 To n
            tmp_x = tmp_x + x(i, j)
            tmp_y = tmp_y + x(i, j) ^ 2
        Next i
        x_avg(j) = tmp_x / n
        x_sd(j) = Sqr(tmp_y / n - (tmp_x / n) ^ 2)
    Next j
    
    ReDim x_pos(1 To n, 1 To 2)
    For j = 1 To 2
        For i = 1 To n
            x_pos(i, j) = (x(i, j) - x_avg(j)) / x_sd(j)
        Next i
    Next j
    Call Init(x_pos)
    x_lines = Me.Plot(uArr)
    Call Me.Reset
    
    n = UBound(x_lines, 1)
    ReDim Preserve x_lines(1 To n, 1 To 3)
    For i = 1 To n
        If Not VBA.IsEmpty(uArr(i)) Then
            x_lines(i, 3) = (x(uArr(i), 3) - x_avg(3)) / x_sd(3)
        End If
    Next i

End Sub

