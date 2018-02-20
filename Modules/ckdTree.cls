VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ckdTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===============================================
'In this implementation, splitting dimension is cyclically selected,
'starting from the first. Splitting node is chosen to be the median
'along that dimension.
'Distance type supports: "EUCLIDEAN", "MANHATTAN", "MAXNORM".
'===============================================
Private pn_dimension As Long
Private pParent As ckdTree, pChild_L As ckdTree, pChild_R As ckdTree
Private pDepth As Long, pisLeaf As Long, pPtr As Long
Private pSplitDim As Long, pSplitVal As Double

'==========================================
'Access Properties
'==========================================
Public Property Get parent() As ckdTree
    Set parent = pParent
End Property

Public Property Get Child_L() As ckdTree
    Set Child_L = pChild_L
End Property

Public Property Get Child_R() As ckdTree
    Set Child_R = pChild_R
End Property

Public Property Get Depth() As Long
    Depth = pDepth
End Property

Public Property Get ptr() As Long
    ptr = pPtr
End Property

Public Property Get SplitDim() As Long
    SplitDim = pSplitDim
End Property


'==========================================
'Main sub used to create k-d Tree from raw data x(1:N,1:D)
'==========================================
Sub Build_Tree(x As Variant, Optional min_size As Long = 1, Optional max_depth As Long = 500)
Dim i As Long, j As Long, k As Long, n As Long
Dim idx() As Long
    n = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    pDepth = 0
    pisLeaf = 0
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    Call z_Build_Tree_Recursive(x, idx, 0, min_size, max_depth, pn_dimension, Nothing)
End Sub

'Recursive step
Sub z_Build_Tree_Recursive(x As Variant, idx() As Long, cur_depth As Long, _
        min_size As Long, max_depth As Long, n_dimension As Long, vparent As ckdTree)
Dim i As Long, j As Long, k As Long, n As Long, n1 As Long, n2 As Long
Dim idx1() As Long, idx2() As Long
Dim xk() As Double
    n = UBound(idx)
    pDepth = cur_depth
    pn_dimension = n_dimension
    Set pParent = vparent
    If n <= min_size Or cur_depth >= max_depth Then
        pisLeaf = 1
        pPtr = idx(1)
        Set pChild_L = Nothing
        Set pChild_R = Nothing
        Exit Sub
    End If

    pSplitDim = cur_depth Mod n_dimension + 1
    ReDim xk(1 To n)
    For i = 1 To n
        xk(i) = x(idx(i), pSplitDim)
    Next i
    If n > 1 Then Call modMath.Sort_Quick_A(xk, 1, n, idx, 0)
    i = n \ 2 + 1
    pSplitVal = xk(i)
    pPtr = idx(i)

    n1 = 0: ReDim idx1(1 To n)
    n2 = 0: ReDim idx2(1 To n)
    For i = 1 To n
        j = idx(i)
        If j <> pPtr Then
            If xk(i) < pSplitVal Then
                n1 = n1 + 1
                idx1(n1) = j
            Else
                n2 = n2 + 1
                idx2(n2) = j
            End If
        End If
    Next i
    If n1 > 0 Then
        ReDim Preserve idx1(1 To n1)
        Set pChild_L = New ckdTree
        Call pChild_L.z_Build_Tree_Recursive(x, idx1, cur_depth + 1, min_size, max_depth, n_dimension, Me)
    Else
        Set pChild_L = Nothing
    End If

    If n2 > 0 Then
        ReDim Preserve idx2(1 To n2)
        Set pChild_R = New ckdTree
        Call pChild_R.z_Build_Tree_Recursive(x, idx2, cur_depth + 1, min_size, max_depth, n_dimension, Me)
    Else
        Set pChild_R = Nothing
    End If
    
    Erase xk, idx1, idx2
End Sub

Sub Reset()
    If Not pChild_L Is Nothing Then Call pChild_L.Reset
    If Not pChild_R Is Nothing Then Call pChild_R.Reset
    Set pChild_L = Nothing
    Set pChild_R = Nothing
End Sub

'===========================================================
'For a given vector x_tgt(1:D), return an integer pointer to
'its nearest neighbor in x(1:N,1:D). Use after tree is built
'===========================================================
Function NN_Search(x_tgt As Variant, x As Variant, Optional dist_type As String = "EUCLIDEAN") As Long
Dim i As Long, j As Long, k As Long, n As Long
Dim best_node As ckdTree, v As ckdTree, u As ckdTree, vparent As ckdTree
Dim vstack As Collection
Dim tmp_min As Double, tmp_x As Double
Dim isVisited() As Long
    If pSplitDim <= 0 Then
        Debug.Print "ckdTree: NN_Search: Tree is not built yet."
        End
    End If
    n = UBound(x, 1)
    ReDim isVisited(1 To n)
    Set vstack = New Collection
    tmp_min = 2.5E+30
    Call z_Add_to_Stack(x_tgt, x, vstack, isVisited, tmp_min, best_node, dist_type)
    Do While vstack.Count > 0
        With vstack
            Set v = .Item(.Count)
            .Remove .Count
        End With
        If v.Depth = 0 Then Exit Do
        Set vparent = v.parent
        With vparent
            j = .SplitDim
            If Intersect_Chk(tmp_min, x_tgt(j), x(.ptr, j), dist_type) = True Then
            'If tmp_min >= ((x_tgt(j) - x(.ptr, j)) ^ 2) Then
                Set u = .Child_L
                If u Is v Then Set u = .Child_R
            End If
        End With
        If Not u Is Nothing Then
            Call u.z_Add_to_Stack(x_tgt, x, vstack, isVisited, tmp_min, best_node, dist_type)
        End If
    Loop
    NN_Search = best_node.ptr
    Erase isVisited
End Function

'Traverse down the tree from current node, adding each node to stack until a leaf is reached
Sub z_Add_to_Stack(x_tgt As Variant, x As Variant, vstack As Collection, isVisited() As Long, _
                    min_dist As Double, best_node As ckdTree, dist_type As String)
Dim tmp_x As Double
Dim u As ckdTree
    If isVisited(pPtr) = 0 Then
        isVisited(pPtr) = 1
        vstack.Add Me
        tmp_x = Dist2Vec(x_tgt, x, dist_type)
        If tmp_x < min_dist Then
            min_dist = tmp_x
            Set best_node = Me
        End If
    End If
    If pisLeaf = 1 Then Exit Sub
    If x_tgt(pSplitDim) < pSplitVal Then
        Set u = pChild_L
        If u Is Nothing Then Set u = pChild_R
    Else
        Set u = pChild_R
        If u Is Nothing Then Set u = pChild_L
    End If
    Call u.z_Add_to_Stack(x_tgt, x, vstack, isVisited, min_dist, best_node, dist_type)
End Sub

'Nearest Neighbor search speed up with pre-calculated distance
Function NN_Search_Speed(x_tgt As Variant, x As Variant, tgt As Long, tgt_dist() As Double, _
    Optional dist_type As String = "EUCLIDEAN") As Long
Dim i As Long, j As Long, k As Long, n As Long
Dim best_node As ckdTree, v As ckdTree, u As ckdTree, vparent As ckdTree
Dim vstack As Collection
Dim tmp_min As Double, tmp_x As Double
Dim isVisited() As Long
    If pSplitDim <= 0 Then
        Debug.Print "ckdTree: NN_Search_Speed: Tree is not built yet."
        End
    End If
    n = UBound(x, 1)
    ReDim isVisited(1 To n)
    Set vstack = New Collection
    tmp_min = 2.5E+30
    Call z_Add_to_Stack_Speed(x_tgt, x, vstack, isVisited, tmp_min, best_node, dist_type, _
            tgt, tgt_dist)
    Do While vstack.Count > 0
        With vstack
            Set v = .Item(.Count)
            .Remove .Count
        End With
        If v.Depth = 0 Then Exit Do
        Set vparent = v.parent
        With vparent
            j = .SplitDim
            If Intersect_Chk(tmp_min, x_tgt(j), x(.ptr, j), dist_type) = True Then
            'If tmp_min >= ((x_tgt(j) - x(.ptr, j)) ^ 2) Then
                Set u = .Child_L
                If u Is v Then Set u = .Child_R
            End If
        End With
        If Not u Is Nothing Then
            Call u.z_Add_to_Stack_Speed(x_tgt, x, vstack, isVisited, tmp_min, best_node, dist_type, _
                    tgt, tgt_dist)
        End If
    Loop
    NN_Search_Speed = best_node.ptr
    Erase isVisited
End Function

Sub z_Add_to_Stack_Speed(x_tgt As Variant, x As Variant, vstack As Collection, isVisited() As Long, _
                    min_dist As Double, best_node As ckdTree, dist_type As String, _
                    tgt As Long, tgt_dist() As Double)
Dim tmp_x As Double
Dim u As ckdTree
    If isVisited(pPtr) = 0 Then
        isVisited(pPtr) = 1
        vstack.Add Me
        'Only calculate distance if centroid has moved or if it's not done before
        If tgt_dist(tgt, pPtr) < 0 Then
            tgt_dist(tgt, pPtr) = Dist2Vec(x_tgt, x, dist_type)
        End If
        tmp_x = tgt_dist(tgt, pPtr)
        If tmp_x < min_dist Then
            min_dist = tmp_x
            Set best_node = Me
        End If
    End If
    If pisLeaf = 1 Then Exit Sub
    If x_tgt(pSplitDim) < pSplitVal Then
        Set u = pChild_L
        If u Is Nothing Then Set u = pChild_R
    Else
        Set u = pChild_R
        If u Is Nothing Then Set u = pChild_L
    End If
    Call u.z_Add_to_Stack_Speed(x_tgt, x, vstack, isVisited, min_dist, best_node, dist_type, _
            tgt, tgt_dist)
End Sub


'=== Find k nearest neighbors for every data point in x()
'Input:  x(1:N,1:D), N datapoints of D dimensional vector
'        k, desired number of neighbors
'        kth_only, if set to 1 then only the k-th neighbor is returned, otherwise a list of k-neighbors is returned
'Output: k_idx(1:N) or k_idx(1:N,1:k)
'        k_dist(1:N) or k_dist(1:N,1:k)
'===========================================================
Sub kNN_All(k_idx() As Long, k_dist() As Double, x() As Double, k As Long, _
        Optional kth_only As Long = 1, Optional dist_type As String = "EUCLIDEAN")
Dim i As Long, j As Long, n As Long
Dim kNeighbor() As Long
Dim kDist() As Double
    n = UBound(x, 1)
    Call Build_Tree(x)
    If kth_only = 1 Then
        ReDim k_idx(1 To n)
        ReDim k_dist(1 To n)
    Else
        ReDim k_idx(1 To n, 1 To k)
        ReDim k_dist(1 To n, 1 To k)
    End If
    For i = 1 To n
        If i Mod 250 = 0 Then
            DoEvents: Application.StatusBar = "kNN (kdtree): " & i & "/" & n
        End If
        Call kNN_Search(i, k, x, kNeighbor, kDist, dist_type)
        If kth_only = 1 Then
            k_idx(i) = kNeighbor(k)
            k_dist(i) = kDist(k)
        Else
            For j = 1 To k
                k_idx(i, j) = kNeighbor(j)
                k_dist(i, j) = kDist(j)
            Next j
        End If
    Next i
    Erase kNeighbor, kDist
    Application.StatusBar = False
End Sub


'===========================================================
'Find k-nearest neighbours of target node. Different from NN_Search(), tgt here is
'a member of x(). So tgt itself is excluded from the search. Use after tree is built.
'Input:  tgt, integer pointer of target data
'        k_nearest, desired number of neighbors
'        x(1:N,1:D), N observations of D-dimensional data
'        allowTie, if set to TRUE, will allow more than K neighbors to return
'Output: kNeighbor(1:K), integer pointers to nearest neighbors
'        kDist(1:K), distance to the k-nearest neighbors
'===========================================================
Sub kNN_Search(tgt As Long, k_nearest As Long, x As Variant, kNeighbor() As Long, kDist() As Double, _
        Optional dist_type As String = "EUCLIDEAN", Optional allowTie As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long
Dim v As ckdTree, u As ckdTree, vparent As ckdTree
Dim vstack As Collection
Dim INFINITY As Double
Dim isVisited() As Long
    INFINITY = Exp(70)
    n = UBound(x, 1)
    ReDim isVisited(1 To n)
    ReDim kNeighbor(1 To k_nearest)
    ReDim kDist(1 To k_nearest)
    Set vstack = New Collection
    For i = 1 To k_nearest
        kDist(i) = INFINITY
    Next i
    Call z_Add_to_Stack_kNN(tgt, x, vstack, isVisited, k_nearest, kNeighbor, kDist, dist_type, allowTie)
    Do While vstack.Count > 0
        With vstack
            Set v = .Item(.Count)
            .Remove .Count
        End With
        If v.Depth = 0 Then Exit Do
        Set vparent = v.parent
        With vparent
            j = .SplitDim
            If Intersect_Chk(kDist(UBound(kDist)), x(tgt, j), x(.ptr, j), dist_type) = True Then
                Set u = .Child_L
                If u Is v Then Set u = .Child_R
            End If
        End With
        If Not u Is Nothing Then
            If isVisited(u.ptr) = 0 Then
                Call u.z_Add_to_Stack_kNN(tgt, x, vstack, isVisited, k_nearest, kNeighbor, kDist, dist_type, allowTie)
            End If
        End If
    Loop
    Set vstack = Nothing
    Erase isVisited
    'Take square root of distance values
    If dist_type = "EUCLIDEAN" Then
        For i = 1 To UBound(kDist)
            kDist(i) = Sqr(kDist(i))
        Next i
    End If
End Sub


'Move down the tree until a leaf is reached, add each visited node to stack
'Evaluate distance of each visited node and add to k-Nearest list if suitable
Sub z_Add_to_Stack_kNN(tgt As Long, x As Variant, vstack As Collection, isVisited() As Long, _
                    k_nearest As Long, kNeighbor() As Long, kDist() As Double, dist_type As String, allowTie As Boolean)
    Dim u As ckdTree
    vstack.Add Me
    'Make comparison only if it's not already compared or if its not the target itself
    If isVisited(pPtr) = 0 Then
        isVisited(pPtr) = 1
        If tgt <> pPtr Then
            Call Queue_Eject(kNeighbor, kDist, Dist2Node(tgt, x, dist_type), pPtr, k_nearest, allowTie)
        End If
    End If
    If pisLeaf = 1 Then Exit Sub
    If x(tgt, pSplitDim) < pSplitVal Then
        Set u = pChild_L
        If u Is Nothing Then Set u = pChild_R
    Else
        Set u = pChild_R
        If u Is Nothing Then Set u = pChild_L
    End If
    Call u.z_Add_to_Stack_kNN(tgt, x, vstack, isVisited, k_nearest, kNeighbor, kDist, dist_type, allowTie)
End Sub


'Remove point from kNN list if point v is closer than top of queue
Private Sub Queue_Eject(kNeighbor() As Long, kDist() As Double, cur_dist As Double, v As Long, k As Long, Optional allowTie As Boolean = False)
Dim i As Long, j As Long, n As Long
    If allowTie = False Then
        If cur_dist < kDist(k) Then
            For i = 1 To k
                If cur_dist < kDist(i) Then
                    For j = k To i + 1 Step -1
                        kDist(j) = kDist(j - 1)
                        kNeighbor(j) = kNeighbor(j - 1)
                    Next j
                    kDist(i) = cur_dist
                    kNeighbor(i) = v
                    Exit For
                End If
            Next i
        End If
    Else
        n = UBound(kDist)
        If cur_dist <= kDist(n) Then
            n = n + 1
            ReDim Preserve kDist(1 To n)
            ReDim Preserve kNeighbor(1 To n)
            For i = 1 To n - 1
                If cur_dist <= kDist(i) Then
                    For j = n To i + 1 Step -1
                        kDist(j) = kDist(j - 1)
                        kNeighbor(j) = kNeighbor(j - 1)
                    Next j
                    kDist(i) = cur_dist
                    kNeighbor(i) = v
                    Exit For
                End If
            Next i
            If n > k Then
                If kDist(k + 1) > kDist(k) Then
                    ReDim Preserve kDist(1 To k)
                    ReDim Preserve kNeighbor(1 To k)
                End If
            End If
        End If
    End If
End Sub


'Find distance from a target node to current node, return is squared
Private Function Dist2Node(tgt As Long, x As Variant, strType As String) As Double
Dim j As Long
Dim tmp_x As Double
    Dist2Node = 0
    If strType = "EUCLIDEAN" Then
        For j = 1 To pn_dimension
            Dist2Node = Dist2Node + (x(tgt, j) - x(pPtr, j)) ^ 2
        Next j
    ElseIf strType = "MANHATTAN" Then
        For j = 1 To pn_dimension
            Dist2Node = Dist2Node + Abs(x(tgt, j) - x(pPtr, j))
        Next j
    ElseIf strType = "MAXNORM" Then
        For j = 1 To pn_dimension
            tmp_x = Abs(x(tgt, j) - x(pPtr, j))
            If tmp_x > Dist2Node Then Dist2Node = tmp_x
        Next j
    Else
        Debug.Print "ckdTree:Dist2Node:Invalid Distance Type."
        End
    End If
End Function

'Find distance from a vector x_tgt() to current node, return is squared
Private Function Dist2Vec(x_tgt As Variant, x As Variant, strType As String) As Double
Dim j As Long
Dim tmp_x As Double
    Dist2Vec = 0
    If strType = "EUCLIDEAN" Then
        For j = 1 To pn_dimension
            Dist2Vec = Dist2Vec + (x_tgt(j) - x(pPtr, j)) ^ 2
        Next j
    ElseIf strType = "MANHATTAN" Then
        For j = 1 To pn_dimension
            Dist2Vec = Dist2Vec + Abs(x_tgt(j) - x(pPtr, j))
        Next j
    ElseIf strType = "MAXNORM" Then
        For j = 1 To pn_dimension
            tmp_x = Abs(x_tgt(j) - x(pPtr, j))
            If tmp_x > Dist2Vec Then Dist2Vec = tmp_x
        Next j
    Else
        Debug.Print "ckdTree:Dist2Vec: Invalid distance metric."
        End
    End If
End Function

Private Function Intersect_Chk(min_dist As Double, search_pt As Variant, node As Variant, strType As String) As Boolean
    Intersect_Chk = False
    If strType = "EUCLIDEAN" Then
        If min_dist >= ((search_pt - node) ^ 2) Then Intersect_Chk = True
    ElseIf strType = "MANHATTAN" Then
        If min_dist >= Abs(search_pt - node) Then Intersect_Chk = True
    ElseIf strType = "MAXNORM" Then
        If min_dist >= Abs(search_pt - node) Then Intersect_Chk = True
    Else
        Debug.Print "ckdTree:Intersect_Chk: Invalid distance metric."
        End
    End If
End Function

'======================================================================
'For 2D-tree, output splitting lines to visualize the area partitions
'======================================================================
Sub Visualize_2DTree(vArr As Variant, x_min As Double, x_max As Double, y_min As Double, y_max As Double)
    If pn_dimension <> 2 Then
        Debug.Print "ckdTree:Visualize: only works for 2D."
        End
    End If
    Call z_Visualize_Recursive(vArr, x_min, x_max, y_min, y_max)
End Sub

Sub z_Visualize_Recursive(vArr As Variant, x_min As Double, x_max As Double, _
            y_min As Double, y_max As Double, Optional n As Long = 1)
    If pisLeaf = 1 Then
        n = n - 3
        Exit Sub
    End If
    If n = 1 Then
        ReDim vArr(1 To 2, 1 To n + 1)
    Else
        ReDim Preserve vArr(1 To 2, 1 To n + 1)
    End If
    If pSplitDim = 1 Then
        vArr(1, n) = pSplitVal
        vArr(2, n) = y_min
        vArr(1, n + 1) = pSplitVal
        vArr(2, n + 1) = y_max
    ElseIf pSplitDim = 2 Then
        vArr(1, n) = x_min
        vArr(2, n) = pSplitVal
        vArr(1, n + 1) = x_max
        vArr(2, n + 1) = pSplitVal
    End If
    If Not pChild_L Is Nothing Then
        n = n + 3
        If pSplitDim = 1 Then
            Call pChild_L.z_Visualize_Recursive(vArr, x_min, pSplitVal, y_min, y_max, n)
        Else
            Call pChild_L.z_Visualize_Recursive(vArr, x_min, x_max, y_min, pSplitVal, n)
        End If
    End If
    If Not pChild_R Is Nothing Then
        n = n + 3
        If pSplitDim = 1 Then
            Call pChild_R.z_Visualize_Recursive(vArr, pSplitVal, x_max, y_min, y_max, n)
        Else
            Call pChild_R.z_Visualize_Recursive(vArr, x_min, x_max, pSplitVal, y_max, n)
        End If
    End If
End Sub


'=====================================================
'*************************************************************
'Dedicated to filtering algorithm for fast k-Means clustering
'*************************************************************
'=====================================================
'Build tree to be used in k-Means filtering algorithm
'=====================================================
Sub kMean_Build_Tree(x As Variant, node_size() As Long, node_wgtcenter As Variant, _
    node_min() As Double, node_max() As Double, Optional dist_type As String = "EUCLIDEAN")
Dim i As Long, j As Long, k As Long, n As Long
Dim idx() As Long
    n = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    pDepth = 0
    pisLeaf = 0
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    ReDim node_size(1 To n)
    If dist_type = "EUCLIDEAN" Then
        ReDim node_wgtcenter(1 To n, 1 To pn_dimension)
    ElseIf dist_type = "MANHATTAN" Then
        ReDim node_wgtcenter(1 To n)
    End If
    ReDim node_min(1 To n, 1 To pn_dimension)
    ReDim node_max(1 To n, 1 To pn_dimension)
    Call z_kMean_Build_Tree_Recursive(x, idx, 0, pn_dimension, _
        node_size, node_wgtcenter, node_min, node_max, dist_type)
End Sub

'Recursive step caled by kMean_Build_Tree
Sub z_kMean_Build_Tree_Recursive(x As Variant, idx() As Long, cur_depth As Long, _
        n_dimension As Long, _
        node_size() As Long, node_wgtcenter As Variant, node_min() As Double, node_max() As Double, _
        dist_type As String)
Dim i As Long, j As Long, k As Long, n As Long, n1 As Long, n2 As Long
Dim idx1() As Long, idx2() As Long
Dim xk() As Double, tmp_vec() As Double
Dim tmp_x As Double, tmp_y As Double, tmp_min As Double, tmp_max As Double
Dim INFINITY As Double
    INFINITY = 2.52E+30
    n = UBound(idx)
    pDepth = cur_depth
    pn_dimension = n_dimension
    If n = 1 Then
        pisLeaf = 1
        pPtr = idx(1)
        Set pChild_L = Nothing
        Set pChild_R = Nothing
        Exit Sub
    End If

    pSplitDim = cur_depth Mod n_dimension + 1
    ReDim xk(1 To n)
    For i = 1 To n
        xk(i) = x(idx(i), pSplitDim)
    Next i
    If n > 1 Then Call modMath.Sort_Quick_A(xk, 1, n, idx, 0)
    i = n \ 2 + 1
    pSplitVal = xk(i)
    pPtr = idx(i)
    
    'Extra information to be used in filtering algorithm
    node_size(pPtr) = n
    If dist_type = "EUCLIDEAN" Then
        For k = 1 To pn_dimension
            tmp_y = 0
            For i = 1 To n
                tmp_y = tmp_y + x(idx(i), k)
            Next i
            node_wgtcenter(pPtr, k) = tmp_y
        Next k
    ElseIf dist_type = "MANHATTAN" Then
        node_wgtcenter(pPtr) = idx
    End If
    For k = 1 To pn_dimension
        tmp_min = INFINITY
        tmp_max = -INFINITY
        For i = 1 To n
            tmp_x = x(idx(i), k)
            If tmp_x < tmp_min Then tmp_min = tmp_x
            If tmp_x > tmp_max Then tmp_max = tmp_x
        Next i
        node_min(pPtr, k) = tmp_min
        node_max(pPtr, k) = tmp_max
    Next k
    
    n1 = 0: ReDim idx1(1 To n)
    n2 = 0: ReDim idx2(1 To n)
    For i = 1 To n
        j = idx(i)
        If j <> pPtr Then
            If xk(i) < pSplitVal Then
                n1 = n1 + 1
                idx1(n1) = j
            Else
                n2 = n2 + 1
                idx2(n2) = j
            End If
        End If
    Next i
    If n1 > 0 Then
        ReDim Preserve idx1(1 To n1)
        Set pChild_L = New ckdTree
        Call pChild_L.z_kMean_Build_Tree_Recursive(x, idx1, cur_depth + 1, n_dimension, _
                node_size, node_wgtcenter, node_min, node_max, dist_type)
    Else
        Set pChild_L = Nothing
    End If

    If n2 > 0 Then
        ReDim Preserve idx2(1 To n2)
        Set pChild_R = New ckdTree
        Call pChild_R.z_kMean_Build_Tree_Recursive(x, idx2, cur_depth + 1, n_dimension, _
                node_size, node_wgtcenter, node_min, node_max, dist_type)
    Else
        Set pChild_R = Nothing
    End If
    
    Erase xk, idx1, idx2
End Sub



'==============================================================
'For a given set of centroids, find associated centroid to each data point,
'and recalculate centroids' postions and sizes.
'==============================================================
Sub kMean_Assign_Center(x() As Double, node_size() As Long, node_wgtcenter As Variant, _
            node_min() As Double, node_max() As Double, _
            centroids() As Double, centroids_size() As Long, x_idx() As Long, x_dist() As Double, _
            Optional dist_type As String = "EUCLIDEAN", _
            Optional use_kdTree As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, n_centroids As Long, n_dimension As Long
Dim CandidateSet() As Long, iArr() As Long
Dim centroids_wgtcenter As Variant, xi() As Double
Dim kT1 As ckdTree
    n = UBound(x, 1)
    n_centroids = UBound(centroids, 1)
    n_dimension = UBound(centroids, 2)
    ReDim x_idx(1 To n)
    ReDim CandidateSet(1 To n_centroids)
    ReDim centroids_size(1 To n_centroids)
    If dist_type = "EUCLIDEAN" Then
        ReDim centroids_wgtcenter(1 To n_centroids, 1 To n_dimension)
    ElseIf dist_type = "MANHATTAN" Then
        ReDim centroids_wgtcenter(1 To n_centroids)
        ReDim iArr(0 To 0)
        For k = 1 To n_centroids
            centroids_wgtcenter(k) = iArr
        Next k
    End If
    
    If use_kdTree = True Then
        Set kT1 = New ckdTree
        Call kT1.Build_Tree(centroids)
    End If

    Call z_kMean_Filter(x, node_size, node_wgtcenter, node_min, node_max, _
        centroids, centroids_size, centroids_wgtcenter, CandidateSet, _
        x_idx, n_centroids, n_dimension, dist_type, kT1, use_kdTree, x_dist)
    
    If use_kdTree = True Then
        Call kT1.Reset
        Set kT1 = Nothing
    End If
    
    If dist_type = "EUCLIDEAN" Then
        For k = 1 To n_centroids
            If centroids_size(k) > 0 Then
                For j = 1 To n_dimension
                    centroids(k, j) = centroids_wgtcenter(k, j) / centroids_size(k)
                Next j
            End If
        Next k
    ElseIf dist_type = "MANHATTAN" Then
        For k = 1 To n_centroids
            If centroids_size(k) > 0 Then
                iArr = centroids_wgtcenter(k)
                For j = 1 To n_dimension
                    Call modMath.Filter_Array(x, xi, iArr, j)
                    centroids(k, j) = modMath.fmedian(xi)
                Next j
            End If
        Next k
    End If
End Sub


Sub z_kMean_Filter(x() As Double, node_size() As Long, node_wgtcenter As Variant, _
            node_min() As Double, node_max() As Double, _
            centroids() As Double, centroids_size() As Long, _
            centroids_wgtcenter As Variant, CandidateSet() As Long, _
            x_idx() As Long, n_centroids As Long, n_dimension As Long, dist_type As String, _
            kT1 As ckdTree, use_kdTree As Boolean, x_dist() As Double)
Dim i As Long, j As Long, z As Long, z_nearest As Long
Dim tmp_x As Double, tmp_min As Double
Dim Z_set() As Long
Dim dist1 As Double, dist2 As Double
Dim v() As Double, u() As Double

    Z_set = CandidateSet    'Inherit candidate set from parent

    'vector represeting current node
    ReDim v(1 To n_dimension)
    For j = 1 To n_dimension
        v(j) = x(pPtr, j)
    Next j

    'Find nearest centroid to current node
    z_nearest = Nearest_Centroid(v, centroids, Z_set, n_centroids, n_dimension, dist_type, use_kdTree, kT1, x_dist)

    If pisLeaf = 1 Then
        x_idx(pPtr) = z_nearest
        centroids_size(z_nearest) = centroids_size(z_nearest) + 1
        Call Update_Centroid_Wgt(centroids_wgtcenter, z_nearest, dist_type, n_dimension, x)
    Else
        i = 0
        'Test if other candidates are farther than z_nearest
        ReDim u(1 To n_dimension)
        For z = 1 To n_centroids
            If z <> z_nearest And Z_set(z) = 0 Then
                For j = 1 To n_dimension
                    If centroids(z, j) < centroids(z_nearest, j) Then
                        u(j) = node_min(pPtr, j)
                    Else
                        u(j) = node_max(pPtr, j)
                    End If
                Next j
                dist1 = Dist2Centroid(u, centroids, z, dist_type, n_dimension)
                dist2 = Dist2Centroid(u, centroids, z_nearest, dist_type, n_dimension)
                If dist1 >= dist2 Then Z_set(z) = -1
            End If
            If Z_set(z) = 0 Then i = i + 1
        Next z
        If i = 1 Then
            'if only one candidate left, then assign all children to z_nearest
            Call z_Assign_Children_to_Centroid(z_nearest, x_idx)
            centroids_size(z_nearest) = centroids_size(z_nearest) + node_size(pPtr)
            Call Update_Centroid_Wgt_Group(centroids_wgtcenter, z_nearest, dist_type, n_dimension, node_wgtcenter)
        Else
            'otherwise assign only itself to z_nearest and recurse on children
            x_idx(pPtr) = z_nearest
            centroids_size(z_nearest) = centroids_size(z_nearest) + 1
            Call Update_Centroid_Wgt(centroids_wgtcenter, z_nearest, dist_type, n_dimension, x)
            
            If Not pChild_L Is Nothing Then Call pChild_L.z_kMean_Filter(x, _
                        node_size, node_wgtcenter, node_min, node_max, _
                        centroids, centroids_size, centroids_wgtcenter, _
                        Z_set, x_idx, n_centroids, n_dimension, dist_type, kT1, use_kdTree, x_dist)
            If Not pChild_R Is Nothing Then Call pChild_R.z_kMean_Filter(x, _
                        node_size, node_wgtcenter, node_min, node_max, _
                        centroids, centroids_size, centroids_wgtcenter, _
                        Z_set, x_idx, n_centroids, n_dimension, dist_type, kT1, use_kdTree, x_dist)
        End If
    End If
End Sub

'Distance from data v() to the z-th centroid
Private Function Dist2Centroid(v() As Double, centroids() As Double, z As Long, dist_type As String, n_dimension As Long)
Dim j As Long
    Dist2Centroid = 0
    If dist_type = "EUCLIDEAN" Then
        For j = 1 To n_dimension
            Dist2Centroid = Dist2Centroid + (v(j) - centroids(z, j)) ^ 2
        Next j
    ElseIf dist_type = "MANHATTAN" Then
        For j = 1 To n_dimension
            Dist2Centroid = Dist2Centroid + Abs(v(j) - centroids(z, j))
        Next j
    End If
End Function

'Find centroid that is nearest to v(), examining only active candidates
Private Function Nearest_Centroid(v() As Double, centroids() As Double, Z_set() As Long, _
        n_centroids As Long, n_dimension As Long, _
        dist_type As String, use_kdTree As Boolean, kT1 As ckdTree, x_dist() As Double) As Long
Dim z As Long
Dim tmp_x As Double, tmp_min As Double
    If use_kdTree = True Then
        Nearest_Centroid = kT1.NN_Search_Speed(v, centroids, pPtr, x_dist, dist_type)
    Else
        tmp_min = 2.52E+30
        For z = 1 To n_centroids
            If Z_set(z) = 0 Then
                If x_dist(pPtr, z) < 0 Then x_dist(pPtr, z) = Dist2Centroid(v, centroids, z, dist_type, n_dimension)
                tmp_x = x_dist(pPtr, z)
                'tmp_x = Dist2Centroid(v, centroids, z, dist_type, n_dimension)
                If tmp_x < tmp_min Then
                    tmp_min = tmp_x
                    Nearest_Centroid = z
                End If
            End If
        Next z
    End If
End Function


'Assign current data to centroid z_nearest and update z_nearest's data
Private Sub Update_Centroid_Wgt(centroids_wgtcenter As Variant, z_nearest As Long, _
        dist_type As String, n_dimension As Long, x() As Double)
Dim j As Long
Dim iArr() As Long
    If dist_type = "EUCLIDEAN" Then
        For j = 1 To n_dimension
            centroids_wgtcenter(z_nearest, j) = centroids_wgtcenter(z_nearest, j) + x(pPtr, j)
        Next j
    ElseIf dist_type = "MANHATTAN" Then
        iArr = centroids_wgtcenter(z_nearest)
        Call Append_1D(iArr, pPtr)
        centroids_wgtcenter(z_nearest) = iArr
    End If
End Sub

'Assign all children to centroid z_nearest and update z_nearest's data
Private Sub Update_Centroid_Wgt_Group(centroids_wgtcenter As Variant, z_nearest As Long, _
        dist_type As String, n_dimension As Long, node_wgtcenter As Variant)
Dim j As Long
Dim iArr() As Long
    If dist_type = "EUCLIDEAN" Then
        For j = 1 To n_dimension
            centroids_wgtcenter(z_nearest, j) = centroids_wgtcenter(z_nearest, j) + node_wgtcenter(pPtr, j)
        Next j
    ElseIf dist_type = "MANHATTAN" Then
        iArr = centroids_wgtcenter(z_nearest)
        Call Append_1D_vec(iArr, node_wgtcenter(pPtr))
        centroids_wgtcenter(z_nearest) = iArr
    End If
End Sub

'Assign all descendants to cluster index k
Sub z_Assign_Children_to_Centroid(k As Long, x_idx() As Long)
    x_idx(pPtr) = k
    If Not pChild_L Is Nothing Then Call pChild_L.z_Assign_Children_to_Centroid(k, x_idx)
    If Not pChild_R Is Nothing Then Call pChild_R.z_Assign_Children_to_Centroid(k, x_idx)
End Sub

Private Sub Append_1D(iArr() As Long, k As Long)
Dim n As Long
    n = UBound(iArr, 1)
    If n = 0 Then
        ReDim iArr(1 To 1)
    Else
        ReDim Preserve iArr(1 To n + 1)
    End If
    iArr(n + 1) = k
End Sub

Private Sub Append_1D_vec(iArr() As Long, idx As Variant)
Dim i As Long, n As Long, m As Long
    n = UBound(iArr, 1)
    m = UBound(idx, 1)
    If n = 0 Then
        ReDim iArr(1 To m)
    Else
        ReDim Preserve iArr(1 To n + m)
    End If
    For i = 1 To m
        iArr(n + i) = idx(i)
    Next i
End Sub
