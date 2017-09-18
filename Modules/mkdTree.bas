Attribute VB_Name = "mkdTree"
Option Explicit

'============================
'kd-tree data structure
'=============================
'Mainly used in speeding up k-nearest neighbor search
'N-observations of D-dimensional vectors are stored in a tree, represented as
'an array kdTree(1 to N, 1 to 7), data from 1 to 7 represents
'1: node index
'2: left-child
'3: right-child
'4: depth of a node
'5: splitting axis of node
'6: parent node
'7: kdtree(i,7)=j means datapoint i is saved in the j-th position of kdtree
'=============================


'=== Build k-d Tree from data points x()
'Input: x(1 to N, 1 to D), N datapoints of D dimensional data
'Output: kdtree(1 to N, 1 to 7)
'Output: tree(i,7)=j, datapoint i is saved in the j-th position of kdtree
Function kdtree(x() As Double) As Long()
Dim i As Long, j As Long, k As Long, n As Long, n_raw As Long
Dim tree() As Long, tree2() As Long, pos_index() As Long
    n_raw = UBound(x, 1)
    
    ReDim tree(1 To 5, 0 To 0)
    ReDim pos_index(1 To n_raw)
    For i = 1 To n_raw
        pos_index(i) = i
    Next i
    
    n = kdtree_recursive(x, pos_index, tree)
    
    ReDim tree2(1 To n_raw, 1 To 7)
    For i = 1 To n_raw
        For j = 1 To 5
            tree2(i, j) = tree(j, i)
        Next j
        tree2(tree(1, i), 7) = i
    Next i
    Erase tree
    
    'Find parent nodes
    For i = 1 To n_raw
        If tree2(i, 2) <> -1 Then tree2(tree2(tree2(i, 2), 7), 6) = tree2(i, 1)
        If tree2(i, 3) <> -1 Then tree2(tree2(tree2(i, 3), 7), 6) = tree2(i, 1)
    Next i
    tree2(1, 6) = -1
    
    kdtree = tree2
    Erase tree, tree2, pos_index
End Function


Private Function kdtree_recursive(x() As Double, pos_index() As Long, tree() As Long, Optional depth As Long = 0) As Long
Dim i As Long, j As Long, k As Long, n As Long, n_raw As Long, n_dimension As Long
Dim median_pos As Long, axis As Long
Dim x_sorted() As Double, x1() As Double, x2() As Double
Dim pos_sorted() As Long, pos1() As Long, pos2() As Long

    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    axis = depth Mod n_dimension + 1
    median_pos = n_raw \ 2 + 1
    
    'Reach a leaf
    If n_raw = 1 Then
        n = UBound(tree, 2) + 1
        ReDim Preserve tree(1 To 5, 0 To n)
        kdtree_recursive = pos_index(1)
        tree(1, n) = pos_index(1)
        tree(2, n) = -1
        tree(3, n) = -1
        tree(4, n) = depth
        tree(5, n) = axis
        Exit Function
    End If
    
    Call Sort_by_axis(x, x_sorted, pos_index, pos_sorted, axis)
    kdtree_recursive = pos_sorted(median_pos)
    
    Call Array_Partition(x_sorted, x1, x2, pos_sorted, pos1, pos2, median_pos)
    
    n = UBound(tree, 2) + 1
    ReDim Preserve tree(1 To 5, 0 To n)
    tree(1, n) = pos_sorted(median_pos)
    tree(4, n) = depth
    tree(5, n) = axis
    tree(2, n) = kdtree_recursive(x1, pos1, tree, depth + 1)
    If UBound(x2, 1) > 0 Then
        tree(3, n) = kdtree_recursive(x2, pos2, tree, depth + 1)
    Else
        tree(3, n) = -1
    End If
    Erase x1, x2, pos1, pos2, x_sorted, pos_sorted
End Function



'=== Find k nearest neighbor for every data point in x()
'Input: x(1 to N, 1 to D), N datapoints of D dimensional data
Sub kNN_All(k_idx() As Long, k_dist() As Double, x() As Double, k As Long, Optional kth_only As Long = 1, Optional dist_type As String = "EUCLIDEAN")
Dim i As Long, j As Long, n_raw As Long
Dim tree() As Long
Dim Output() As Double, y() As Double
Dim kNeighbor() As Long
Dim kDist() As Double

n_raw = UBound(x, 1)

tree = kdtree(x)
If kth_only = 1 Then
    ReDim k_idx(1 To n_raw)
    ReDim k_dist(1 To n_raw)
    For i = 1 To n_raw
        If i Mod 250 = 0 Then
            DoEvents
            Application.StatusBar = "kNN (kdtree): " & i & "/" & n_raw
        End If
        Call kNNSearch(kNeighbor, kDist, i, k, x, tree, dist_type)
        k_idx(i) = kNeighbor(k)
        k_dist(i) = kDist(k)
    Next i
Else
    ReDim k_idx(1 To n_raw, 1 To k)
    ReDim k_dist(1 To n_raw, 1 To k)
    For i = 1 To n_raw
        If i Mod 250 = 0 Then
            DoEvents
            Application.StatusBar = "kNN (kdtree): " & i & "/" & n_raw
        End If
        Call kNNSearch(kNeighbor, kDist, i, k, x, tree, dist_type)
        For j = 1 To k
            k_idx(i, j) = kNeighbor(j)
            k_dist(i, j) = kDist(j)
        Next j
    Next i
End If
Erase tree
Application.StatusBar = False
End Sub


'=== Search for k-Nearest Neighbors of tgt node
'Output: kNeighbor(1 to k) - node lable of neighbors
'Output: kDist(1 to k) - distance from tgt node
Sub kNNSearch(kNeighbor() As Long, kDist() As Double, tgt As Long, k_nearest As Long, x() As Double, tree() As Long, Optional dist_type As String = "EUCLIDEAN")
Dim i As Long, j As Long, u As Long, v As Long, vparent As Long
Dim tmp_x As Double, INFINITY As Double
Dim visited() As Long
Dim vstack As Collection

INFINITY = Exp(70)
Set vstack = New Collection
ReDim visited(1 To UBound(tree, 1))
ReDim kNeighbor(1 To k_nearest)
ReDim kDist(1 To k_nearest)
visited(tgt) = 1 'so that target itself won't be selected
For i = 1 To k_nearest
    kNeighbor(i) = -1
    kDist(i) = INFINITY
Next i

Call Add_to_Stack(kNeighbor, kDist, tgt, k_nearest, x, tree, tree(1, 1), visited, vstack, dist_type)

Do While vstack.count > 0
    With vstack
        v = .Item(.count)
        .Remove .count
    End With
    If v = tree(1, 1) Then Exit Do
    vparent = tree(tree(v, 7), 6)
    If hyper_intersect(tgt, vparent, kDist(k_nearest), x, tree) = 1 Then
        u = tree(tree(vparent, 7), 2)
        If u = v Then u = tree(tree(vparent, 7), 3)
        If u <> -1 Then
            If visited(u) = 0 Then Call Add_to_Stack(kNeighbor, kDist, tgt, k_nearest, x, tree, u, visited, vstack, dist_type)
        End If
    End If
Loop
Set vstack = Nothing
Erase visited

If dist_type = "EUCLIDEAN" Then
    For i = 1 To k_nearest
        kDist(i) = Sqr(kDist(i))
    Next i
End If
End Sub

'Starting from v_start, traverse down the tree until a leaf is reached
Private Sub Add_to_Stack(kNeighbor() As Long, kDist() As Double, tgt As Long, k_nearest As Long, x() As Double, tree() As Long, _
            v_start As Long, visited() As Long, vstack As Collection, Optional dist_type As String = "EUCLIDEAN")
Dim v As Long, i As Long
v = v_start
vstack.Add v
Do
    If visited(v) = 0 Then
        visited(v) = 1
        Call Queue_Eject(kNeighbor, kDist, v, node2node_dist(v, tgt, x, dist_type), k_nearest)
    End If
    If isLeaf(v, tree) = 1 Then Exit Do
    i = hyper_LR(tgt, v, x, tree)
    If tree(tree(v, 7), i) = -1 Then
        If i = 3 Then
            If tree(tree(v, 7), 2) = -1 Then Exit Do
            v = tree(tree(v, 7), 2)
            vstack.Add v
        End If
    Else
        v = tree(tree(v, 7), i)
        vstack.Add v
    End If
Loop
End Sub


Private Sub Queue_Eject(kNeighbor() As Long, kDist() As Double, v As Long, r As Double, k As Long)
If r < kDist(k) Then
    ReDim Preserve kNeighbor(1 To k + 1)
    ReDim Preserve kDist(1 To k + 1)
    kNeighbor(k + 1) = v
    kDist(k + 1) = r
    Call modMath.Sort_Quick_A(kDist, 1, k + 1, kNeighbor, 0)
    ReDim Preserve kNeighbor(1 To k)
    ReDim Preserve kDist(1 To k)
End If
End Sub


'Input: hypersphere of radius r (squared) centered at tgt node
'Input: tree node v
'Output: 1 - Hypersphere intersects hyperplane
'        2 - Does not intersect, hypesphere lies on the left side of hyperplane
'        3 - Does not intersect, hypesphere lies on the right side of hyperplane
Private Function hyper_intersect(tgt As Long, v As Long, r As Double, x() As Double, tree() As Long) As Long
Dim axis As Long
    axis = tree(tree(v, 7), 5)
    If r >= ((x(v, axis) - x(tgt, axis)) ^ 2) Then
        hyper_intersect = 1
    Else
        If x(tgt, axis) <= x(v, axis) Then
            hyper_intersect = 2
        Else
            hyper_intersect = 3
        End If
    End If
End Function

'Determines if tgt node lies on the left or right side of v's hyperplane
'Output: 2 - Left, 3 - Right
Private Function hyper_LR(tgt As Long, v As Long, x() As Double, tree() As Long) As Long
Dim axis As Long
    axis = tree(tree(v, 7), 5)
    If x(tgt, axis) <= x(v, axis) Then
        hyper_LR = 2
    Else
        hyper_LR = 3
    End If
End Function


Private Function NodeInfo(node As Long, tree() As Long, Optional strinfo As String = "LEFT") As Long
Dim i As Long
i = tree(node, 7)
strinfo = UCase(strinfo)
If strinfo = "LEFT" Then
    NodeInfo = tree(i, 2)
ElseIf strinfo = "RIGHT" Then
    NodeInfo = tree(i, 3)
ElseIf strinfo = "DEPTH" Then
    NodeInfo = tree(i, 4)
ElseIf strinfo = "AXIS" Then
    NodeInfo = tree(i, 5)
ElseIf strinfo = "PARENT" Then
    NodeInfo = tree(i, 6)
Else
    Debug.Print "NodeInfo: " & strinfo & " is not a valid field."
    NodeInfo = -1
End If
End Function


Private Function node2node_dist(u As Long, v As Long, x() As Double, Optional strType As String = "EUCLIDEAN") As Double
Dim i As Long
    node2node_dist = 0
    If strType = "EUCLIDEAN" Then
        For i = 1 To UBound(x, 2)
            node2node_dist = node2node_dist + (x(u, i) - x(v, i)) ^ 2
        Next i
    ElseIf strType = "MAXNORM" Then
        For i = 1 To UBound(x, 2)
            If Abs(x(u, i) - x(v, i)) > node2node_dist Then node2node_dist = Abs(x(u, i) - x(v, i))
        Next i
    End If
End Function

Private Function x2node_dist(x_tgt() As Double, v As Long, x() As Double, Optional strType As String = "EUCLIDEAN") As Double
Dim i As Long
    x2node_dist = 0
    If strType = "EUCLIDEAN" Then
        For i = 1 To UBound(x, 2)
            x2node_dist = x2node_dist + (x_tgt(i) - x(v, i)) ^ 2
        Next i
    ElseIf strType = "MAXNORM" Then
        For i = 1 To UBound(x, 2)
            If Abs(x_tgt(i) - x(v, i)) > x2node_dist Then x2node_dist = Abs(x_tgt(i) - x(v, i))
        Next i
    End If
End Function

Private Function isRoot(node As Long, tree() As Long) As Long
    isRoot = 0
    If tree(1, 1) = node Then isRoot = 1
End Function

Private Function isLeaf(node As Long, tree() As Long) As Long
    isLeaf = 0
    If tree(tree(node, 7), 2) = -1 And tree(tree(node, 7), 3) = -1 Then isLeaf = 1
End Function





'Extract the subtree that roots at the node-th data point
Private Function subtree(tree() As Long, node As Long) As Long()
Dim i As Long, j As Long, k As Long, n As Long, n_raw As Long
Dim tree2() As Long, subtree_index() As Long

If isRoot(node, tree) = 1 Then 'node is root, just copy the whole tree
    subtree = tree
    Exit Function
End If

n_raw = UBound(tree, 1)

ReDim subtree_index(0 To 0)
n = subtree_recursive(tree, node, subtree_index)

n = UBound(subtree_index, 1)

ReDim tree2(1 To n, 1 To 6)
For i = 1 To n
    k = subtree_index(i)
    For j = 1 To 6
        tree2(i, j) = tree(k, j)
    Next j
Next i

subtree = tree2
End Function


Private Function subtree_recursive(tree() As Long, node As Long, subtree_index() As Long) As Long
Dim i As Long, j As Long, k As Long, n As Long
If node = -1 Then 'node is nothing
    Exit Function
ElseIf isLeaf(node, tree) = 1 Then 'node is leaf, add leaf and stop recursion
    n = UBound(subtree_index, 1) + 1
    ReDim Preserve subtree_index(0 To n)
    subtree_index(n) = tree(node, 7)
Else
    k = tree(node, 7)
    n = UBound(subtree_index, 1) + 1
    ReDim Preserve subtree_index(0 To n)
    subtree_index(n) = k
    If tree(k, 2) > -1 Then i = subtree_recursive(tree, tree(k, 2), subtree_index)
    If tree(k, 3) > -1 Then j = subtree_recursive(tree, tree(k, 3), subtree_index)
End If
End Function





Private Sub Array_Partition(x As Variant, x1 As Variant, x2 As Variant, _
        pos_index() As Long, pos1() As Long, pos2() As Long, k As Long)
Dim i As Long, j As Long, n As Long, n_dimension As Long
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    ReDim x1(1 To k - 1, 1 To n_dimension)
    ReDim pos1(1 To k - 1)
    For i = 1 To k - 1
        pos1(i) = pos_index(i)
        For j = 1 To n_dimension
            x1(i, j) = x(i, j)
        Next j
    Next i
    
    If n > 2 Then
        ReDim x2(1 To n - k, 1 To n_dimension)
        ReDim pos2(1 To n - k)
        For i = 1 To n - k
            pos2(i) = pos_index(k + i)
            For j = 1 To n_dimension
                x2(i, j) = x(k + i, j)
            Next j
        Next i
    Else
        ReDim x2(0 To 0, 1 To n_dimension)
        ReDim pos2(0 To 0)
    End If
End Sub

Private Sub Sort_by_axis(x As Variant, x_sorted As Variant, pos_index() As Long, pos_sorted() As Long, axis As Long)
Dim i As Long, j As Long, k As Long, n_raw As Long, n_dimension As Long
Dim tmpVec As Variant
Dim sort_index() As Long

    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    ReDim tmpVec(1 To n_raw)
    For i = 1 To n_raw
        tmpVec(i) = x(i, axis)
    Next i
    Call modMath.Sort_Quick_A(tmpVec, 1, n_raw, sort_index, 1)
    
    ReDim x_sorted(1 To n_raw, 1 To n_dimension)
    ReDim pos_sorted(1 To n_raw)
    For i = 1 To n_raw
        j = sort_index(i)
        For k = 1 To n_dimension
            x_sorted(i, k) = x(j, k)
        Next k
        pos_sorted(i) = pos_index(j)
    Next i

End Sub


''Slower implementation
''=== Search for k-Nearest Neighbors of tgt node
''Output: kNeighbor(1 to k) - node lable of neighbors
''Output: kDist(1 to k) - distance from tgt node
'Sub kNNSearch(kNeighbor() As Long, kDist() As Double, tgt As Long, k_nearest As Long, x() As Double, tree() As Long)
'Dim i As Long, j As Long, v As Long, n_raw As Long, vparent As Long
'Dim tmp_x As Double, INFINITY As Double
'Dim vStack As Collection
'Dim visited() As Long
'Dim y() As Double
'
'INFINITY = Exp(70)
'n_raw = UBound(tree, 1)
'
'ReDim visited(1 To n_raw)
'visited(tgt) = 1 'so that target itself won't be selected
'
'ReDim kNeighbor(1 To k_nearest)
'ReDim kDist(1 To k_nearest)
'For i = 1 To k_nearest
'    kNeighbor(i) = -1
'    kDist(i) = INFINITY
'Next i
'
'v = tree(1, 1)
'Do
'    If visited(v) = 0 Then
'        visited(v) = 1
'        Call Queue_Eject(kNeighbor, kDist, v, node2node_dist(v, tgt, x), k_nearest)
'    End If
'    If isLeaf(v, tree) = 1 Then Exit Do
'    j = hyper_LR(tgt, v, x, tree)
'    If tree(tree(v, 7), j) <> -1 Then
'        v = tree(tree(v, 7), j)
'    Else
'        Exit Do
'    End If
'Loop
'
'Set vStack = New Collection
'vStack.Add tree(1, 1)
'
'Do While vStack.count > 0
'
'    With vStack
'        i = .count
'        v = .Item(i)
'        .Remove i
'    End With
'
'    If visited(v) = 0 Then
'        visited(v) = 1
''    If v <> tgt Then
'        'Call Queue_Eject(kNeighbor, kDist, v, node2node_dist(v, tgt, x), k_nearest)
'        tmp_x = node2node_dist(v, tgt, x)
'        If tmp_x < kDist(k_nearest) Then
'            ReDim Preserve kNeighbor(1 To k_nearest + 1)
'            ReDim Preserve kDist(1 To k_nearest + 1)
'            kNeighbor(k_nearest + 1) = v
'            kDist(k_nearest + 1) = tmp_x
'            Call Sort_Quick_A(kDist, 1, k_nearest + 1, kNeighbor, 0)
'            ReDim Preserve kNeighbor(1 To k_nearest)
'            ReDim Preserve kDist(1 To k_nearest)
'        End If
'    End If
'
'    i = tree(v, 7)
'    If tree(i, 2) <> -1 Or tree(i, 3) <> -1 Then
'        j = hyper_intersect(tgt, v, kDist(k_nearest), x, tree)
'        If j = 1 Then
'            'Hypersphere intersects hyperplane, need to search both children
'            If tree(i, 2) <> -1 Then vStack.Add tree(i, 2)
'            If tree(i, 3) <> -1 Then vStack.Add tree(i, 3)
'        Else
'            'Does not intersect, search only the side where it lies in.
'            If tree(i, j) <> -1 Then vStack.Add tree(i, j)
'        End If
'    End If
'Loop
'
'Set vStack = Nothing
'Erase visited
'For i = 1 To k_nearest
'    kDist(i) = Sqr(kDist(i))
'Next i
'End Sub
