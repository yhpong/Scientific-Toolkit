VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGraphAlgo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************
'This module is built for undirected graph, represented
'by edge list as the main data structure
'pEdgeList(): a N x 2 array storing the start and end node index of each edge
'pEdgeDist(): a N x1 array storing the length of each edge
'***********************************************
'Requires: cqtree, cqtree_point, cHeap
'Requires: cPFMG_extFaceLinkRec, _Graph, _lcnode, _ListColl, _Node, _Stack, _VertexRec
'Requires: gp, gp_Embed
'***********************************************

Private pSize As Long           'number of nodes
Private pn_edge As Long         'number of edges
Private pEdgeList() As Long     'edge list
Private pEdgeDist() As Double   'edge length
Private pnode_idx() As Long     'integer index of each node
Private pnode_pos() As Double   'Nx2 array storing the 2D-layout of graph
Private pnode_degree() As Long  'Nx1 array storing degree of each node
Private pnode_degree_wgt() As Double    'Nx1 array storing degree (weighted) of each node
Private pnode_closeness() As Double     'Nx1 arary storing closeness of each node
Private pnode_eigen() As Double 'Nx1 array storing eigen vector centrality of each node
Private pnode_next() As Long    'Nx1 mapping of current set of nodes to nodes in a coarsened graph


Public Property Get Size() As Long
Size = pSize
End Property

Public Property Get n_edge() As Long
n_edge = pn_edge
End Property

Public Property Get EdgeList() As Long()
EdgeList = pEdgeList
End Property

Public Property Get EdgeDist() As Double()
EdgeDist = pEdgeDist
End Property

Public Property Get node_pos() As Double()
node_pos = pnode_pos
End Property

Public Property Let node_pos(x() As Double)
pnode_pos = x
End Property

Public Property Get node_degree() As Long()
node_degree = pnode_degree
End Property

Public Property Get node_degree_wgt() As Double()
node_degree_wgt = pnode_degree_wgt
End Property

Public Property Get node_idx() As Long()
node_idx = pnode_idx
End Property

Public Property Get node_next() As Long()
node_next = pnode_next
End Property

Public Property Let node_next(x() As Long)
pnode_next = x
End Property

Public Property Get node_closeness() As Double()
node_closeness = pnode_closeness
End Property

Public Property Get node_eigen() As Double()
node_eigen = pnode_eigen
End Property


'=== Build Minimum Spanning Tree from a pairwise distance matrix
'Input: Symmetric Distance(N x N) matrix
Sub MST_Build(distance() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_pairs As Long
Dim temp1 As Long, temp2 As Long
Dim parent() As Long, sort_index() As Long
Dim tmp_x As Double

pSize = UBound(distance, 1)
n_pairs = pSize * (pSize - 1) / 2
ReDim pnode_idx(1 To pSize)
For i = 1 To pSize
    pnode_idx(i) = i
Next i

k = 0
ReDim d(1 To n_pairs)
For i = 1 To pSize - 1
    For j = i + 1 To pSize
        k = k + 1
        d(k) = distance(i, j)
    Next j
Next i

Application.StatusBar = "Building MST: Sorting distances...."
Call modMath.Sort_Quick_A(d, 1, n_pairs, sort_index)

pn_edge = 0
ReDim parent(1 To pSize)
ReDim pEdgeList(1 To pSize - 1, 1 To 2)
ReDim pEdgeDist(1 To pSize - 1)
For k = 1 To n_pairs
    
    Call condense2sq(sort_index(k), pSize, m, n)
    
    If parent(m) <> parent(n) Or (parent(m) + parent(n)) = 0 Then
    
        'Union the pair into the same parent
        temp1 = parent(m)
        temp2 = parent(n)
        parent(m) = m
        parent(n) = m
        
        For i = 1 To pSize
            If (parent(i) = temp1 Or parent(i) = temp2) And parent(i) > 0 Then
                parent(i) = m
            End If
        Next i

        pn_edge = pn_edge + 1
        pEdgeList(pn_edge, 1) = m
        pEdgeList(pn_edge, 2) = n
        pEdgeDist(pn_edge) = d(k)
        If pn_edge = (pSize - 1) Then Exit For
        
        DoEvents
        Application.StatusBar = "cMST.Build: " & pn_edge & " / " & (pSize - 1)
        
    End If

Next k

If pn_edge <> (pSize - 1) Then
    Debug.Print "MST_Build Fail: Num of edges is not equal to N-1"
End If

Erase parent, sort_index, d
Application.StatusBar = False

Randomize
ReDim pnode_pos(1 To pSize, 1 To 2)
For i = 1 To pSize
    pnode_pos(i, 1) = (-0.5 + Rnd()) * Sqr(pSize)
    pnode_pos(i, 2) = (-0.5 + Rnd()) * Sqr(pSize)
Next i
End Sub


'=== Import graph from user supplied EdgeList, EdgeDist and node_pos
Sub Init(EdgeList() As Long, EdgeDist() As Double, node_pos() As Double)
Dim i As Long
    pn_edge = UBound(EdgeDist, 1)
    pSize = UBound(node_pos, 1)
    pEdgeList = EdgeList
    pEdgeDist = EdgeDist
    pnode_pos = node_pos
    ReDim pnode_idx(1 To pSize)
    For i = 1 To pSize
        pnode_idx(i) = i
    Next i
End Sub


'=== Release memory
Sub Reset()
pSize = 0
pn_edge = 0
Erase pEdgeList, pEdgeDist, pnode_pos, pnode_degree, pnode_closeness, pnode_eigen, pnode_idx, pnode_next
End Sub


'=== Create a duplicate of input graph g
Sub Copy(g As cGraphAlgo)
With g
    pSize = .Size
    pn_edge = .n_edge
    pEdgeList = .EdgeList
    pEdgeDist = .EdgeDist
    pnode_pos = .node_pos
    pnode_idx = .node_idx
    pnode_next = .node_next
    pnode_degree = .node_degree
    pnode_closeness = .node_closeness
    pnode_eigen = .node_eigen
End With
End Sub


'=== Return list of edges in format that can be printed as excel chart
Function Print_Edges()
Dim i As Long, j As Long, k As Long, n As Long
Dim vArr As Variant
ReDim vArr(1 To pn_edge * 3, 1 To 2)
k = 1
For n = 1 To pn_edge
    i = pEdgeList(n, 1)
    j = pEdgeList(n, 2)
    vArr(k, 1) = pnode_pos(i, 1)
    vArr(k, 2) = pnode_pos(i, 2)
    vArr(k + 1, 1) = pnode_pos(j, 1)
    vArr(k + 1, 2) = pnode_pos(j, 2)
    k = k + 3
Next n
Print_Edges = vArr
Erase vArr
End Function


'=== Arrange graph layout by force directed algorithm
'=== c1 is spring strength, c2 is sprint natural length, c3 is repulsive strength
Sub ForceDirectedLayout(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 500)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim dist() As Double, force() As Double, force_xy() As Double
Dim energy As Double, energy_prev As Double, temperature As Double

energy_prev = Exp(70)
temperature = c2
ReDim dist(1 To pSize, 1 To pSize)
ReDim force(1 To pSize, 1 To pSize)
For iterate = 1 To iter_max

    If iterate Mod 20 = 0 Then
        DoEvents
        Application.StatusBar = "ForceDirected: " & iterate & "/" & iter_max
    End If

    ReDim force_xy(1 To pSize, 1 To 2)

    For i = 1 To pSize - 1
        For j = i + 1 To pSize
            dist(i, j) = Sqr((pnode_pos(i, 1) - pnode_pos(j, 1)) ^ 2 + (pnode_pos(i, 2) - pnode_pos(j, 2)) ^ 2)
            If dist(i, j) > 0 Then force(i, j) = c3 / (dist(i, j) ^ 3)
        Next j
    Next i

    For k = 1 To pn_edge
        i = pEdgeList(k, 1)
        j = pEdgeList(k, 2)
        If i > j Then
            m = i
            i = j
            j = m
        End If
        tmp_x = dist(i, j)
        If tmp_x > 0 Then force(i, j) = force(i, j) - c1 * Log(tmp_x / c2) / tmp_x
    Next k

    For i = 1 To pSize - 1
        For j = i + 1 To pSize
            tmp_x = pnode_pos(i, 1) - pnode_pos(j, 1)
            tmp_y = pnode_pos(i, 2) - pnode_pos(j, 2)
            force_xy(i, 1) = force_xy(i, 1) + force(i, j) * tmp_x
            force_xy(i, 2) = force_xy(i, 2) + force(i, j) * tmp_y
            force_xy(j, 1) = force_xy(j, 1) - force(i, j) * tmp_x
            force_xy(j, 2) = force_xy(j, 2) - force(i, j) * tmp_y
        Next j
    Next i

    energy = 0
    For i = 1 To pSize
        tmp_y = Sqr(force_xy(i, 1) ^ 2 + force_xy(i, 2) ^ 2)
        tmp_x = min2(tmp_y, temperature)
        pnode_pos(i, 1) = pnode_pos(i, 1) + tmp_x * force_xy(i, 1) / tmp_y
        pnode_pos(i, 2) = pnode_pos(i, 2) + tmp_x * force_xy(i, 2) / tmp_y
        energy = energy + tmp_x
    Next i
    energy = energy / pSize
    If energy < energy_prev Then
        temperature = 1.05 * temperature
    Else
        temperature = 0.9 * temperature
    End If
    energy_prev = energy
    If energy < (c2 * 0.01) Then Exit For
    
Next iterate
Erase force, force_xy, dist
Application.StatusBar = False
End Sub


'=== Force Directed algorithm with Barnes-Hut acceleration
Sub ForceDirectedLayout_BarnesHut(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 500)
Dim i As Long, j As Long, k As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, tmp As Double
Dim energy As Double, energy_prev As Double, temperature As Double
Dim quadtree1 As cqtree
Dim force_xy() As Double

energy_prev = Exp(70)
temperature = c2
For iterate = 1 To iter_max
    
    If iterate Mod 50 = 0 Then
        DoEvents
        Application.StatusBar = "ForceDirected (Barnes-Hut): " & iterate & "/" & iter_max
    End If
    
    Set quadtree1 = New cqtree
    force_xy = quadtree1.NetForce(pnode_pos, c3)
    
    For k = 1 To pn_edge
        i = pEdgeList(k, 1)
        j = pEdgeList(k, 2)
        tmp_x = pnode_pos(i, 1) - pnode_pos(j, 1)
        tmp_y = pnode_pos(i, 2) - pnode_pos(j, 2)
        tmp = Sqr(tmp_x ^ 2 + tmp_y ^ 2)
        If tmp > 0 Then
            tmp = c1 * Log(tmp / c2) / tmp
            force_xy(i, 1) = force_xy(i, 1) - tmp * tmp_x
            force_xy(i, 2) = force_xy(i, 2) - tmp * tmp_y
            force_xy(j, 1) = force_xy(j, 1) + tmp * tmp_x
            force_xy(j, 2) = force_xy(j, 2) + tmp * tmp_y
        End If
    Next k
    
    energy = 0
    For i = 1 To pSize
        tmp = Sqr(force_xy(i, 1) ^ 2 + force_xy(i, 2) ^ 2)
        tmp_x = min2(tmp, temperature)
        pnode_pos(i, 1) = pnode_pos(i, 1) + tmp_x * force_xy(i, 1) / tmp
        pnode_pos(i, 2) = pnode_pos(i, 2) + tmp_x * force_xy(i, 2) / tmp
        energy = energy + tmp_x
    Next i
    energy = energy / pSize
    If energy < energy_prev Then
        temperature = 1.05 * temperature
    Else
        temperature = 0.9 * temperature
    End If
    energy_prev = energy
    If energy < (c2 * 0.01) Then Exit For
Next iterate

Set quadtree1 = Nothing
Erase force_xy
Application.StatusBar = False
End Sub

Private Function min2(x As Double, y As Double) As Double
min2 = x
If y < x Then min2 = y
End Function


'=== Force Directed algorithm with multi-level and Barnes-Hut acceleration
Sub ForceDirected_MultiLevel(Optional c1 As Double = 2, Optional c2 As Double = 1, Optional c3 As Double = 1, _
        Optional iter_max As Long = 300)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim g2 As cGraphAlgo, g_tmp As cGraphAlgo
Dim G_List As Collection
Dim node_pos() As Double

Set G_List = New Collection
Set g_tmp = New cGraphAlgo
Call g_tmp.Copy(Me)
Do
    Set g2 = New cGraphAlgo
    Call g2.Collapse(g_tmp)
    G_List.Add g_tmp
    If g2.Size <= 5 Then Exit Do
    Set g_tmp = New cGraphAlgo
    Call g_tmp.Copy(g2)
Loop
G_List.Add g2

i = 0
Do While G_List.count > 0
    DoEvents
    Application.StatusBar = "ForceDirected (Multilevel): " & G_List.count & "->1"
    i = i + 1
    With G_List
        Set g2 = .Item(.count)
        .Remove .count
    End With
    With g2
        If i > 1 Then Call .IntrapolateLayout(node_pos)
        Call .ForceDirectedLayout_BarnesHut(c1, c2, c3, iter_max)
        node_pos = .node_pos
        Call .Reset
    End With
Loop
pnode_pos = node_pos

'Release memory
Erase node_pos
Set g2 = Nothing
Set g_tmp = Nothing
Set G_List = Nothing
End Sub


'Given the layout of a coarsened graph, recover layout of the finer graph
'Input: node_pos() from a coarsened graph with layout already found
Sub IntrapolateLayout(node_pos() As Double)
Dim i As Long, j As Long
ReDim pnode_pos(1 To pSize, 1 To 2)
For i = 1 To pSize
    j = pnode_next(i)
    pnode_pos(i, 1) = node_pos(j, 1) + (-0.1 + 0.2 * Rnd())
    pnode_pos(i, 2) = node_pos(j, 2) + (-0.1 + 0.2 * Rnd())
Next i
End Sub


'=== Coarsen an input graph by collapsing edges until graph is half-size
'=== mapping of nodes is saved in .node_next attribute of input graph
Sub Collapse(g As cGraphAlgo)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim u As Long, v As Long, w As Long
Dim EdgeList() As Long, EdgeDist() As Double, node_pos() As Double
Dim node_next() As Long

With g
    n = .Size
    m = .n_edge
    EdgeList = .EdgeList
    EdgeDist = .EdgeDist
    node_pos = .node_pos
End With

If n = 2 Then
    Debug.Print "Only 2 nodes left. Cannot collapse."
    Exit Sub
End If

w = n
pSize = n
pn_edge = m
ReDim node_next(1 To n)
For k = pn_edge To 1 Step -1
    u = EdgeList(k, 1)
    v = EdgeList(k, 2)
    If u <= n And v <= n Then
        w = w + 1
        node_next(u) = w
        node_next(v) = w
        pSize = pSize - 1
        Call EdgeList_Relabel(EdgeList, u, v, w)
    ElseIf u > n And v <= n Then
        node_next(v) = u
        pSize = pSize - 1
        Call EdgeList_Relabel(EdgeList, u, v, u)
    ElseIf v > n And u <= n Then
        node_next(u) = v
        pSize = pSize - 1
        Call EdgeList_Relabel(EdgeList, u, v, v)
    End If
    If pSize = 2 Or pSize < (n \ 2) Then Exit For
Next k

For i = 1 To n
    If node_next(i) = 0 Then
        w = w + 1
        node_next(i) = w
    End If
Next i

For k = 1 To pn_edge
    u = EdgeList(k, 1)
    v = EdgeList(k, 2)
    If u <= n Then EdgeList(k, 1) = node_next(u)
    If v <= n Then EdgeList(k, 2) = node_next(v)
    EdgeList(k, 1) = EdgeList(k, 1) - n
    EdgeList(k, 2) = EdgeList(k, 2) - n
Next k
For i = 1 To n
    node_next(i) = node_next(i) - n
Next i

g.node_next = node_next
Call EdgeList_Purge(EdgeList)
pn_edge = UBound(EdgeList, 1)
pEdgeList = EdgeList

ReDim pnode_pos(1 To pSize, 1 To 2)
For i = 1 To n
    j = node_next(i)
    pnode_pos(j, 1) = node_pos(i, 1)
    pnode_pos(j, 2) = node_pos(i, 2)
Next i
End Sub

'=== Relabel u and v to w in EdgeList()
Private Sub EdgeList_Relabel(EdgeList() As Long, u As Long, v As Long, w As Long)
Dim i As Long, j As Long, k As Long, n As Long
n = UBound(EdgeList, 1)
For k = 1 To n
    If (EdgeList(k, 1) = u Or EdgeList(k, 1) = v) Then EdgeList(k, 1) = w
    If (EdgeList(k, 2) = u Or EdgeList(k, 2) = v) Then EdgeList(k, 2) = w
Next k
End Sub

'=== Remove duplicate edges and self-cycle from EdgeList()
Private Sub EdgeList_Purge(EdgeList() As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_edge As Long, isUnique As Long, count As Long
Dim newEdgeList() As Long
n_edge = UBound(EdgeList, 1)
count = 0
ReDim newEdgeList(1 To 2, 1 To n_edge)
For k = 1 To n_edge
    m = EdgeList(k, 1)
    n = EdgeList(k, 2)
    If m <> n Then
        isUnique = 1
        For i = 1 To count
            If (newEdgeList(1, i) = m And newEdgeList(2, i) = n) _
                Or (newEdgeList(1, i) = n And newEdgeList(2, i) = m) Then
                isUnique = 0
                Exit For
            End If
        Next i
        If isUnique = 1 Then
            count = count + 1
            newEdgeList(1, count) = m
            newEdgeList(2, count) = n
        End If
    End If
Next k
ReDim Preserve newEdgeList(1 To 2, 1 To count)
Call mTranspose(newEdgeList, EdgeList)
Erase newEdgeList
End Sub


'=== Find degree of each node
Sub Find_degree()
Dim i As Long, j As Long, k As Long
    ReDim pnode_degree(1 To pSize)
    For k = 1 To pn_edge
        i = pEdgeList(k, 1)
        j = pEdgeList(k, 2)
        pnode_degree(i) = pnode_degree(i) + 1
        pnode_degree(j) = pnode_degree(j) + 1
    Next k
End Sub


'=== Find weighted degree of each node
Sub Find_degree_wgt()
Dim i As Long, j As Long, k As Long
    ReDim pnode_degree_wgt(1 To pSize)
    For k = 1 To pn_edge
        i = pEdgeList(k, 1)
        j = pEdgeList(k, 2)
        pnode_degree_wgt(i) = pnode_degree_wgt(i) + 1# / pEdgeDist(k)
        pnode_degree_wgt(j) = pnode_degree_wgt(j) + 1# / pEdgeDist(k)
    Next k
End Sub

'=== Find closeness of each node
Sub Find_closeness()
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double
Dim dist() As Double
    ReDim pnode_closeness(1 To pSize)
    Call Dijkstra_Algorithm(dist)
    For i = 1 To pSize
        tmp_x = 0
        For j = 1 To pSize
            If i <> j Then tmp_x = tmp_x + dist(i, j)
        Next j
        pnode_closeness(i) = (pSize - 1) / tmp_x
    Next i
    Erase dist
End Sub


'=== Find shortest path between all pairs using Dijkstra's algorithm
Sub Dijkstra_Algorithm(dist() As Double)
Dim i As Long, j As Long, k As Long, n As Long, s As Long, u As Long, v As Long
Dim tmp_x As Double, d_min As Double, INFINITY As Double
Dim max_degree As Long
Dim neighbour() As Long, sort_index() As Long
Dim d() As Double, neighbour_dist() As Double
Dim q As cHeap

INFINITY = Exp(70)

max_degree = 0
Call Me.Find_degree
For i = 1 To pSize
    If pnode_degree(i) > max_degree Then max_degree = pnode_degree(i)
Next i

'First identify the neigbours of each node
ReDim neighbour(1 To pSize, 1 To max_degree)
ReDim neighbour_dist(1 To pSize, 1 To max_degree)
ReDim sort_index(1 To pSize)
For n = 1 To pn_edge
    i = pEdgeList(n, 1)
    j = pEdgeList(n, 2)
    tmp_x = pEdgeDist(n)
    sort_index(i) = sort_index(i) + 1
    sort_index(j) = sort_index(j) + 1
    neighbour(i, sort_index(i)) = j
    neighbour(j, sort_index(j)) = i
    neighbour_dist(i, sort_index(i)) = tmp_x
    neighbour_dist(j, sort_index(j)) = tmp_x
Next n
Erase sort_index

'Dijkstra's algorithm
ReDim d(1 To pSize)
ReDim dist(1 To pSize, 1 To pSize)
For s = 1 To pSize

    If s Mod 50 = 0 Then
        DoEvents
        Application.StatusBar = "cMST: Dijkstra..." & s & "/" & pSize
    End If

    Set q = New cHeap
    Call q.Init

    For v = 1 To pSize
        d(v) = INFINITY
    Next v

    Call q.Add(0, s)
    d(s) = 0
    Do While q.Size > 0
        Call q.Pop_Min(d_min, u)
        For i = 1 To pnode_degree(u)
            v = neighbour(u, i)
            tmp_x = d(u) + neighbour_dist(u, i)
            If tmp_x < d(v) Then
                d(v) = tmp_x
                Call q.Add(tmp_x, v)
            End If
        Next i
    Loop

    For v = 1 To pSize
        dist(s, v) = d(v)
    Next v

Next s
Call q.Reset
Set q = Nothing
Erase d, neighbour, neighbour_dist
Application.StatusBar = False
End Sub


'Sub Find_Eigen()
'Dim i As Long, j As Long, k As Long, iterate As Long
'Dim A() As Double, eigen_val() As Double, eigen_vec() As Double
'ReDim A(1 To psize, 1 To psize)
'For k = 1 To pn_edge
'    i = pEdgeList(k, 1)
'    j = pEdgeList(k, 2)
'    A(i, j) = 1
'    A(j, i) = 1
'Next k
'Call mEigen.Eigen_QR(A, eigen_val, eigen_vec)
'ReDim pnode_eigen(1 To psize)
'Debug.Print "A"
'For i = 1 To psize
'    pnode_eigen(i) = eigen_vec(1, i)
'Next i
'Debug.Print "B"
'For i = 1 To 100
'    Debug.Print "Eigen value, " & i & ", " & eigen_val(i)
'Next i
'End Sub


Sub Find_Eigen(Optional iter_max As Long = 10000, Optional tolerance As Double = 0.0000000001)
Dim i As Long, j As Long, k As Long, iterate As Long
Dim x() As Double, tmp_x As Double, eigen_val As Double
Dim tmpBool As Boolean

Randomize
Call Me.Find_degree
eigen_val = 0
ReDim pnode_eigen(1 To pSize)
For i = 1 To pSize
    pnode_eigen(i) = pnode_degree(i)
    eigen_val = eigen_val + pnode_eigen(i) ^ 2
Next i
eigen_val = Sqr(eigen_val)
For i = 1 To pSize
    pnode_eigen(i) = pnode_eigen(i) / eigen_val
Next i

tmpBool = PowerIterate(pnode_eigen, tmp_x, iter_max, tolerance)

If tmpBool = False Then
    Debug.Print "cGraphAlgo:Find_Eigen: has not converged. Possible degeneracy. Err=" & Format(tmp_x, "0.00E+00") & _
        ". Taking average of last two iterations."
    x = pnode_eigen
    tmpBool = PowerIterate(x, tmp_x, 1, tolerance)
    For i = 1 To pSize
        pnode_eigen(i) = (pnode_eigen(i) + x(i)) / 2
    Next i
    tmpBool = PowerIterate(pnode_eigen, tmp_x, iter_max, tolerance)
    If tmpBool = False Then
        Debug.Print "cGraphAlgo:Find_Eigen: Still no convergence. Err=" & Format(tmp_x, "0.00E+00")
    Else
        Debug.Print "cGraphAlgo:Find_Eigen: Reachieve convergence. Err=" & Format(tmp_x, "0.00E+00")
    End If
End If

End Sub


Private Function PowerIterate(vec_guess() As Double, ConvChk As Double, Optional iter_max As Long = 10000, Optional tolerance As Double = 0.0000000001) As Boolean
Dim i As Long, j As Long, k As Long, iterate As Long
Dim x() As Double, tmp_x As Double, tmp_y As Double, eigen_val As Double
    For iterate = 1 To iter_max
        ReDim x(1 To pSize)
        For k = 1 To pn_edge
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            x(i) = x(i) + vec_guess(j)
            x(j) = x(j) + vec_guess(i)
        Next k
        eigen_val = 0
        tmp_x = 0
        For i = 1 To pSize
            x(i) = x(i) + 1 'Find eigenvector of (A+I) instead of A
            eigen_val = eigen_val + x(i) ^ 2
            tmp_x = tmp_x + x(i) * vec_guess(i)
        Next i
        eigen_val = Sgn(tmp_x) * Sqr(eigen_val)
        For i = 1 To pSize
            x(i) = x(i) / eigen_val
        Next i
        tmp_x = tmp_x / eigen_val
        tmp_x = Abs(tmp_x - 1)
        vec_guess = x
        If tmp_x < tolerance Then Exit For
    Next iterate
    Erase x
    If iterate >= iter_max Then
        PowerIterate = False
    Else
        PowerIterate = True
    End If
    ConvChk = tmp_x
    Erase x
End Function


Private Sub mTranspose(A As Variant, B As Variant)
Dim i As Long, j As Long, m As Long, n As Long
    m = UBound(A, 1)
    n = UBound(A, 2)
    ReDim B(1 To n, 1 To m)
    For i = 1 To m
        For j = 1 To n
            B(j, i) = A(i, j)
        Next j
    Next i
End Sub

Private Sub condense2sq(k As Long, n As Long, i As Long, j As Long)
    i = Application.WorksheetFunction.Ceiling(((n - 0.5) - Sqr((n - 0.5) ^ 2 - 2 * k)), 1)
    j = k + i - (i - 1) * (2 * n - i) / 2
End Sub



'=== Build Planar Maximally Filtered Graph from a pairwise distance matrix
'Input: Symmetric Distance(N x N) matrix
Sub PMFG_Build(distance() As Double)
Dim i As Long, j As Long, m As Long, n As Long, k As Long, n_pairs As Long
Dim node_index() As Long, sort_index() As Long, AdjMatrix() As Long, included_nodes() As Long
Dim theGraph As cPMFG_Graph
Dim isPlanar As String
Dim M_is_new As Boolean, N_is_new As Boolean
Dim d() As Double

pSize = UBound(distance, 1)
n_pairs = pSize * (pSize - 1) / 2
ReDim pnode_idx(1 To pSize)
For i = 1 To pSize
    pnode_idx(i) = i
Next i

k = 0
ReDim d(1 To n_pairs)
For i = 1 To pSize - 1
    For j = i + 1 To pSize
        k = k + 1
        d(k) = distance(i, j)
    Next j
Next i

Application.StatusBar = "Building PMFG: Sorting distances...."
Call modMath.Sort_Quick_A(d, 1, n_pairs, sort_index)

'=== Construct PMFG
pn_edge = 0
ReDim included_nodes(0 To 0)
ReDim node_index(1 To pSize)
ReDim edge_stat(1 To n_pairs)
ReDim AdjMatrix(1 To 2, 1 To 1)
ReDim pEdgeList(1 To 3 * pSize - 6, 1 To 2)
ReDim pEdgeDist(1 To 3 * pSize - 6)
For i = 1 To n_pairs
    edge_stat(i) = False
Next i

For i = 1 To n_pairs

    DoEvents
    Application.StatusBar = "Contructing PMFG..." & pn_edge & " / " & (3 * pSize - 6)

    j = sort_index(i)
    Call condense2sq(j, pSize, m, n)

    'Temporarily include this edge
    edge_stat(j) = True
    pn_edge = pn_edge + 1
    M_is_new = Add_Node(included_nodes, m, node_index)
    N_is_new = Add_Node(included_nodes, n, node_index)

    'Construct temporary adjacency matrix
    ReDim Preserve AdjMatrix(1 To 2, 1 To pn_edge)
    AdjMatrix(1, pn_edge) = node_index(m) - 1
    AdjMatrix(2, pn_edge) = node_index(n) - 1

    'Check if the graph is planar, it not try the next edge
    Call gp.ReadAdjMatrix(theGraph, AdjMatrix, UBound(included_nodes))
    isPlanar = gp_Embed.Embed(theGraph)

    If isPlanar <> "OK" Then
        'restore to previous state
        edge_stat(j) = False
        pn_edge = pn_edge - 1
        ReDim Preserve AdjMatrix(1 To 2, 1 To pn_edge)
        If N_is_new = True Then Call Delete_Node(included_nodes, n, node_index)
        If M_is_new = True Then Call Delete_Node(included_nodes, m, node_index)
    ElseIf isPlanar = "OK" Then
        pEdgeList(pn_edge, 1) = m
        pEdgeList(pn_edge, 2) = n
        pEdgeDist(pn_edge) = d(i)
    End If

    If pn_edge >= (3 * pSize - 6) Then Exit For

Next i
'==============================

If pn_edge <> (3 * pSize - 6) Or UBound(included_nodes) <> pSize Then
    Debug.Print "PMFG_Build Fail: Num of edges is not equal to 3N-6"
End If

Erase sort_index, d
Erase included_nodes, node_index, edge_stat, AdjMatrix
Application.StatusBar = False

Randomize
ReDim pnode_pos(1 To pSize, 1 To 2)
For i = 1 To pSize
    pnode_pos(i, 1) = (-0.5 + Rnd()) * Sqr(pSize)
    pnode_pos(i, 2) = (-0.5 + Rnd()) * Sqr(pSize)
Next i
End Sub

Private Function Add_Node(included_nodes() As Long, m As Long, node_index() As Long) As Boolean
Dim i As Long, k As Long, n As Long
Add_Node = False
If node_index(m) = 0 Then
    n = UBound(included_nodes) + 1
    ReDim Preserve included_nodes(0 To n)
    included_nodes(n) = m
    node_index(m) = n
    Add_Node = True
End If
End Function

Private Sub Delete_Node(included_nodes() As Long, m As Long, node_index() As Long)
Dim i As Long, k As Long, n As Long
If node_index(m) > 0 Then
    n = UBound(included_nodes)
    If node_index(m) < n Then
        For i = node_index(m) + 1 To n
            node_index(included_nodes(i)) = node_index(included_nodes(i)) - 1
            included_nodes(i - 1) = included_nodes(i)
        Next i
    End If
    ReDim Preserve included_nodes(0 To n - 1)
    node_index(m) = 0
End If
End Sub

