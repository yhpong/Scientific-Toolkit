VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGraph_Undir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSize As Long, pn_edge As Long
Private pnode_pos() As Double
Private pEdgeList() As Long, pEdgeWgt() As Double
Private pnode_eigen() As Double, pnode_pagerank() As Double, pnode_katz() As Double
Private pnode_reach() As Long
Private pnode_deg() As Long, pnode_core() As Long, pnode_onion() As Long
Private pnode_deg_wgt() As Double
Private pnode_closeness() As Double
Private pnode_betweenness() As Double
Private pnode_cluster_coeff() As Double, pcluster_coeff As Double
Private pnode_cluster_coeff_wgt() As Double, pcluster_coeff_wgt As Double
Private pKey2Idx() As Long
Private pNodeList As cAdjList
Private pAdjLists As Collection
Private pCollapse_Map() As Long

'========================================================
'Access Properties
'========================================================

Public Property Get Collapse_Map() As Long()
    Collapse_Map = pCollapse_Map
End Property

Public Property Let Collapse_Map(Collapse_Map() As Long)
    pCollapse_Map = Collapse_Map
End Property

Public Property Get Size() As Long
    Size = pSize
End Property

Public Property Get n_edge() As Long
    n_edge = pn_edge
End Property

Public Property Get node_deg() As Long()
    node_deg = pnode_deg
End Property

Public Property Get node_deg_wgt() As Double()
    node_deg_wgt = pnode_deg_wgt
End Property

Public Property Get node_reach() As Long()
    node_reach = pnode_reach
End Property

Public Property Get node_closeness() As Double()
    node_closeness = pnode_closeness
End Property

Public Property Get node_betweenness() As Double()
    node_betweenness = pnode_betweenness
End Property


Public Property Get node_eigen() As Double()
    node_eigen = pnode_eigen
End Property

Public Property Get node_katz() As Double()
    node_katz = pnode_katz
End Property

Public Property Get node_pagerank() As Double()
    node_pagerank = pnode_pagerank
End Property

Public Property Get node_core() As Long()
    node_core = pnode_core
End Property

Public Property Get node_onion() As Long()
    node_onion = pnode_onion
End Property

Public Property Let node_pos(x() As Double)
    pnode_pos = x
End Property

Public Property Get node_pos() As Double()
    node_pos = pnode_pos
End Property


Public Property Get node_cluster_coeff() As Double()
    node_cluster_coeff = pnode_cluster_coeff
End Property

Public Property Get cluster_coeff() As Double
    cluster_coeff = pcluster_coeff
End Property


Public Property Get node_cluster_coeff_wgt() As Double()
    node_cluster_coeff_wgt = pnode_cluster_coeff_wgt
End Property

Public Property Get cluster_coeff_wgt() As Double
    cluster_coeff_wgt = pcluster_coeff_wgt
End Property

Public Property Get EdgeList(Optional outputkey As Boolean = False) As Long()
Dim i As Long, j As Long, k As Long
Dim iArr() As Long, jArr() As Long
    If outputkey = False Then
        EdgeList = pEdgeList
    ElseIf outputkey = True Then
        ReDim iArr(1 To pn_edge, 1 To 2)
        jArr = pNodeList.GetList
        For k = 1 To pn_edge
            iArr(k, 1) = jArr(pEdgeList(k, 1))
            iArr(k, 2) = jArr(pEdgeList(k, 2))
        Next k
        EdgeList = iArr
        Erase iArr, jArr
    End If
End Property

Public Property Get EdgeWgts() As Double()
    EdgeWgts = pEdgeWgt
End Property

Public Function NodeList() As Long()
    NodeList = pNodeList.GetList
End Function

Public Function NodeList_Inv() As Long()
    Call Calc_Key2Idx
    NodeList_Inv = pKey2Idx
End Function

Public Function NodeWgts() As Double()
    NodeWgts = pNodeList.GetListWgt
End Function

Public Property Get AdjMatrix() As Double()
Dim i As Long, j As Long, k As Long
Dim A() As Double
Dim iArr() As Long, xArr() As Double
    Call Calc_Key2Idx
    ReDim A(1 To pSize, 1 To pSize)
    For i = 1 To pSize
        Call pAdjLists(i).GetListnWgt(iArr, xArr)
        For j = 1 To UBound(iArr)
            A(i, pKey2Idx(iArr(j))) = xArr(j)
        Next j
    Next i
    AdjMatrix = A
    Erase A, iArr, xArr
End Property

Sub SubGraph(g1 As cGraph_Undir, node_list() As Long, Optional isKey As Boolean = True, _
            Optional KeepNeighbors As Boolean = False, _
            Optional adjpos As Boolean = False)
Dim i As Long, j As Long, k As Long
Dim isIncl() As Long, node_keys() As Long, neigh_list() As Long
    ReDim isIncl(1 To pSize)
    Call Calc_Key2Idx
    node_keys = pNodeList.GetList
    If isKey = False Then
        For i = 1 To UBound(node_list)
            isIncl(node_list(i)) = 1
            If KeepNeighbors = True Then
                neigh_list = pAdjLists(node_list(i)).GetList
                For j = 1 To UBound(neigh_list)
                    isIncl(pKey2Idx(neigh_list(j))) = 1
                Next j
            End If
        Next i
    Else
        For i = 1 To UBound(node_list)
            isIncl(pKey2Idx(node_list(i))) = 1
            If KeepNeighbors = True Then
                neigh_list = pAdjLists(pKey2Idx(node_list(i))).GetList
                For j = 1 To UBound(neigh_list)
                    isIncl(pKey2Idx(neigh_list(j))) = 1
                Next j
            End If
        Next i
    End If
    Set g1 = New cGraph_Undir
    With g1
        Call .Copy_from(Me)
        If adjpos = True Then .node_pos = pnode_pos
        For i = 1 To pSize
            If isIncl(i) = 0 Then
                Call .NodeDelete(node_keys(i), adjpos)
            End If
        Next i
    End With
End Sub

Function node_neighbors(v As Long) As Long()
    node_neighbors = pAdjLists(Key2Idx(v)).GetList
End Function

Function node_neighbors_wgt(v As Long) As Double()
    node_neighbors_wgt = pAdjLists(Key2Idx(v)).GetListWgt
End Function

'========================================================
'Initialization
'========================================================

Sub Init(n_node As Long, Optional node_wgt As Variant, Optional isWgt As Boolean = False, _
            Optional node_key As Variant, Optional custom_key As Boolean = False)
Dim i As Long, j As Long
Dim v As cAdjList
    pSize = n_node
    Set pAdjLists = New Collection
    For i = 1 To pSize
        Set v = New cAdjList
        pAdjLists.Add v
    Next i
    Set pNodeList = New cAdjList
    If custom_key = False Then
        With pNodeList
            If isWgt = True Then
                For i = 1 To pSize
                    j = .Append(i, VBA.CDbl(node_wgt(i)))
                Next i
            Else
                For i = 1 To pSize
                    j = .Append(i, 1)
                Next i
            End If
        End With
    Else
        With pNodeList
            If isWgt = True Then
                For i = 1 To pSize
                    j = .Append(Int(node_key(i)), VBA.CDbl(node_wgt(i)))
                Next i
            Else
                For i = 1 To pSize
                    j = .Append(Int(node_key(i)), 1)
                Next i
            End If
        End With
    End If
End Sub

Sub Reset()
Dim i As Long, n As Long
    Erase pKey2Idx
    Erase pnode_deg, pnode_deg_wgt, pnode_reach, pnode_core, pnode_onion
    Erase pnode_eigen, pnode_pagerank, pnode_closeness, pnode_katz, pnode_betweenness
    Erase pEdgeList, pEdgeWgt, pnode_pos, pCollapse_Map
    n = pAdjLists.count
    For i = n To 1 Step -1
        pAdjLists.Remove (i)
    Next i
    Set pAdjLists = Nothing
    Call pNodeList.Init
    Set pNodeList = Nothing
    pSize = 0
    pn_edge = 0
End Sub

Sub Convert_from_Directed(g1 As cGraph_Dir)
Dim i As Long, j As Long, n As Long
Dim idx() As Long, iArr() As Long, xArr() As Double
    With g1
        n = .Size
        xArr = .NodeWgts
        idx = .NodeList
    End With
    
    Call Init(n, xArr, True, idx, True)
    pn_edge = 0
    For i = 1 To n
        If i Mod 300 = 0 Then
            DoEvents: Application.StatusBar = "Copying Edges " & i & "/" & n
        End If
        iArr = g1.node_neighbors(idx(i), "TOTAL", xArr, True)
        j = pAdjLists(i).ReplaceList(iArr, xArr, True)
        For j = 1 To UBound(iArr)
            If iArr(j) >= idx(i) Then pn_edge = pn_edge + 1
        Next j
    Next i
    Erase idx, iArr, xArr
    Application.StatusBar = False
End Sub

'Make current graph into an identical copy of g1
Sub Copy_from(g1 As cGraph_Undir)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim idx() As Long
Dim xArr() As Double, iArr() As Long
    If pSize > 0 Then Call Reset
    With g1
        n = .Size
        xArr = .NodeWgts
        idx = .NodeList
    End With
    Call Init(n, xArr, True, idx, True)
    pn_edge = g1.n_edge
    For i = 1 To n
        If i Mod 300 = 0 Then
            DoEvents: Application.StatusBar = "Copying Edges " & i & "/" & n
        End If
        iArr = g1.node_neighbors(idx(i))
        xArr = g1.node_neighbors_wgt(idx(i))
        j = pAdjLists(i).ReplaceList(iArr, xArr, True)
    Next i
    Erase idx, iArr, xArr
    Application.StatusBar = False
End Sub


Sub AdjM_Add(AdjMatrix As Variant, Optional isWgt As Boolean = True, _
                Optional ignore_selfloop As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long
Dim xArr() As Double, iArr() As Long, node_keys() As Long, node_keys_inv() As Long
    If pSize <> UBound(AdjMatrix, 1) Or pSize <> UBound(AdjMatrix, 2) Then
        Debug.Print "cGraph_undir: AdjM_Add: size of matrix inconsistent with number of nodes."
        Exit Sub
    End If
    node_keys = NodeList
    node_keys_inv = NodeList_Inv
    'Clear all existing edges
    pn_edge = 0
    For i = 1 To pSize
        Call pAdjLists(i).Init
    Next i
    For i = 1 To pSize
        u = node_keys(i)
        If i Mod 100 = 0 Then
            DoEvents: Application.StatusBar = "Importing Adjacency matrix: " & i & "/" & pSize
        End If
        k = 0
        ReDim xArr(1 To pSize)
        ReDim iArr(1 To pSize)
        For j = 1 To pSize
            If j <> i And AdjMatrix(i, j) <> 0 Then
                k = k + 1
                iArr(k) = node_keys(j)
                xArr(k) = AdjMatrix(i, j)
            End If
        Next j
        If k > 0 Then
            ReDim Preserve iArr(1 To k)
            If isWgt = True Then
                ReDim Preserve xArr(1 To k)
                m = pAdjLists(i).ReplaceList(iArr, xArr, True)
            Else
                m = pAdjLists(i).ReplaceList(iArr)
            End If
            For j = 1 To k
                If node_keys_inv(iArr(j)) > i Then pn_edge = pn_edge + 1
            Next j
        End If
    Next i
    
    'Add self-loop
    If ignore_selfloop = False Then
        For i = 1 To pSize
            If AdjMatrix(i, i) <> 0 Then
                u = node_keys(i)
                Call EdgeAdd(u, u, VBA.CDbl(AdjMatrix(i, i)))
            End If
        Next i
    End If
    
    Erase xArr, iArr, node_keys, node_keys_inv
    Application.StatusBar = False
End Sub


'Set graph to some preset shapes, available shapes and the corresponding input n_size include:
'"LINE", number of vertex
'"POLYGON", number of vertex
'"GRID_2D"/"GRID_3D", number of vertex on each side
'"CYLINDER", number of vertex on the top & bottom polygon
'"C60", Carbon 60 football, n_size not relevant
'"STAR"/"STAR2", number leaves
'"WHEEL", number vertex -1
'"HYPERCUBE", number of dimensions
'"CAYLEY", degree of non-leaf node, create up to 5 layers
Sub Init_Shapes(Optional strType As String = "POLYGON", Optional n_size As Long = 6)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_node As Long, n_edge As Long, ii As Long, jj As Long
Dim EdgeList() As Long
    If pSize > 0 Then Call Reset
    If strType = "LINE" Then
        n_node = n_size
        n_edge = n_node - 1: ReDim EdgeList(1 To n_edge, 1 To 2)
        For i = 1 To n_node - 1
            EdgeList(i, 1) = i: EdgeList(i, 2) = i + 1
        Next
    ElseIf strType = "POLYGON" Then
        n_node = n_size
        n_edge = n_node: ReDim EdgeList(1 To n_edge, 1 To 2)
        For i = 1 To n_node
            If i < n_node Then
                EdgeList(i, 1) = i: EdgeList(i, 2) = i + 1
            Else
                EdgeList(i, 1) = i: EdgeList(i, 2) = 1
            End If
        Next i
    ElseIf strType = "GRID_3D" Then
        m = n_size ^ 2
        n_node = n_size ^ 3
        n = 0: n_edge = 0: ReDim EdgeList(1 To 3 * m * (n_size - 1), 1 To 2)
        For k = 1 To n_size
            For j = 1 To n_size
                For i = 1 To n_size
                    n = n + 1
                    If i < n_size Then
                        n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n + 1
                    End If
                    If j < n_size Then
                        n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n + n_size
                    End If
                    If k < n_size Then
                        n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n + m
                    End If
                Next i
            Next j
        Next k
    ElseIf strType = "GRID_2D" Then
        n_node = n_size ^ 2
        n = 0: n_edge = 0: ReDim EdgeList(1 To 2 * (n_node - n_size), 1 To 2)
        For j = 1 To n_size
            For i = 1 To n_size
                n = n + 1
                If i < n_size Then
                    n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n + 1
                End If
                If j < n_size Then
                    n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n + n_size
                End If
            Next i
        Next j
    ElseIf strType = "CYLINDER" Then
        n_node = n_size ^ 2
        n = 0: n_edge = 0: ReDim EdgeList(1 To 2 * (n_node - n_size) + n_size, 1 To 2)
        For j = 1 To n_size
            For i = 1 To n_size
                n = n + 1
                If i < n_size Then
                    n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n + 1
                End If
                If j < n_size Then
                    n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n + n_size
                End If
                If j = n_size Then
                    n_edge = n_edge + 1: EdgeList(n_edge, 1) = n: EdgeList(n_edge, 2) = n - n_size * (n_size - 1)
                End If
            Next i
        Next j
    ElseIf strType = "C60" Then
        n_node = 60
        n_edge = 0: ReDim EdgeList(1 To 90, 1 To 2)
        For i = 1 To 60 Step 5
            For j = 1 To 4
                n_edge = n_edge + 1: EdgeList(n_edge, 1) = i + j - 1: EdgeList(n_edge, 2) = i + j
            Next j
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i: EdgeList(n_edge, 2) = i + 4
        Next i
        For i = 1 To 5
            m = 5 * i - (i - 1)
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i: EdgeList(n_edge, 2) = i + m
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i + 30: EdgeList(n_edge, 2) = i + 30 + m
        Next i
        n_edge = n_edge + 1: EdgeList(n_edge, 1) = 7: EdgeList(n_edge, 2) = 30
        n_edge = n_edge + 1: EdgeList(n_edge, 1) = 37: EdgeList(n_edge, 2) = 60
        For m = 2 To 5
            i = 5 * m
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i: EdgeList(n_edge, 2) = i + 2
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i + 30: EdgeList(n_edge, 2) = i + 30 + 2
        Next m
        For m = 1 To 3
            i = 8 + (m - 1) * 5
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i: EdgeList(n_edge, 2) = i + 41
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i + 11: EdgeList(n_edge, 2) = i + 11 + 19
        Next m
        For m = 1 To 2
            i = 9 + (m - 1) * 5
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i: EdgeList(n_edge, 2) = i + 44
            n_edge = n_edge + 1: EdgeList(n_edge, 1) = i + 14: EdgeList(n_edge, 2) = i + 14 + 16
        Next m
    ElseIf strType = "STAR" Then
        n_node = n_size + 1
        ReDim EdgeList(1 To n_size, 1 To 2)
        For i = 1 To n_size
            EdgeList(i, 1) = 1: EdgeList(i, 2) = 1 + i
        Next i
    ElseIf strType = "STAR2" Then
        n_node = n_size + 1 + n_size ^ 2
        ReDim EdgeList(1 To n_size + n_size ^ 2, 1 To 2)
        For i = 1 To n_size
            EdgeList(i, 1) = 1: EdgeList(i, 2) = 1 + i
        Next i
        n = n_size
        For i = 1 To n_size
            For j = 1 To n_size
                n = n + 1: EdgeList(n, 1) = i + 1: EdgeList(n, 2) = n + 1
            Next j
        Next i
    ElseIf strType = "WHEEL" Then
        n_node = n_size + 1
        ReDim EdgeList(1 To 2 * n_size, 1 To 2)
        For i = 1 To n_size
            EdgeList(i, 1) = 1: EdgeList(i, 2) = 1 + i
            If i < n_size Then
                EdgeList(n_size + i, 1) = 1 + i: EdgeList(n_size + i, 2) = 1 + i + 1
            Else
                EdgeList(n_size + i, 1) = 1 + i: EdgeList(n_size + i, 2) = 2
            End If
        Next i
    ElseIf strType = "PRISM" Then
        n_node = 2 * n_size
        ReDim EdgeList(1 To 3 * n_size, 1 To 2)
        For i = 1 To n_size
            If i < n_size Then
                EdgeList(i, 1) = i: EdgeList(i, 2) = i + 1
                EdgeList(n_size + i, 1) = n_size + i: EdgeList(n_size + i, 2) = n_size + i + 1
            Else
                EdgeList(i, 1) = i: EdgeList(i, 2) = 1
                EdgeList(n_size + i, 1) = n_size + i: EdgeList(n_size + i, 2) = n_size + 1
            End If
            EdgeList(2 * n_size + i, 1) = i: EdgeList(2 * n_size + i, 2) = n_size + i
        Next i
    ElseIf strType = "HYPERCUBE" Then
        n_node = 2 ^ n_size
        n_edge = 0: ReDim EdgeList(1 To (n_node / 2) * n_size, 1 To 2)
        n = 0: m = 0
        For k = 1 To n_size
            n = 2 ^ (k - 1)
            If k > 1 Then
                m = (2 ^ (k - 2)) * (k - 1)
                For i = 1 To m
                    n_edge = n_edge + 1
                    EdgeList(n_edge, 1) = EdgeList(n_edge - m, 1) + n
                    EdgeList(n_edge, 2) = EdgeList(n_edge - m, 2) + n
                Next i
            End If
            For i = 1 To n
                n_edge = n_edge + 1: EdgeList(n_edge, 1) = i: EdgeList(n_edge, 2) = i + n
            Next i
        Next k
    ElseIf strType = "CAYLEY" Then
        jj = 5
        n_node = (((n_size - 1) ^ jj - 1) / (n_size - 2)) * n_size + 1: n_edge = n_node - 1
        ReDim EdgeList(1 To n_edge, 1 To 2)
        For i = 1 To n_size
            EdgeList(i, 1) = 1
            EdgeList(i, 2) = i + 1
        Next i
        n_edge = n_size: ii = 1
        For k = 2 To jj
            m = n_size * ((n_size - 1) ^ (k - 2))
            n = n_size * ((n_size - 1) ^ (k - 1))
            For i = 1 To m
                For j = 1 To n_size - 1
                    n_edge = n_edge + 1
                    EdgeList(n_edge, 1) = ii + i
                    EdgeList(n_edge, 2) = n_edge + 1
                Next j
            Next i
            ii = ii + m
        Next k
    End If
    Call Init(n_node)
    Call EdgeAddList(EdgeList)
    Erase EdgeList
End Sub


'Initialize a random graph with of size n_node, and probability x_prob for an edge to exists
Sub Init_Random_Graph(n_node As Long, Optional x_prob As Double = 0.02, Optional allow_selfloop As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double
Dim x_idx() As Long, y_idx() As Long
    VBA.Randomize
    Call Init(n_node)
    
    If x_prob > 1 Then
        x_prob = 1
    ElseIf x_prob < 0 Then
        x_prob = 0
    End If
    
    If allow_selfloop = False Then
        For i = 1 To n_node - 1
            For j = i + 1 To n_node
                tmp_x = VBA.Rnd
                If tmp_x < x_prob Then Call EdgeAdd(i, j)
            Next j
        Next i
    Else
        For i = 1 To n_node
            For j = i To n_node
                tmp_x = VBA.Rnd
                If tmp_x < x_prob Then Call EdgeAdd(i, j)
            Next j
        Next i
    End If
End Sub


'===============================================================
'Node Operations
'===============================================================

'Returns position index if node key v exist in graph, 0 otherwise
Function NodeExists(v As Long) As Long
Dim tmp_x As Double
    NodeExists = pNodeList.isConnect(v, tmp_x)
End Function

'Add a node to graph using v as key, if node already exists no action is performed
Sub NodeAdd(v As Long, Optional Wgt As Double = 1)
Dim i As Long, j As Long, k As Long, iArr() As Long, xArr() As Double
Dim vList As cAdjList
    i = pNodeList.Append(v, Wgt)
    pSize = pSize + i
    If i > 0 Then 'Node is indeed new
        Set vList = New cAdjList
        pAdjLists.Add vList
    End If
End Sub

'Delete a node from graph by its key
Sub NodeDelete(v As Long, Optional adjpos As Boolean = False)
Dim i As Long, j As Long, k As Long, idx As Long
Dim iArr() As Long, xArr() As Double
    idx = Key2Idx(v)
    iArr = pAdjLists(idx).GetList
    If UBound(iArr) > 0 Then
        For j = 1 To UBound(iArr)
            Call EdgeDelete(v, iArr(j))
        Next j
    End If
    Erase iArr
    i = pNodeList.Remove(v)
    pSize = pSize - i
    If i > 0 Then 'Node is indeed removed
        pAdjLists.Remove idx
    End If
    If i > 0 And adjpos = True Then
        xArr = pnode_pos
        ReDim pnode_pos(1 To pSize, 1 To UBound(xArr, 2))
        For k = 1 To UBound(pnode_pos, 2)
            For j = 1 To idx - 1
                pnode_pos(j, k) = xArr(j, k)
            Next j
            For j = idx To pSize
                pnode_pos(j, k) = xArr(j + 1, k)
            Next j
        Next k
        Erase xArr
    End If
End Sub

'Get the weight of node v
Function NodeWgt(v As Long) As Double
Dim k As Long, tmp_x As Double
    k = pNodeList.isConnect(v, tmp_x)
    NodeWgt = tmp_x
End Function

'Update the weight of node v
'0=no action perform
'1=Replace existing value
'2=Accumlate existing value
Sub NodeWgt_Update(v As Long, Wgt As Double, Optional replace_wgt As Long = 1)
Dim k As Long
    k = pNodeList.Append(v, Wgt, replace_wgt)
    If k > 0 Then Debug.Print "NodeWgt_Update: node " & v & " does not exist. New node created."
End Sub

'For a node key v, return its node index in graph
Function Key2Idx(v As Long) As Long
Dim tmp_x As Double
    Key2Idx = pNodeList.isConnect(v, tmp_x)
End Function

'Save node index of all node keys in an array
Sub Calc_Key2Idx()
Dim i As Long, v As Long
Dim iArr() As Long
    iArr = pNodeList.GetList
    v = -1
    For i = 1 To pSize
        If iArr(i) > v Then v = iArr(i)
    Next i
    ReDim pKey2Idx(1 To v)
    For i = 1 To pSize
        pKey2Idx(iArr(i)) = i
    Next i
    Erase iArr
End Sub

'Merge v into u, reference to v is destroyed
Sub NodeMerge(u As Long, v As Long, Optional adjpos As Boolean = False)
Dim i As Long, j As Long, k As Long, w As Long
Dim iArr() As Long, xArr() As Double, tmp_x As Double, tmp_y As Double
    If u = v Then Exit Sub
    i = Key2Idx(u)
    j = Key2Idx(v)
    If adjpos = True Then
        tmp_x = NodeWgt(u)
        tmp_y = NodeWgt(v)
        For k = 1 To UBound(pnode_pos, 2)
            pnode_pos(i, k) = (pnode_pos(i, k) * tmp_x + pnode_pos(j, k) * tmp_y) / (tmp_x + tmp_y)
        Next k
    End If
    Call EdgeDelete(u, v)
    'edges from v are added to those of u
    Call pAdjLists(j).GetListnWgt(iArr, xArr)
    For k = 1 To UBound(iArr)
        w = iArr(k)
        Call EdgeDelete(v, w)
        Call EdgeAdd(u, w, xArr(k), 2)
    Next k
    Call pNodeList.Append(u, NodeWgt(v), 2) 'weight of v is added to u
    Call NodeDelete(v, adjpos) 'Destroy v
End Sub


'Reorder nodes in current graph
'Input: iorder(), integer vector, ioder(i)=v means that node key v is put in the i-th position
'Output: swap_idx(), optional, swap_idx(i)=j means that the current i-th node was previously in the j-th position
Sub NodeReorder(iorder() As Long, _
        Optional swap_idx As Variant, Optional output_swap_idx As Boolean = False, _
        Optional adjpos As Boolean = False)
Dim i As Long, j As Long, k As Long
Dim swap_idx_tmp() As Long
Dim new_AdjLists As Collection, xArr() As Double
    'Reorder node list
    Call pNodeList.Reorder(iorder, swap_idx_tmp, True)
    'Reorder adjacency list
    Set new_AdjLists = New Collection
    For i = 1 To pSize
        new_AdjLists.Add pAdjLists(swap_idx_tmp(i))
    Next i
    For i = pSize To 1 Step -1
        pAdjLists.Remove (i)
    Next i
    Set pAdjLists = new_AdjLists
    'Output swap_idx()
    If output_swap_idx = True Then swap_idx = swap_idx_tmp
    'Adjust node position
    If adjpos = True Then
        xArr = pnode_pos
        For j = 1 To UBound(pnode_pos, 2)
            For i = 1 To pSize
                pnode_pos(i, j) = xArr(swap_idx_tmp(i), j)
            Next i
        Next j
        Erase xArr
    End If
End Sub


'===============================================================
'Edge Operations
'===============================================================

Function EdgeWgt(u As Long, v As Long) As Double
Dim i As Long, tmp_x As Double
    i = pAdjLists(Key2Idx(u)).isConnect(v, tmp_x)
    EdgeWgt = tmp_x
End Function

Function EdgeExists(u As Long, v As Long) As Boolean
Dim i As Long, tmp_x As Double
    EdgeExists = False
    If pAdjLists(Key2Idx(u)).isConnect(v, tmp_x) > 0 Then EdgeExists = True
End Function

'Add an edge between u & v
Sub EdgeAdd(u As Long, v As Long, Optional Wgt As Double = 1, Optional replace_wgt As Long = 0, _
        Optional check_duplicate As Boolean = True)
Dim i As Long, j As Long
    If check_duplicate = True Then
        i = pAdjLists(Key2Idx(u)).Append(v, Wgt, replace_wgt)
        j = pAdjLists(Key2Idx(v)).Append(u, Wgt, replace_wgt)
        pn_edge = pn_edge + i
    Else
        i = pAdjLists(Key2Idx(u)).Prepend(v, Wgt, replace_wgt, False)
        j = pAdjLists(Key2Idx(v)).Prepend(u, Wgt, replace_wgt, False)
        pn_edge = pn_edge + i
    End If
End Sub

'remove edge between u & v
Sub EdgeDelete(u As Long, v As Long)
Dim i As Long, j As Long
    i = pAdjLists(Key2Idx(u)).Remove(v)
    j = pAdjLists(Key2Idx(v)).Remove(u)
    pn_edge = pn_edge - i
End Sub

'Import edges from EdgeList(1:n_edge, 1:2) and EdgeWgt(1:n_edge)
Sub EdgeAddList(EdgeList() As Long, Optional EdgeWgt As Variant, Optional isWgt As Boolean = False, _
    Optional replace_wgt As Long = 0, Optional check_duplicate As Boolean = True)
Dim i As Long, n_edge As Long
    n_edge = UBound(EdgeList, 1)
    If isWgt = False Then
        For i = 1 To n_edge
            If i Mod 1000 = 0 Then
                DoEvents: Application.StatusBar = "EdgeAddList: " & i & "/" & n_edge
            End If
            Call EdgeAdd(EdgeList(i, 1), EdgeList(i, 2), 1, replace_wgt, check_duplicate)
        Next i
    Else
        For i = 1 To n_edge
            If i Mod 1000 = 0 Then
                DoEvents: Application.StatusBar = "EdgeAddList: " & i & "/" & n_edge
            End If
            Call EdgeAdd(EdgeList(i, 1), EdgeList(i, 2), VBA.CDbl(EdgeWgt(i)), replace_wgt, check_duplicate)
        Next i
    End If
    Application.StatusBar = False
End Sub

'remove all edges from u
Sub EdgeDelete_All(u As Long)
Dim uList As cAdjList
Dim iArr() As Long, i As Long
    Set uList = pAdjLists(Key2Idx(u))
    iArr = uList.GetList
    For i = 1 To UBound(iArr)
        Call EdgeDelete(u, iArr(i))
    Next i
    Call uList.Init
End Sub

Sub Calc_EdgeList()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim iArr() As Long, xArr() As Double, tmp_x As Double
Dim n_loop As Long, self_loop_list() As Long, self_loop_wgt() As Double
    n = 0
    Call Calc_Key2Idx
    ReDim pEdgeList(1 To pn_edge, 1 To 2)
    ReDim pEdgeWgt(1 To pn_edge)
    n_loop = 0
    ReDim self_loop_list(0 To pSize)
    ReDim self_loop_wgt(0 To pSize)
    For i = 1 To pSize
        Call pAdjLists(i).GetListnWgt(iArr, xArr)
        For k = 1 To UBound(iArr)
            j = pKey2Idx(iArr(k))
            If j > i Then
                n = n + 1
                pEdgeList(n, 1) = i
                pEdgeList(n, 2) = j
                pEdgeWgt(n) = xArr(k)
            ElseIf j = i Then
                n_loop = n_loop + 1
                self_loop_list(n_loop) = i
                self_loop_wgt(n_loop) = xArr(k)
            End If
        Next k
    Next i
    For k = 1 To n_loop
        n = n + 1
        pEdgeList(n, 1) = self_loop_list(k)
        pEdgeList(n, 2) = self_loop_list(k)
        pEdgeWgt(n) = self_loop_wgt(k)
    Next k
    Erase self_loop_list, self_loop_wgt
    If n <> pn_edge Then Debug.Print "Calc_EdgeList: Something wrong. Self-Loop?"
    Erase iArr, xArr
End Sub


'===============================================================
'Visualization
'===============================================================

'Randomly place nodes in an n-dimensional hypercube of volume pSize
Sub node_pos_randomize(Optional n_dimension As Long = 2)
Dim i As Long, j As Long, tmp_x As Double
    Randomize
    tmp_x = pSize ^ (1# / n_dimension)
    ReDim pnode_pos(1 To pSize, 1 To n_dimension)
    For j = 1 To n_dimension
        For i = 1 To pSize
            pnode_pos(i, j) = Rnd() * tmp_x
        Next i
    Next j
End Sub

'Calculate nodes layout, create edge list and use algorithm from undirected graph to calculate layout
'1. layout of each disconnected component is calculated separately, which are then placed side by side
'2. isWgt controls whether edges are weighted in sepectral layout
'3. WgtType controls node weight to use in force directed layout, which can be "NONE","NODE","DEG","DEG_WGT"
Sub Calc_Layout(Optional strMethod As String = "FORCE_MULTI", _
            Optional n_dimension As Long = 2, Optional isWgt As Boolean = True, _
            Optional tol As Double = 0.00000001, Optional iter_max As Long = 5000, _
            Optional isDegNorm As Boolean = True, Optional WgtType As String = "NONE", _
            Optional c_spring As Double = 1, Optional c_repulsive As Double = 1, Optional c_power As Double = 2)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_component As Long
Dim comp_idx() As Long, node_pos() As Double
Dim x_offset As Double, x_min As Double
Dim g_components As Collection
Dim g1 As cGraph_Undir
    n_component = component_count(comp_idx)
    If n_component = 1 Then
        Call Calc_Layout_Single(strMethod, n_dimension, isWgt, tol, iter_max, isDegNorm, _
                    WgtType, c_spring, c_repulsive, c_power)
        Exit Sub
    End If
    Call Disconnected_Components(g_components, comp_idx, n_component)
    ReDim pnode_pos(1 To pSize, 1 To n_dimension)
    k = 0
    Do While g_components.count > 0
        k = k + 1
        Set g1 = g_components(1)
        With g1
            n = g1.Size
            Call .Calc_Layout_Single(strMethod, n_dimension, isWgt, tol, iter_max, isDegNorm, _
                    WgtType, c_spring, c_repulsive, c_power)
            node_pos = .node_pos
            Call .Reset
        End With
        Set g1 = Nothing
        g_components.Remove 1

        If k > 1 Then
            x_min = 2.5E+30
            For i = 1 To n
                If node_pos(i, 1) < x_min Then x_min = node_pos(i, 1)
            Next i
            x_min = (x_offset - x_min) + Rnd()
            For i = 1 To n
                node_pos(i, 1) = node_pos(i, 1) + x_min
            Next i
        End If
        If k < n_component Then
            x_offset = -999999
            For i = 1 To n
                If node_pos(i, 1) > x_offset Then x_offset = node_pos(i, 1)
            Next i
        End If

        m = 0
        For i = 1 To pSize
            If comp_idx(i) = k Then
                m = m + 1
                For j = 1 To n_dimension
                    pnode_pos(i, j) = node_pos(m, j)
                Next j
            End If
        Next i
    Loop
    Set g_components = Nothing
End Sub


Sub Calc_Layout_Single(Optional strMethod As String = "FORCE_MULTI", _
            Optional n_dimension As Long = 2, Optional isWgt As Boolean = True, _
            Optional tol As Double = 0.00000001, Optional iter_max As Long = 5000, _
            Optional isDegNorm As Boolean = True, Optional WgtType As String = "NONE", _
            Optional c_spring As Double = 1, Optional c_repulsive As Double = 1, Optional c_power As Double = 2)
Dim i As Long, j As Long
'Dim g1 As cGraphAlgo
Dim tmp_x As Double
    If pSize <= 2 Then
        ReDim pnode_pos(1 To pSize, 1 To n_dimension)
        If pSize = 2 Then
            For j = 1 To n_dimension
                pnode_pos(2, j) = 1
            Next j
        End If
        Exit Sub
    End If
    If VBA.UCase$(strMethod) = "SPECTRAL" Then
        Call Layout_Spectral(n_dimension, isWgt, tol, iter_max)
    ElseIf VBA.UCase$(strMethod) = "SPECTRAL_HDE" Then
        Call Layout_Spectral_HDE(n_dimension, , isDegNorm, isWgt)
    ElseIf VBA.UCase$(strMethod) = "HDE" Then
        Call Layout_HDE(pnode_pos, n_dimension + 2)
        Call Layout_Orthonormalize(pnode_pos, tol, True, isWgt)
        ReDim Preserve pnode_pos(1 To pSize, 1 To n_dimension)
        tmp_x = pSize ^ (1# / n_dimension)
        For j = 1 To n_dimension
            For i = 1 To pSize
                pnode_pos(i, j) = pnode_pos(i, j) * tmp_x
            Next i
        Next j
    ElseIf VBA.UCase$(strMethod) = "FORCE" Then
        Call Layout_Force(n_dimension, c_spring, c_repulsive, c_power, 500, , WgtType)
    ElseIf VBA.UCase$(strMethod) = "FORCE_MULTI" Then
        Call Layout_Force_Multi(n_dimension, c_spring, c_repulsive, c_power, 500, WgtType)
    End If
End Sub


'=== Degree-normalized spectral layout
'"Drawing Graphs by Eigenvectors: Theory and Practice", Koren (2005)
Private Sub Layout_Spectral(Optional n_dimension As Long = 2, Optional isWgt As Boolean = True, _
        Optional tol As Double = 0.00000001, Optional iter_max As Long = 5000)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, pp As Long, iterate As Long
Dim x() As Double, x_prev() As Double, tmp_x As Double, tmp_y As Double, tmp_vec() As Double
Dim x0() As Double, node_deg As Variant, x_HDE() As Double, xnorm() As Double

    Call Calc_EdgeList
    If isWgt = True Then
        If isAllocated(pnode_deg_wgt) = False Then Call Calc_Node_Deg_Wgt
        node_deg = pnode_deg_wgt
    Else
        If isAllocated(pnode_deg) = False Then Call Calc_Node_Deg
        node_deg = pnode_deg
    End If

    tmp_x = Sqr(pSize)
    ReDim x0(1 To pSize)
    For i = 1 To pSize
        x0(i) = 1# / tmp_x
    Next i
    
    Randomize
    ReDim pnode_pos(1 To pSize, 1 To n_dimension)
    ReDim xnorm(0 To n_dimension)
    For i = 1 To pSize
        xnorm(0) = xnorm(0) + (x0(i) ^ 2) * node_deg(i)
    Next i

    For p = 1 To n_dimension

        ReDim x_prev(1 To pSize)
        For i = 1 To pSize
            x_prev(i) = Rnd()
        Next i
        Call zNormalize_Vec(x_prev, tmp_x)

        For iterate = 1 To iter_max
            If iterate Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = "Layout_Spectral: " & iterate & "/" & iter_max & ", " & p & "/" & n_dimension
            End If
            x = x_prev

            tmp_x = 0: tmp_y = xnorm(0)
            For i = 1 To pSize
                tmp_x = tmp_x + node_deg(i) * x0(i) * x(i)
            Next i
            For i = 1 To pSize
                x(i) = x(i) - (tmp_x / tmp_y) * x0(i)
            Next i

            If p > 1 Then
                For pp = 1 To p - 1
                    tmp_x = 0: tmp_y = xnorm(pp)
                    For i = 1 To pSize
                        tmp_x = tmp_x + node_deg(i) * pnode_pos(i, pp) * x(i)
                    Next i
                    For i = 1 To pSize
                        x(i) = x(i) - (tmp_x / tmp_y) * pnode_pos(i, pp)
                    Next i
                Next pp
            End If

            ReDim tmp_vec(1 To pSize)
            If isWgt = True Then
                For k = 1 To pn_edge
                    i = pEdgeList(k, 1)
                    j = pEdgeList(k, 2)
                    tmp_x = pEdgeWgt(k)
                    tmp_vec(i) = tmp_vec(i) + tmp_x * x(j)
                    tmp_vec(j) = tmp_vec(j) + tmp_x * x(i)
                Next k
            Else
                For k = 1 To pn_edge
                    i = pEdgeList(k, 1)
                    j = pEdgeList(k, 2)
                    tmp_vec(i) = tmp_vec(i) + x(j)
                    tmp_vec(j) = tmp_vec(j) + x(i)
                Next k
            End If
            For i = 1 To pSize
                x_prev(i) = 0.5 * (x(i) + tmp_vec(i) / node_deg(i))
            Next i
            
            Call zNormalize_Vec(x_prev, tmp_x)

            tmp_x = 0
            For i = 1 To pSize
                tmp_x = tmp_x + x(i) * x_prev(i)
            Next i
            If Abs(1 - tmp_x) < tol Then Exit For
        Next iterate
        For i = 1 To pSize
            pnode_pos(i, p) = x_prev(i)
        Next i

        For i = 1 To pSize
            xnorm(p) = xnorm(p) + (pnode_pos(i, p) ^ 2) * node_deg(i)
        Next i

    Next p

    tmp_x = pSize ^ (1# / n_dimension)
    For p = 1 To n_dimension
        For i = 1 To pSize
            pnode_pos(i, p) = pnode_pos(i, p) * tmp_x
        Next i
    Next p
    
    Application.StatusBar = False
End Sub


'=== Spectral Layout with subspace embedding speedup
'"Drawing Graphs by Eigenvectors: Theory and Practice", Koren (2005)
Private Sub Layout_Spectral_HDE(Optional n_dimension As Long = 2, Optional tol As Double = 0.001, _
        Optional isDegNorm As Boolean = False, Optional isWgt As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_y As Double, mu As Double
Dim n_subspace As Long
Dim x() As Double, deg As Variant, Lx() As Double
Dim eigen_vec() As Double, eigen_val() As Double, EdgeWgt() As Double
    n_subspace = 50
    If pSize < n_subspace Then n_subspace = pSize
    Call Layout_HDE(x, n_subspace)
    Call Layout_Orthonormalize(x, tol, isDegNorm, isWgt)
    n_subspace = UBound(x, 2)
    Call Calc_EdgeList
    ReDim EdgeWgt(1 To pn_edge)
    If isWgt = True Then
        Call Calc_Node_Deg_Wgt
        deg = pnode_deg_wgt
        EdgeWgt = pEdgeWgt
    Else
        Call Calc_Node_Deg
        deg = pnode_deg
        For k = 1 To pn_edge
            EdgeWgt(k) = 1
        Next k
    End If
    
    ReDim Lx(1 To pSize, 1 To n_subspace)
    For m = 1 To n_subspace
        For k = 1 To pn_edge
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            tmp_x = EdgeWgt(k)
            Lx(i, m) = Lx(i, m) - tmp_x * x(j, m)
            Lx(j, m) = Lx(j, m) - tmp_x * x(i, m)
        Next k
        For i = 1 To pSize
            Lx(i, m) = Lx(i, m) + deg(i) * x(i, m)
        Next i
    Next m

    Lx = modMath.M_Dot(x, Lx, 1)
    Call modMath.Eigen_Power(Lx, eigen_vec, eigen_val, 1)
    mu = eigen_val(1)
    For i = 1 To n_subspace - 1
        For j = i + 1 To n_subspace
            Lx(i, j) = -Lx(i, j)
            Lx(j, i) = Lx(i, j)
        Next j
    Next i
    For i = 1 To n_subspace
        Lx(i, i) = mu - Lx(i, i)
    Next i
    Call modMath.Eigen_Power(Lx, eigen_vec, eigen_val, n_dimension)
    pnode_pos = modMath.M_Dot(x, eigen_vec)
    Call modMath.Normalize_x(pnode_pos, x, Lx, "AVGSD")
    tmp_x = pSize ^ (1# / n_dimension)
    For j = 1 To n_dimension
        For i = 1 To pSize
            pnode_pos(i, j) = pnode_pos(i, j) * tmp_x
        Next i
    Next j
    Erase x, Lx, eigen_vec, eigen_val
End Sub

Private Sub Layout_HDE(x() As Double, Optional n_dimension As Long = 50)
Dim i As Long, j As Long, k As Long, n As Long, m As Long, p As Long
Dim INFINITY As Double, tmp_max As Double
Dim dist() As Double, xp() As Double
Dim i_seed As Long, v_seed As Long, node_list() As Long, isVisited() As Long
    ReDim x(1 To pSize, 1 To n_dimension)
    ReDim xp(1 To pSize)
    INFINITY = Exp(70)
    ReDim dist(1 To pSize)
    For i = 1 To pSize
        dist(i) = INFINITY
    Next i
    Call Calc_Key2Idx
    node_list = pNodeList.GetList
    i_seed = Int(Rnd() * (pSize - 1)) + 1
    v_seed = node_list(i_seed)
    For p = 1 To n_dimension
        Call Dijkstra_Algorithm(v_seed, xp, isVisited, False)
        i_seed = -1
        tmp_max = -INFINITY
        For i = 1 To pSize
            x(i, p) = xp(i)
            If xp(i) < dist(i) Then dist(i) = xp(i)
            If dist(i) > tmp_max Then
                tmp_max = dist(i)
                i_seed = i
            End If
        Next i
        v_seed = node_list(i_seed)
    Next p
    Erase dist, xp, node_list, isVisited
End Sub

Private Sub Layout_Orthonormalize(x() As Double, Optional tol As Double = 0.001, _
        Optional isDegNorm As Boolean = False, Optional isWgt As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, m As Long, p As Long, q As Long, n_dimension As Long
Dim p_idx() As Long, w As Variant
Dim tmp_x As Double, tmp_y As Double
Dim x0() As Double
    n_dimension = UBound(x, 2)
    If isDegNorm = True Then
        If isWgt = False Then
            Call Calc_Node_Deg
            w = pnode_deg
        Else
            Call Calc_Node_Deg_Wgt
            w = pnode_deg_wgt
        End If
    Else
        ReDim w(1 To pSize)
        For i = 1 To pSize
            w(i) = 1
        Next i
    End If
    ReDim x0(1 To pSize)
    For i = 1 To pSize
        x0(i) = 1# / Sqr(pSize * w(i))
    Next i

    k = 0
    ReDim p_idx(1 To n_dimension)
    For p = 1 To n_dimension
        tmp_x = 0
        For i = 1 To pSize
            tmp_x = tmp_x + x0(i) * x(i, p) * w(i)
        Next i
        For i = 1 To pSize
            x(i, p) = x(i, p) - tmp_x * x0(i)
        Next i
        If p > 1 Then
            For q = 1 To p - 1
                tmp_x = 0
                For i = 1 To pSize
                    tmp_x = tmp_x + x(i, q) * x(i, p) * w(i)
                Next i
                For i = 1 To pSize
                    x(i, p) = x(i, p) - tmp_x * x(i, q)
                Next i
            Next q
        End If
        tmp_x = 0
        For i = 1 To pSize
            tmp_x = tmp_x + (x(i, p) ^ 2) * w(i)
        Next i
        tmp_x = Sqr(tmp_x)
        If tmp_x < tol Then
            For i = 1 To pSize
                x(i, p) = 0
            Next i
        Else
            k = k + 1
            p_idx(k) = p
            For i = 1 To pSize
                x(i, p) = x(i, p) / tmp_x
            Next i
        End If
    Next p
    
    For p = 1 To k
        q = p_idx(p)
        For i = 1 To pSize
            x(i, p) = x(i, q)
        Next i
    Next p
    ReDim Preserve x(1 To pSize, 1 To k)
End Sub


'=== Force Directed algorithm with Barnes-Hut acceleration
'c_spring: spring constant, c_repulsive: repulsive strength, c_power: power law of repulsive force
Sub Layout_Force(Optional n_dimension As Long = 2, _
        Optional c_spring As Double = 1, Optional c_repulsive As Double = 1, Optional c_power As Double = 2, _
        Optional iter_max As Long = 500, Optional rescale_pos As Boolean = True, _
        Optional WgtType As String = "NONE")
Dim i As Long, j As Long, k As Long, iterate As Long, p As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp As Double
Dim temperature As Double, max_chg As Double, step_size As Double
Dim node_wgts() As Double, force_xy() As Double, delta() As Double
Dim quadtree1 As cqtree
Dim x_shift() As Double, x_scale() As Double, edge_len_avg As Double
Dim energy As Double, energy_prev As Double, energy_prev2 As Double, conv_count As Long
Dim isAdaptive As Boolean, c3 As Double
Dim gain() As Double
'Dim vArr As Variant, uArr As Variant, vTransform As Variant

    'Randomize node positions if not already exists
    isAdaptive = False
    If isAllocated(pnode_pos) = False Then
        Call node_pos_randomize(n_dimension): isAdaptive = True
    Else
        If UBound(pnode_pos, 2) <> n_dimension Or UBound(pnode_pos, 1) <> pSize Then
            Call node_pos_randomize(n_dimension): isAdaptive = True
        End If
    End If

    'if node_pos already exists, rescale to lie within a ranges order 1~10
    If rescale_pos = True Then
        tmp_x = c_spring * pSize ^ (1# / n_dimension)
        Call modMath.Normalize_x(pnode_pos, x_shift, x_scale, "MINMAX")
        For j = 1 To n_dimension
            For i = 1 To pSize
                pnode_pos(i, j) = pnode_pos(i, j) * tmp_x
            Next i
        Next j
    End If

    'Weights to use in scaling repulsive forces
    If WgtType = "NODE" Then
        node_wgts = pNodeList.GetListWgt
    ElseIf WgtType = "DEG" Then
        Call Calc_Node_Deg
        ReDim node_wgts(1 To pSize)
        For i = 1 To pSize
            node_wgts(i) = pnode_deg(i)
        Next i
    ElseIf WgtType = "DEG_WGT" Then
        Call Calc_Node_Deg_Wgt
        ReDim node_wgts(1 To pSize)
        For i = 1 To pSize
            node_wgts(i) = pnode_deg_wgt(i)
        Next i
    Else
        ReDim node_wgts(1 To pSize)
        For i = 1 To pSize
            node_wgts(i) = 1
        Next i
    End If
    tmp_x = 0
    For i = 1 To pSize
        tmp_x = tmp_x + node_wgts(i)
    Next i
    For i = 1 To pSize
        node_wgts(i) = pSize * node_wgts(i) / tmp_x
    Next i

    If pSize = 2 Then
        ReDim pnode_pos(1 To 2, 1 To n_dimension)
        For p = 1 To n_dimension
            pnode_pos(2, p) = c_spring / Sqr(n_dimension)
        Next p
    End If

    Call Calc_EdgeList
    temperature = c_spring
    c3 = c_repulsive * (c_spring ^ (c_power + 1))
    energy_prev = Exp(70): conv_count = 0
    ReDim delta(1 To n_dimension)

    For iterate = 1 To iter_max
'        If (iterate - 1) Mod 5 = 0 Then
'            DoEvents
'            With ActiveWorkbook.Sheets("Lattice")
'                .Range("R5:V1048576").Clear
'                .Range("R5").Resize(pSize, 2).Value = pnode_pos
'                .Range("U5").Resize(pn_edge * 3 - 1, 2).Value = Print_Edges
'            End With
'            Application.ScreenUpdating = True
'            Application.Calculation = xlCalculationAutomatic
'            DoEvents
'            Application.ScreenUpdating = False
'            Application.Calculation = xlCalculationManual
'        End If

        '=== Node-Node Repulsive Force
        If pSize > 64 Then
            Set quadtree1 = New cqtree
            force_xy = quadtree1.NetForce(pnode_pos, c3, , node_wgts, True, c_power)
            Set quadtree1 = Nothing
        Else
            ReDim force_xy(1 To pSize, 1 To n_dimension)
            For i = 1 To pSize - 1
                For j = i + 1 To pSize
                    tmp = 0
                    For p = 1 To n_dimension
                        delta(p) = pnode_pos(i, p) - pnode_pos(j, p)
                        tmp = tmp + delta(p) ^ 2
                    Next p
                    If tmp > 0 Then
                        tmp = c3 * node_wgts(i) * node_wgts(j) / (tmp ^ ((c_power + 1) / 2))
                        For p = 1 To n_dimension
                            force_xy(i, p) = force_xy(i, p) + delta(p) * tmp
                            force_xy(j, p) = force_xy(j, p) - delta(p) * tmp
                        Next p
                    End If
                Next j
            Next i
        End If

        '=== edge as spring force
        For k = 1 To UBound(pEdgeList, 1)
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            tmp = 0
            For p = 1 To n_dimension
                delta(p) = pnode_pos(i, p) - pnode_pos(j, p)
                tmp = tmp + delta(p) ^ 2
            Next p
            If tmp = 0 Then
                For p = 1 To n_dimension
                    delta(p) = c_spring * (-0.0005 + 0.001 * Rnd())
                    tmp = tmp + delta(p) ^ 2
                Next p
            End If
            tmp = Sqr(tmp)
            tmp_x = c3 * node_wgts(i) * node_wgts(j) / (tmp ^ (c_power + 1))
            tmp = (tmp - c_spring) / tmp
            For p = 1 To n_dimension
                force_xy(i, p) = force_xy(i, p) - tmp * delta(p) - delta(p) * tmp_x
                force_xy(j, p) = force_xy(j, p) + tmp * delta(p) + delta(p) * tmp_x
            Next p
        Next k

        '=== Update position
        max_chg = 0: energy = 0
        For i = 1 To pSize
            tmp = 0
            For p = 1 To n_dimension
                tmp = tmp + force_xy(i, p) ^ 2
            Next p
            energy = energy + tmp
            tmp = Sqr(tmp)

            step_size = tmp
            If step_size > temperature Then
                step_size = temperature
            End If
            If tmp > 0 Then
                For p = 1 To n_dimension
                    pnode_pos(i, p) = pnode_pos(i, p) + step_size * force_xy(i, p) / tmp
                Next p
            End If
            If step_size > max_chg Then max_chg = step_size
        Next i

        energy = energy / pSize
        
        If max_chg < (c_spring * 0.001) Then Exit For
        If energy < energy_prev Then
            conv_count = conv_count + 1
            If conv_count >= 5 Then
                If Abs(energy / energy_prev - 1) < 0.005 Then Exit For
            End If
        Else
            conv_count = 0
        End If
        If isAdaptive = False Then
            temperature = temperature * 0.9
        Else
            If energy < energy_prev Then
                If conv_count >= 5 Then
                    conv_count = 0
                    temperature = temperature / 0.95
                End If
            Else
                conv_count = 0
                temperature = temperature * 0.95
            End If
        End If
        energy_prev = energy

        If iterate Mod 50 = 0 Then
            DoEvents
            Application.StatusBar = "Layout_Force: " & iterate & "/" & iter_max
        End If
    Next iterate
'    Debug.Print pSize & ": " & iterate & "/" & iter_max
    If iterate >= iter_max Then Debug.Print "Force Directed Layout: Failed to achieve convergence."
    Erase force_xy
    Application.StatusBar = False
End Sub


'=== Force directed graph layout with multilevel speed up
'"A Multilevel Algorithm for Force-Directed Graph Drawing", Walshaw (2003)
Private Sub Layout_Force_Multi(Optional n_dimension As Long = 2, _
        Optional c_spring As Double = 1, Optional c_repulsive As Double = 0.2, Optional c_power As Double = 1, _
        Optional iter_max As Long = 500, Optional WgtType As String = "NONE")
Dim i As Long, j As Long, k As Long, n_graph As Long
Dim gs As Collection
Dim g1 As cGraph_Undir, g2 As cGraph_Undir
Dim x_pos() As Double, y_pos() As Double, idx_map() As Long, c2_tmp As Double, tmp_x As Double
Dim node_wgts() As Double, node_wgts_original() As Double, node_list() As Long
    
    Call pNodeList.GetListnWgt(node_list, node_wgts_original)
    'Weights to use in scaling repulsive forces
    If WgtType = "NODE" Then
        node_wgts = pNodeList.GetListWgt
    ElseIf WgtType = "DEG" Then
        Call Calc_Node_Deg
        ReDim node_wgts(1 To pSize)
        For i = 1 To pSize
            node_wgts(i) = pnode_deg(i)
        Next i
    ElseIf WgtType = "DEG_WGT" Then
        Call Calc_Node_Deg_Wgt
        ReDim node_wgts(1 To pSize)
        For i = 1 To pSize
            node_wgts(i) = pnode_deg_wgt(i)
        Next i
    Else
        ReDim node_wgts(1 To pSize)
        For i = 1 To pSize
            node_wgts(i) = 1
        Next i
    End If
    tmp_x = 0
    For i = 1 To pSize
        tmp_x = tmp_x + node_wgts(i)
    Next i
    For i = 1 To pSize
        node_wgts(i) = node_wgts(i) * pSize / tmp_x
    Next i
    
    'Temporarily replace node weights
    For i = 1 To pSize
        Call pNodeList.Append(node_list(i), node_wgts(i), 1)
    Next i
    
    'Repeatedly coarsen graph
    Set gs = New Collection
    Set g1 = Me
    Do
        Set g2 = New cGraph_Undir
        With g2
            Call .Copy_from(g1)
            Call .Multi_Collapse(idx_map)
            If g2.Size < 2 Or g2.Size = g1.Size Then Exit Do
        End With
        g1.Collapse_Map = idx_map
        gs.Add g2
        Set g1 = g2
    Loop While g2.Size > 2
    
    'Calculate layout starting from coarsest graph
    n_graph = gs.count
    For k = n_graph To 0 Step -1
        DoEvents
        Application.StatusBar = "Layout_Force_Multi: " & (n_graph - k) & "/" & n_graph
        If k > 0 Then
            Set g1 = gs(k)
            gs.Remove k
        Else
            Set g1 = Me
        End If
        If k < n_graph Then
            Randomize
            With g1
                ReDim y_pos(1 To .Size, 1 To n_dimension)
                idx_map = .Collapse_Map
                tmp_x = 2 ^ (1# / n_dimension)
                For j = 1 To n_dimension
                    For i = 1 To .Size
                        y_pos(i, j) = (x_pos(idx_map(i), j) + c2_tmp * (-0.01 + Rnd() * 0.02)) * tmp_x
                    Next i
                Next j
                .node_pos = y_pos
            End With
        End If
        c2_tmp = c_spring * (Sqr(7# / 4#) ^ k)    'Longer spring length for coarser graph
        If k > 0 Then
            With g1
                Call .Layout_Force(n_dimension, c2_tmp, c_repulsive, c_power, iter_max, False, "NODE")
                x_pos = .node_pos
                Call .Reset
            End With
            Set g1 = Nothing
        End If
    Next k
    DoEvents: Application.StatusBar = "Layout_Force_Multi: Final Layout...."
    Call Layout_Force(n_dimension, c_spring, c_repulsive, c_power, iter_max, False, WgtType)
    Erase x_pos, y_pos, idx_map, pCollapse_Map
    Set g1 = Nothing
    Set g2 = Nothing
    Set gs = Nothing
    
    'Restore node weights
    For i = 1 To pSize
        Call pNodeList.Append(node_list(i), node_wgts_original(i), 1)
    Next i
    
    Application.StatusBar = False
End Sub

'Collapse current graph by merging each vertex to its neighbor of
'smallest weight. Also return mapping index to its parent.
'idx_map(i)=j means that node i is collapsed into node j at the coarser level
Sub Multi_Collapse(idx_map() As Long)
Dim i As Long, j As Long, k As Long, n_neighbor As Long, v_min As Long, i_min As Long, u As Long, v As Long
Dim tmp_x As Double, wgt_min As Double
Dim iArr() As Long, isProcess() As Long, iParent() As Long
Dim size_old As Long, n_edge_old As Long, xArr() As Double
Dim x_time As Double, node_wgts_old() As Double
'    x_time = Timer
    Call Calc_Key2Idx
    Call pNodeList.GetListnWgt(idx_map, xArr)
    ReDim isProcess(1 To UBound(pKey2Idx))
    ReDim iParent(1 To UBound(pKey2Idx))
    ReDim node_wgts_old(1 To UBound(pKey2Idx))
    For i = 1 To pSize
        u = idx_map(i)
        iParent(u) = u
        node_wgts_old(u) = xArr(i)
    Next i
    Erase xArr
    
    size_old = pSize
    n_edge_old = pn_edge
    For i = 1 To size_old
        If i Mod 200 = 0 Then
            DoEvents: Application.StatusBar = "Collapsing Graph: " & i & "/" & size_old
        End If
        u = idx_map(i)
        If isProcess(u) = 0 Then
            isProcess(u) = 1
            iArr = pAdjLists(Key2Idx(u)).GetList
            n_neighbor = UBound(iArr)
            If n_neighbor > 0 Then
                v_min = iArr(1)
                If n_neighbor > 1 Then
                    wgt_min = node_wgts_old(v_min) 'NodeWgt(v_min)
                    For j = 2 To n_neighbor
                        tmp_x = node_wgts_old(iArr(j)) ' NodeWgt(iArr(j))
                        If tmp_x < wgt_min Then
                            wgt_min = tmp_x
                            v_min = iArr(j)
                        End If
                    Next j
                End If
                If u <> v_min Then
                    Call NodeMerge(u, v_min)
                    isProcess(v_min) = 1
                    iParent(v_min) = Multi_Parent(u, iParent)
                    node_wgts_old(u) = node_wgts_old(u) + node_wgts_old(v_min)
                End If
            End If
            If pSize = 2 Then Exit For
        End If
    Next i
    Erase isProcess, iArr, node_wgts_old

    Call Calc_Key2Idx
    For i = 1 To size_old
        u = Multi_Parent(idx_map(i), iParent)
        idx_map(i) = pKey2Idx(u)
    Next i
'    Debug.Print size_old & "->" & pSize & ", " & Round(Timer - x_time, 4)
    Application.StatusBar = False
End Sub

Private Function Multi_Parent(u As Long, iParent() As Long) As Long
    If u <> iParent(u) Then
        iParent(u) = Multi_Parent(iParent(u), iParent)
    End If
    Multi_Parent = iParent(u)
End Function


Function Print_Edges() As Variant
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim n_dimension As Long
Dim vArr As Variant
    n = 0
    n_dimension = UBound(pnode_pos, 2)
    Call Calc_EdgeList
    ReDim vArr(1 To pn_edge * 3 - 1, 1 To n_dimension)
    For k = 1 To pn_edge
        i = pEdgeList(k, 1)
        j = pEdgeList(k, 2)
        For m = 1 To n_dimension
            vArr(n + 1, m) = pnode_pos(i, m)
            vArr(n + 2, m) = pnode_pos(j, m)
        Next m
        n = n + 3
    Next k
    Print_Edges = vArr
    Erase vArr
End Function


'=============================================
'Graph algorithms
'=============================================
Function DFS(u As Long, Optional plot_Path As Variant, Optional show_path As Boolean = False, _
            Optional check_full As Boolean = False) As Long()
Dim i As Long, j As Long, k As Long, v As Long, m As Long, n As Long, v_prev As Long, v2 As Long
Dim u_idx As Long
Dim tmp_x As Double
Dim DFS_Order() As Long, DFS_parent() As Long, vPath() As Long, n_visited As Long
    
    Call Calc_Key2Idx
    u_idx = pKey2Idx(u)
    
    If u_idx = 0 Then  'no path can be shown
        Debug.Print "DFS Failed: " & u & " does not exist in graph."
        ReDim DFS_Order(1 To pSize)
        DFS = DFS_Order
        If show_path = True Then ReDim plot_Path(1 To 1, 1 To 2)
        Exit Function
    End If
    
    ReDim DFS_Order(1 To pSize)
    ReDim DFS_parent(1 To pSize)
    ReDim vPath(1 To pSize)
    n_visited = 0
    DFS_parent(u_idx) = -1
    Call DFS_Recursive(u_idx, DFS_Order, n_visited, DFS_parent, vPath)
    If check_full = True Then
        For i = 1 To pSize
            If DFS_Order(i) = 0 Then
                DFS_parent(i) = -1
                Call DFS_Recursive(i, DFS_Order, n_visited, DFS_parent, vPath)
            End If
        Next i
    End If
    DFS = DFS_Order
    
    ReDim Preserve vPath(1 To n_visited)
    If show_path = True Then
        v_prev = -1
        ReDim vArr(1 To 2, 0 To 0)
        For i = 1 To n_visited
            v = vPath(i)
            If v > 0 Then
                m = UBound(vArr, 2)
                If DFS_parent(v) <> v_prev Then
                    ReDim Preserve vArr(1 To 2, 0 To m + 3)
                    v2 = DFS_parent(v)
                    If v2 > 0 Then
                        vArr(1, m + 2) = pnode_pos(v2, 1)
                        vArr(2, m + 2) = pnode_pos(v2, 2)
                        vArr(1, m + 3) = pnode_pos(v, 1)
                        vArr(2, m + 3) = pnode_pos(v, 2)
                    Else
                        ReDim Preserve vArr(1 To 2, 0 To m + 2)
                        vArr(1, m + 2) = pnode_pos(v, 1)
                        vArr(2, m + 2) = pnode_pos(v, 2)
                    End If
                Else
                    ReDim Preserve vArr(1 To 2, 0 To m + 1)
                    vArr(1, m + 1) = pnode_pos(v, 1)
                    vArr(2, m + 1) = pnode_pos(v, 2)
                End If
                v_prev = v
            End If
        Next i

        ReDim plot_Path(1 To UBound(vArr, 2), 1 To 2)
        For i = 1 To UBound(vArr, 2)
            plot_Path(i, 1) = vArr(1, i)
            plot_Path(i, 2) = vArr(2, i)
        Next i
        Erase vArr, vPath
    End If
End Function


'*** Note: the input u here is node index, not key
Private Sub DFS_Recursive(u As Long, _
    DFS_Order() As Long, n_visited As Long, DFS_parent() As Long, vPath() As Long)
Dim k As Long, v As Long
Dim iArr() As Long
Dim uList As cAdjList
    n_visited = n_visited + 1
    DFS_Order(u) = n_visited
    vPath(n_visited) = u
    Set uList = pAdjLists(u)
    If uList.Size > 0 Then
        iArr = uList.GetList
        For k = 1 To UBound(iArr)
            v = pKey2Idx(iArr(k))
            If DFS_Order(v) = 0 And DFS_parent(v) = 0 Then
                DFS_parent(v) = u
                Call DFS_Recursive(v, DFS_Order, n_visited, DFS_parent, vPath)
            End If
        Next k
    End If
    Set uList = Nothing
    Erase iArr
End Sub


Function BFS(u As Long, Optional plot_Path As Variant, Optional show_path As Boolean = False) As Long()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim v As Long, v2 As Long, v_prev As Long, u_idx As Long
Dim iArr() As Long
Dim vList As cAdjList
Dim BFS_queue() As Long, isQueue() As Long, n_head As Long, n_tail As Long
Dim BFS_order() As Long, n_visited As Long, BFS_Parent() As Long, vPath() As Long
Dim vArr As Variant
    Call Calc_Key2Idx
    u_idx = pKey2Idx(u)
    
    ReDim isQueue(1 To pSize)
    ReDim BFS_queue(1 To 3 * pSize)
    ReDim BFS_order(1 To pSize)
    ReDim BFS_Parent(1 To pSize)
    ReDim vPath(1 To pSize)
    
    n_visited = 0
    n_head = 1: n_tail = 1
    BFS_queue(n_head) = u_idx: isQueue(u_idx) = 1
    BFS_Parent(u_idx) = -1
    Do While n_head <= n_tail
        v = BFS_queue(n_head): isQueue(v) = 0
        n_head = n_head + 1
        Set vList = pAdjLists(v)
        If vList.Size > 0 Then
            iArr = vList.GetList
            For i = 1 To UBound(iArr)
                j = pKey2Idx(iArr(i))
                If isQueue(j) = 0 And BFS_order(j) = 0 Then
                    BFS_Parent(j) = v
                    n_tail = n_tail + 1
                    BFS_queue(n_tail) = j: isQueue(j) = 1
                End If
            Next i
        End If
        If BFS_order(v) = 0 Then
            n_visited = n_visited + 1
            BFS_order(v) = n_visited
            vPath(n_visited) = v
        End If
    Loop
    
    BFS = BFS_order
    
    ReDim Preserve vPath(1 To n_visited)
    If show_path = True Then
        ReDim vArr(1 To 2, 0 To 0)
        v_prev = -1
        For k = 1 To n_visited
            v = vPath(k)
            If v > 0 Then
                If BFS_Parent(v) <> v_prev Then
                    v2 = BFS_Parent(v)
                    m = UBound(vArr, 2) + 3
                    ReDim Preserve vArr(1 To 2, 0 To m)
                    vArr(1, m - 1) = pnode_pos(v2, 1)
                    vArr(2, m - 1) = pnode_pos(v2, 2)
                    vArr(1, m) = pnode_pos(v, 1)
                    vArr(2, m) = pnode_pos(v, 2)
                Else
                    m = UBound(vArr, 2) + 1
                    ReDim Preserve vArr(1 To 2, 0 To m)
                    vArr(1, m) = pnode_pos(v, 1)
                    vArr(2, m) = pnode_pos(v, 2)
                End If
                v_prev = v
            End If
        Next k
        
        ReDim plot_Path(1 To UBound(vArr, 2), 1 To 2)
        For i = 1 To UBound(vArr, 2)
            plot_Path(i, 1) = vArr(1, i)
            plot_Path(i, 2) = vArr(2, i)
        Next i
        Erase vArr, vPath
        
    End If
End Function

Function component_count(comp_idx() As Long) As Long
Dim i As Long, j As Long, k As Long, v As Long, n_visited As Long
Dim iArr() As Long, jArr() As Long, comp_size() As Long, node_list() As Long
    node_list = NodeList
    n_visited = 0
    component_count = 0
    ReDim comp_idx(1 To pSize)
    For i = 1 To pSize
        If comp_idx(i) = 0 Then
            component_count = component_count + 1
            iArr = DFS(node_list(i), , , False)
            For j = 1 To pSize
                If iArr(j) > 0 And comp_idx(j) = 0 Then
                    comp_idx(j) = component_count
                    n_visited = n_visited + 1
                End If
            Next j
        End If
        If n_visited = pSize Then Exit For
    Next i

    ReDim comp_size(1 To component_count)
    For i = 1 To pSize
        comp_size(comp_idx(i)) = comp_size(comp_idx(i)) + 1
    Next i

    'Label largest component as #1
    Call modMath.Sort_Quick_A(comp_size, 1, component_count, iArr, 1)
    jArr = comp_idx
    For k = 1 To component_count
        For i = 1 To pSize
            If jArr(i) = iArr(k) Then
                comp_idx(i) = component_count - k + 1
            End If
        Next i
    Next k
    Erase iArr, jArr, comp_size, node_list
End Function

Sub Calc_ArticulationPoints(cut_pts() As Long, Optional outputkey As Boolean = True)
Dim i As Long, n_cut As Long, iArr() As Long
Dim isVisited() As Long, DFS_depth() As Long, LowPt() As Long, DFS_parent() As Long
    n_cut = 0
    Call Calc_Key2Idx
    ReDim cut_pts(0 To 0)
    ReDim isVisited(1 To pSize)
    ReDim DFS_depth(1 To pSize)
    ReDim LowPt(1 To pSize)
    ReDim DFS_parent(1 To pSize)
    For i = 1 To pSize
        If isVisited(i) = 0 Then
            Call GetArticulationPoints(i, 0, cut_pts, n_cut, isVisited, DFS_depth, LowPt, DFS_parent)
        End If
    Next i
    If outputkey = True Then
        iArr = pNodeList.GetList
        For i = 1 To n_cut
            cut_pts(i) = iArr(cut_pts(i))
        Next i
    End If
End Sub

Private Sub GetArticulationPoints(i As Long, u_depth As Long, cut_pts() As Long, n_cut As Long, _
            isVisited() As Long, DFS_depth() As Long, LowPt() As Long, DFS_parent() As Long)
Dim j As Long, k As Long, iArr() As Long
Dim childCount As Long, isArticulation As Boolean
    isVisited(i) = 1
    DFS_depth(i) = u_depth
    LowPt(i) = u_depth
    childCount = 0
    isArticulation = False
    iArr = pAdjLists(i).GetList
    For k = 1 To UBound(iArr)
        j = pKey2Idx(iArr(k))
        If isVisited(j) = 0 Then
            DFS_parent(j) = i
            Call GetArticulationPoints(j, u_depth + 1, cut_pts, n_cut, _
                    isVisited, DFS_depth, LowPt, DFS_parent)
            childCount = childCount + 1
            If LowPt(j) >= DFS_depth(i) Then isArticulation = True
            If LowPt(i) > LowPt(j) Then LowPt(i) = LowPt(j)
        ElseIf j <> DFS_parent(i) Then
            If LowPt(i) > DFS_depth(j) Then LowPt(i) = DFS_depth(j)
        End If
    Next k
    If (DFS_parent(i) > 0 And isArticulation = True) Or (DFS_parent(i) = 0 And childCount > 1) Then
        n_cut = n_cut + 1
        If n_cut = 1 Then
            ReDim cut_pts(1 To n_cut)
        Else
            ReDim Preserve cut_pts(1 To n_cut)
        End If
        cut_pts(n_cut) = i
    End If
End Sub


Sub Calc_Bridges(BridgeList() As Long, Optional outputkey As Boolean = True)
Dim i As Long, n_bridge As Long, iArr() As Long
Dim isVisited() As Long, DFS_Order() As Long, LowPt() As Long, DFS_parent() As Long
    Call Calc_Key2Idx
    n_bridge = 0
    ReDim BridgeList(1 To 2, 0 To 0)
    ReDim isVisited(1 To pSize)
    ReDim LowPt(1 To pSize)
    ReDim DFS_Order(1 To pSize)
    ReDim DFS_parent(1 To pSize)
    For i = 1 To pSize
        If isVisited(i) = 0 Then
            Call GetBridges(i, 0, BridgeList, n_bridge, isVisited, DFS_Order, LowPt, DFS_parent)
        End If
    Next i
    If n_bridge = 0 Then
        ReDim BridgeList(0 To 0, 1 To 2)
        Exit Sub
    End If
    iArr = BridgeList
    ReDim BridgeList(1 To n_bridge, 1 To 2)
    For i = 1 To n_bridge
        BridgeList(i, 1) = iArr(1, i)
        BridgeList(i, 2) = iArr(2, i)
    Next i
    If outputkey = True Then
        iArr = pNodeList.GetList
        For i = 1 To n_bridge
            BridgeList(i, 1) = iArr(BridgeList(i, 1))
            BridgeList(i, 2) = iArr(BridgeList(i, 2))
        Next i
    End If
End Sub

Private Sub GetBridges(i As Long, n_visited As Long, BridgeList() As Long, n_bridge As Long, _
            isVisited() As Long, DFS_Order() As Long, LowPt() As Long, DFS_parent() As Long)
Dim j As Long, k As Long, iArr() As Long
    isVisited(i) = 1
    DFS_Order(i) = n_visited
    LowPt(i) = n_visited
    iArr = pAdjLists(i).GetList
    For k = 1 To UBound(iArr)
        j = pKey2Idx(iArr(k))
        If isVisited(j) = 0 Then
            DFS_parent(j) = i
            Call GetBridges(j, n_visited + 1, BridgeList, n_bridge, _
                    isVisited, DFS_Order, LowPt, DFS_parent)
            If LowPt(i) > LowPt(j) Then LowPt(i) = LowPt(j)
            If LowPt(j) > DFS_Order(i) Then
                n_bridge = n_bridge + 1
                ReDim Preserve BridgeList(1 To 2, 0 To n_bridge)
                BridgeList(1, n_bridge) = i
                BridgeList(2, n_bridge) = j
            End If
        ElseIf j <> DFS_parent(i) Then
            If LowPt(i) > DFS_Order(j) Then LowPt(i) = DFS_Order(j)
        End If
    Next k
End Sub


'Make a new graph of each disconnected component and put into a collection
Sub Disconnected_Components(g_components As Collection, comp_idx() As Long, n_component As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim comp_size() As Long, iArr() As Long, jArr() As Long
Dim xArr() As Double, yArr() As Double
Dim node_keys() As Long, node_wgts() As Double
Dim g1 As cGraph_Undir, x As cAdjList
    Set g_components = New Collection
    n_component = component_count(comp_idx)
    ReDim comp_size(1 To n_component)
    For i = 1 To pSize
        j = comp_idx(i)
        comp_size(j) = comp_size(j) + 1
    Next i

    Call pNodeList.GetListnWgt(node_keys, node_wgts)
    For k = 1 To n_component
        j = 0
        ReDim iArr(1 To comp_size(k))
        ReDim xArr(1 To comp_size(k))
        For i = 1 To pSize
            If comp_idx(i) = k Then
                j = j + 1
                iArr(j) = node_keys(i)
                xArr(j) = node_wgts(i)
            End If
        Next i
        
        Set g1 = New cGraph_Undir
        With g1
            Call .Init(comp_size(k), xArr, True, iArr, True)
            For m = 1 To comp_size(k)
                i = iArr(m)
                jArr = node_neighbors(i)
                yArr = node_neighbors_wgt(i)
                For n = 1 To UBound(jArr)
                    Call .EdgeAdd(i, jArr(n), yArr(n), 0)
                Next n
            Next m
        End With
        g_components.Add g1
    Next k
    
    Erase node_keys, node_wgts, iArr, jArr, xArr, yArr
End Sub


'Find shortest distance from s to every other node
Sub Dijkstra_Algorithm(s As Long, dist() As Double, isVisited() As Long, _
        Optional isWgt As Boolean = True)
Dim i As Long, j As Long, k As Long, n As Long, u As Long, v As Long
Dim tmp_x As Double, INFINITY As Double
Dim iArr() As Long
Dim xArr() As Double
Dim q As cHeap
    INFINITY = Exp(70)
    ReDim isVisited(1 To pSize)
    ReDim dist(1 To pSize)
    For v = 1 To pSize
        dist(v) = INFINITY
    Next v
    Set q = New cHeap
    Call q.Init("MIN")
    i = pKey2Idx(s)
    Call q.Add(0, i)
    dist(i) = 0
    If isWgt = True Then
        Do While q.Size > 0
            Call q.Pop_Min(tmp_x, i)
            Call pAdjLists(i).GetListnWgt(iArr, xArr)
            For k = 1 To UBound(iArr)
                j = pKey2Idx(iArr(k))
                isVisited(j) = 1
                tmp_x = dist(i) + xArr(k)
                If tmp_x < dist(j) Then
                    dist(j) = tmp_x
                    Call q.Add(tmp_x, j)
                End If
            Next k
        Loop
    Else
        Do While q.Size > 0
            Call q.Pop_Min(tmp_x, i)
            iArr = pAdjLists(i).GetList
            For k = 1 To UBound(iArr)
                j = pKey2Idx(iArr(k))
                isVisited(j) = 1
                tmp_x = dist(i) + 1
                If tmp_x < dist(j) Then
                    dist(j) = tmp_x
                    Call q.Add(tmp_x, j)
                End If
            Next k
        Loop
    End If

    Call q.Reset
    Set q = Nothing
End Sub


Function Cycle_Chk() As Boolean
Dim i As Long, j As Long, k As Long, kk As Long, m As Long, n As Long, u As Long, v As Long
Dim tmp_x As Double
Dim sort_idx() As Long, xArr() As Long, node_list() As Long
Dim DFS_Order() As Long, n_visited As Long, DFS_parent() As Long, vPath() As Long
    Call Calc_Key2Idx
    Cycle_Chk = False
    node_list = pNodeList.GetList
    ReDim DFS_Order(1 To pSize)
    ReDim DFS_parent(1 To pSize)
    ReDim vPath(1 To pSize)
    n_visited = 0
    For i = 1 To pSize
        If DFS_Order(i) = 0 Then
            DFS_parent(i) = -1
            Call Cycle_Chk_Recursive(Cycle_Chk, i, DFS_Order, n_visited, DFS_parent, vPath)
            If Cycle_Chk = True Then Exit For
        End If
    Next i
End Function

Private Sub Cycle_Chk_Recursive(isCycle As Boolean, u As Long, _
    DFS_Order() As Long, n_visited As Long, DFS_parent() As Long, vPath() As Long)
Dim k As Long, v As Long
Dim iArr() As Long
Dim uList As cAdjList
    n_visited = n_visited + 1
    DFS_Order(u) = n_visited
    vPath(n_visited) = u
    Set uList = pAdjLists(u)
    If uList.Size > 0 Then
        iArr = uList.GetList
        For k = 1 To UBound(iArr)
            v = pKey2Idx(iArr(k))
            If DFS_Order(v) = 0 And DFS_parent(v) = 0 Then
                DFS_parent(v) = u
                Call Cycle_Chk_Recursive(isCycle, v, DFS_Order, n_visited, DFS_parent, vPath)
            ElseIf DFS_Order(v) > 0 And DFS_parent(u) <> v Then
                isCycle = True
            End If
        Next k
    End If
    Set uList = Nothing
    Erase iArr
End Sub



'========================================================
'Calculate network attributes
'========================================================

Sub Calc_Node_Deg()
Dim i As Long
    ReDim pnode_deg(1 To pSize)
    For i = 1 To pSize
        pnode_deg(i) = pAdjLists(i).Size
    Next i
End Sub

Sub Calc_Node_Deg_Wgt()
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double, tmp_y As Double, xArr() As Double
    ReDim pnode_deg_wgt(1 To pSize)
    For i = 1 To pSize
        tmp_x = 0
        xArr = pAdjLists(i).GetListWgt
        For k = 1 To UBound(xArr)
            tmp_x = tmp_x + xArr(k)
        Next k
        pnode_deg_wgt(i) = tmp_x
    Next i
End Sub

Sub Calc_Node_Reach()
Dim i As Long, n_visited As Long
Dim DFS_Order() As Long, DFS_parent() As Long, vPath() As Long
    Call Calc_Key2Idx
    ReDim pnode_reach(1 To pSize)
    For i = 1 To pSize
        ReDim DFS_Order(1 To pSize)
        ReDim DFS_parent(1 To pSize)
        ReDim vPath(1 To pSize)
        n_visited = 0
        DFS_parent(i) = -1
        Call DFS_Recursive(i, DFS_Order, n_visited, DFS_parent, vPath)
        pnode_reach(i) = n_visited - 1
    Next i
    Erase DFS_Order, DFS_parent, vPath
End Sub


Sub Calc_Node_PageRank(Optional damping As Double = 0.85, Optional isWgt As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim iterate As Long, iter_max As Long
Dim tmp_x As Double, tmp_y As Double, tol As Double, x_sum As Double, tmp_const As Double
Dim x() As Double, x_new() As Double, iArr() As Long, jArr() As Long, kArr() As Long
Dim xArr() As Double, yArr() As Double
Dim empty_list() As Long
    Randomize
    tol = 0.000000000001
    iter_max = 5000
    tmp_const = (1 - damping) / pSize
    ReDim x(1 To pSize)
    For i = 1 To pSize
        x(i) = 1# / pSize
    Next i

    'Create transition probability matrix
    Call Calc_EdgeList
    Call Calc_Node_Deg
    If isWgt = True Then Call Calc_Node_Deg_Wgt
    ReDim empty_list(0 To 0)
    For i = 1 To pSize
        If pnode_deg(i) = 0 Then
            j = UBound(empty_list) + 1
            ReDim Preserve empty_list(0 To j)
            empty_list(j) = i
        End If
    Next i
    
    For iterate = 1 To iter_max
        If (iterate - 1) Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = "Calc_Node_PageRank: " & iterate & "/" & iter_max
        End If
        ReDim x_new(1 To pSize)
        If isWgt = True Then
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                x_new(i) = x_new(i) + x(j) * pEdgeWgt(k) / pnode_deg_wgt(j)
                x_new(j) = x_new(j) + x(i) * pEdgeWgt(k) / pnode_deg_wgt(i)
            Next k
        Else
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                x_new(i) = x_new(i) + x(j) / pnode_deg(j)
                x_new(j) = x_new(j) + x(i) / pnode_deg(i)
            Next k
        End If
        tmp_x = 0
        For i = 1 To UBound(empty_list)
            tmp_x = tmp_x + x(empty_list(i))
        Next i
        tmp_x = damping * tmp_x / pSize
        
        For i = 1 To pSize
            x_new(i) = tmp_const + tmp_x + damping * x_new(i)
        Next i
        tmp_x = 0: tmp_y = 0
        For i = 1 To pSize
            tmp_x = tmp_x + x_new(i)
        Next i

        For i = 1 To pSize
            x_new(i) = x_new(i) / tmp_x
            tmp_y = tmp_y + Abs(x(i) - x_new(i))
        Next i
        x = x_new
        If tmp_y < tol Then Exit For
    Next iterate
    pnode_pagerank = x
    Application.StatusBar = False
End Sub


'Calculate eigenvector centrality for each disconnected component separately
Sub Calc_Node_Eigen(Optional isWgt As Boolean = False, Optional eigenval As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim g_components As Collection, comp_idx() As Long, n_component As Long
Dim xArr() As Double, iArr() As Long
Dim tmp_x As Double
Dim g1 As cGraph_Undir
    Call Disconnected_Components(g_components, comp_idx, n_component)
    If n_component = 1 Then
        Call Calc_Node_Eigen_Single(isWgt, eigenval)
        Exit Sub
    End If

    ReDim pnode_eigen(1 To pSize)
    Call Calc_Key2Idx
    For k = 1 To n_component
        Set g1 = g_components(k)
        With g1
            n = .Size
            Call .Calc_Node_Eigen_Single(isWgt)
            iArr = .NodeList
            xArr = .node_eigen
        End With
        tmp_x = Sqr(pSize * 1# / n)
        For i = 1 To UBound(iArr)
            j = pKey2Idx(iArr(i))
            pnode_eigen(j) = xArr(i) / tmp_x
        Next i
    Next k
    
    Erase iArr, xArr, comp_idx
    For k = n_component To 1 Step -1
        Set g1 = g_components(k)
        Call g1.Reset
        g_components.Remove k
        Set g1 = Nothing
    Next k
    Set g_components = Nothing
End Sub

Sub Calc_Node_Eigen_Single(Optional isWgt As Boolean = False, Optional eigenval As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim iterate As Long, iter_max As Long
Dim tmp_x As Double, tmp_y As Double, tol As Double
Dim x() As Double, x_new() As Double
    
    If pSize = 1 Then
        ReDim pnode_eigen(1 To 1)
        pnode_eigen(1) = 1
        Exit Sub
    End If
    
    Randomize
    tol = 0.000000000000001
    iter_max = 1000
    ReDim x(1 To pSize)
    For i = 1 To pSize
        x(i) = Rnd()
    Next i
    Call zNormalize_Vec(x, tmp_x)

    'Create Edge List
    Call Calc_EdgeList

    For iterate = 1 To iter_max

        ReDim x_new(1 To pSize)
        If isWgt = False Then
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                x_new(i) = x_new(i) + x(j)
                x_new(j) = x_new(j) + x(i)
            Next k
        Else
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                x_new(i) = x_new(i) + x(j) * pEdgeWgt(k)
                x_new(j) = x_new(j) + x(i) * pEdgeWgt(k)
            Next k
        End If

        tmp_x = 0: tmp_y = 0
        For i = 1 To pSize
            tmp_y = tmp_y + x_new(i) ^ 2
            tmp_x = tmp_x + x_new(i) * x(i)
        Next i
        tmp_y = Sgn(tmp_x) * Sqr(tmp_y)
        
        If tmp_y = 0 Then Exit For

        For i = 1 To pSize
            x_new(i) = x_new(i) / tmp_y
        Next i

        x = x_new
        If Abs(1 - tmp_x / tmp_y) < tol Then Exit For
    Next iterate
    pnode_eigen = x
    eigenval = tmp_y
End Sub


Sub Calc_Node_Katz(Optional alpha As Double = 0.1, Optional beta As Double = 1, _
        Optional isWgt As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim g_components As Collection, comp_idx() As Long, n_component As Long
Dim xArr() As Double, iArr() As Long
Dim tmp_x As Double
Dim g1 As cGraph_Undir
    Call Disconnected_Components(g_components, comp_idx, n_component)
    If n_component = 1 Then
        Call Calc_Node_Katz_Single(alpha, beta, isWgt)
        Exit Sub
    End If

    ReDim pnode_katz(1 To pSize)
    Call Calc_Key2Idx
    For k = 1 To n_component
        Set g1 = g_components(k)
        With g1
            n = .Size
            Call .Calc_Node_Katz_Single(alpha, beta, isWgt)
            iArr = .NodeList
            xArr = .node_katz
        End With
        tmp_x = Sqr(pSize * 1# / n)
        For i = 1 To UBound(iArr)
            j = pKey2Idx(iArr(i))
            pnode_katz(j) = xArr(i) / tmp_x
        Next i
    Next k
    
    Erase iArr, xArr, comp_idx
    For k = n_component To 1 Step -1
        Set g1 = g_components(k)
        Call g1.Reset
        g_components.Remove k
        Set g1 = Nothing
    Next k
    Set g_components = Nothing
End Sub

Sub Calc_Node_Katz_Single(Optional alpha As Double = 0.1, Optional beta As Double = 1, _
        Optional isWgt As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim iterate As Long, iter_max As Long
Dim tmp_x As Double, tmp_y As Double, tol As Double
Dim x() As Double, x_new() As Double, iArr() As Long, jArr() As Long, kArr() As Long
Dim xArr() As Double, yArr() As Double, lambda_max As Double, alpha1 As Double, beta1 As Double
Dim eigen_old() As Double, eigen_allocated As Boolean
    eigen_allocated = isAllocated(pnode_eigen)
    If eigen_allocated = True Then eigen_old = pnode_eigen
    Call Calc_Node_Eigen_Single(isWgt, lambda_max): Erase pnode_eigen
    If eigen_allocated = True Then pnode_eigen = eigen_old
    alpha1 = alpha: beta1 = beta / pSize
    If lambda_max > 0 Then
        If alpha1 > (1 / lambda_max) Then
            Debug.Print "Calc_Node_Katz: Invalid alpha. Max Eigenvalue=" & Round(lambda_max, 4) _
                    & ", alpha should be smaller than " & Round(1 / lambda_max, 4)
            alpha1 = (1 / lambda_max) * 0.9
        End If
    End If
    If pSize = 1 Then
        ReDim pnode_katz(1 To 1)
        pnode_katz(1) = 1
        Exit Sub
    End If

    Randomize
    tol = 0.0000000001
    iter_max = 1000
    ReDim x(1 To pSize)
    For i = 1 To pSize
        x(i) = Rnd()
    Next i
    Call zNormalize_Vec(x, tmp_x)

    'Create Edge List
    Call Calc_EdgeList
    
    For iterate = 1 To iter_max

        ReDim x_new(1 To pSize)
        If isWgt = True Then
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                x_new(i) = x_new(i) + x(j) * pEdgeWgt(k)
                x_new(j) = x_new(j) + x(i) * pEdgeWgt(k)
            Next k
        Else
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                x_new(i) = x_new(i) + x(j)
                x_new(j) = x_new(j) + x(i)
            Next k

        End If
        
        tmp_x = 0: tmp_y = 0
        For i = 1 To pSize
            x_new(i) = x_new(i) * alpha1 + beta1
            tmp_y = tmp_y + x_new(i) ^ 2
            tmp_x = tmp_x + x_new(i) * x(i)
        Next i
        tmp_y = Sgn(tmp_x) * Sqr(tmp_y)

        If tmp_y = 0 Then Exit For

        For i = 1 To pSize
            x_new(i) = x_new(i) / tmp_y
        Next i

        x = x_new
        If Abs(1 - tmp_x / tmp_y) < tol Then Exit For
    Next iterate
    If iterate > iter_max Then Debug.Print "Calc_Node_Katz: Failed to converge."
    pnode_katz = x
End Sub


Private Sub zNormalize_Vec(x() As Double, x_mag As Double)
Dim i As Long, n As Long
    n = UBound(x, 1)
    x_mag = 0
    For i = 1 To n
        x_mag = x_mag + x(i) ^ 2
    Next i
    x_mag = Sqr(x_mag)
    For i = 1 To n
        x(i) = x(i) / x_mag
    Next i
End Sub

Private Sub zNormalize_Vec_deg(x() As Double, x_mag As Double, deg As Variant)
Dim i As Long, n As Long
    n = UBound(x, 1)
    x_mag = 0
    For i = 1 To n
        x_mag = x_mag + deg(i) * (x(i) ^ 2)
    Next i
    x_mag = Sqr(x_mag)
    For i = 1 To n
        x(i) = x(i) / x_mag
    Next i
End Sub

Sub Calc_Node_Closeness(Optional isWgt As Boolean = True)
Dim i As Long, j As Long, k As Long, n As Long
Dim iArr() As Long, isVisited() As Long
Dim dist() As Double, tmp_x As Double
    Call Calc_Key2Idx
    ReDim pnode_closeness(1 To pSize)
    iArr = pNodeList.GetList
    For i = 1 To pSize
        If i Mod 10 = 0 Then DoEvents: Application.StatusBar = "cGraph_Undir:Calc_Node_Closeness: " & i & "/" & pSize
        Call Dijkstra_Algorithm(iArr(i), dist, isVisited, isWgt)
        n = 0
        tmp_x = 0
        For j = 1 To pSize
            If j <> i And isVisited(j) > 0 Then
                n = n + 1
                tmp_x = tmp_x + dist(j)
            End If
        Next j
        If n > 0 Then pnode_closeness(i) = ((n / pSize) ^ 2) * (pSize - 1) / tmp_x
    Next i
    Application.StatusBar = False
End Sub


'"A Faster Algorithm for Betweenness Centrality (2001)", Brandes (2001)
'"On Variants of Shortest-Path Betweenness Centrality and their Generic Computation.", Brandes (2008)
Sub Calc_Node_Betweenness(Optional isWgt As Boolean = False, Optional incl_endpts As Boolean = False, _
        Optional isNormalize As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, v As Long, w As Long
Dim n_stack As Long
Dim vStack() As Long, iArr() As Long, vPred As Variant
Dim sigma() As Double, delta() As Double, tmp_x As Double
    
    ReDim pnode_betweenness(1 To pSize)
    Call Calc_Key2Idx
    For i = 1 To pSize
        If i Mod 50 = 0 Then
            DoEvents: Application.StatusBar = "Calculating Betweenness: " & i & "/" & pSize
        End If
        
        If isWgt = False Then
            Call Calc_Node_Betweenness_BFS(i, vStack, n_stack, vPred, sigma)
        ElseIf isWgt = True Then
            Call Calc_Node_Betweenness_Dijkstra(i, vStack, n_stack, vPred, sigma)
        End If
        
        k = 0
        If incl_endpts = True Then
            pnode_betweenness(i) = pnode_betweenness(i) + n_stack - 1
            k = 1
        End If

        ReDim delta(1 To pSize)
        Do While n_stack > 0
            w = vStack(n_stack): n_stack = n_stack - 1
            iArr = vPred(w)
            tmp_x = (1 + delta(w)) / sigma(w)
            For j = 1 To UBound(iArr)
                v = iArr(j)
                delta(v) = delta(v) + sigma(v) * tmp_x
            Next j
            If w <> i Then pnode_betweenness(w) = pnode_betweenness(w) + delta(w) + k
        Loop
    
    Next i
    
    'Half the betweenness for undirected graph
    For i = 1 To pSize
        pnode_betweenness(i) = pnode_betweenness(i) / 2
    Next i
    Erase vStack, iArr, sigma, delta, vPred
    
    'Normalize by factor (n-1)*(n-2)
    If isNormalize = True Then
        If pSize > 2 Then
            m = (pSize - 1) * (pSize - 2)
            For i = 1 To pSize
                pnode_betweenness(i) = pnode_betweenness(i) / m
            Next i
        End If
    End If
    Application.StatusBar = False
End Sub

Private Sub Calc_Node_Betweenness_BFS(i As Long, vStack() As Long, n_stack As Long, vPred As Variant, sigma() As Double)
Dim j As Long, k As Long, m As Long, n As Long, v As Long, w As Long
Dim n_head As Long, n_tail As Long
Dim vQueue() As Long, dist() As Long, iArr() As Long, jArr() As Long
    ReDim vStack(1 To pSize * 3): n_stack = 0
    
    ReDim vPred(1 To pSize)
    ReDim iArr(0 To 0)
    For j = 1 To pSize
        vPred(j) = iArr
    Next j
    
    ReDim sigma(1 To pSize): sigma(i) = 1

    ReDim dist(1 To pSize)
    For j = 1 To pSize
        dist(j) = -1
    Next j
    dist(i) = 0
    
    ReDim vQueue(1 To pSize * 3)
    n_head = 1: n_tail = 1: vQueue(1) = i
    
    Do While n_tail >= n_head
        v = vQueue(n_head): n_head = n_head + 1
        n_stack = n_stack + 1: vStack(n_stack) = v
        iArr = pAdjLists(v).GetList
        For j = 1 To UBound(iArr)
            w = pKey2Idx(iArr(j))
            If dist(w) < 0 Then
                n_tail = n_tail + 1: vQueue(n_tail) = w
                dist(w) = dist(v) + 1
            End If
            If dist(w) = dist(v) + 1 Then
                sigma(w) = sigma(w) + sigma(v)
                jArr = vPred(w)
                m = UBound(jArr) + 1
                ReDim Preserve jArr(0 To m)
                jArr(m) = v
                vPred(w) = jArr
            End If
        Next j
    Loop
 
    Erase vQueue, iArr, jArr, dist
End Sub


Private Sub Calc_Node_Betweenness_Dijkstra(i As Long, vStack() As Long, n_stack As Long, vPred As Variant, sigma() As Double)
Dim j As Long, k As Long, m As Long, n As Long, v As Long, w As Long, n_count As Long
Dim iArr() As Long, jArr() As Long, xArr() As Double
Dim dist() As Double
Dim vQueue As cHeap, heap_tuple() As Long, isSeen() As Double
Dim tmp_x As Double, vw_dist As Double

    ReDim vStack(1 To pSize * 3): n_stack = 0
    
    ReDim vPred(1 To pSize)
    ReDim iArr(0 To 0)
    For j = 1 To pSize
        vPred(j) = iArr
    Next j
    
    ReDim sigma(1 To pSize): sigma(i) = 1

    ReDim dist(1 To pSize)
    ReDim isSeen(1 To pSize)
    For j = 1 To pSize
        dist(j) = -1
        isSeen(j) = -1
    Next j
    isSeen(i) = 0
    
    n_count = 0
    Set vQueue = New cHeap
    ReDim heap_tuple(1 To 2, 0 To 0)
    Call vQueue.Init("MIN")
    Call MinHeap_Add_Tuple(vQueue, 0, n_count, i, i, heap_tuple)
    
    Do While vQueue.Size > 0
        Call MinHeap_Pop_Tuple(vQueue, tmp_x, j, w, v, heap_tuple)
        If dist(v) < 0 Then
            sigma(v) = sigma(v) + sigma(w)
            n_stack = n_stack + 1: vStack(n_stack) = v
            dist(v) = tmp_x
            Call pAdjLists(v).GetListnWgt(iArr, xArr)
            For j = 1 To UBound(iArr)
                w = pKey2Idx(iArr(j))
                vw_dist = tmp_x + xArr(j)
                If dist(w) < 0 And (isSeen(w) < 0 Or vw_dist < isSeen(w)) Then
                    isSeen(w) = vw_dist
                    Call MinHeap_Add_Tuple(vQueue, vw_dist, n_count, v, w, heap_tuple)
                    sigma(w) = 0
                    ReDim jArr(0 To 1)
                    jArr(1) = v
                    vPred(w) = jArr
                ElseIf vw_dist = isSeen(w) Then
                    sigma(w) = sigma(w) + sigma(v)
                    jArr = vPred(w)
                    m = UBound(jArr) + 1
                    ReDim Preserve jArr(0 To m)
                    jArr(m) = v
                    vPred(w) = jArr
                End If
            Next j
        End If
    Loop
    Call vQueue.Reset
    Erase heap_tuple, iArr, jArr, xArr, dist, isSeen, heap_tuple
End Sub


'"An O(m) algorithm for Cores Decomposition of Networks", Vladimir Batagelj (2003)
Sub Calc_Node_Cores()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long, jj As Long
Dim w As Long, deg_max As Long, deg_u As Long, pos_u As Long, pos_w As Long
Dim deg_bin() As Long, vpos() As Long, vert() As Long, neighbor_list() As Long
Dim deg() As Long

    Call Calc_Key2Idx
    Call Calc_Node_Deg
    deg_max = -1
    For i = 1 To pSize
        If pnode_deg(i) > deg_max Then deg_max = pnode_deg(i)
    Next i
    
    ReDim deg_bin(0 To deg_max)
    For i = 1 To pSize
        deg_bin(pnode_deg(i)) = deg_bin(pnode_deg(i)) + 1
    Next i

    i = 1
    For k = 0 To deg_max
        j = deg_bin(k)
        deg_bin(k) = i
        i = i + j
    Next k

    ReDim vpos(1 To pSize)
    ReDim vert(1 To pSize)
    For i = 1 To pSize
        k = pnode_deg(i)
        vpos(i) = deg_bin(k)
        vert(vpos(i)) = i
        deg_bin(k) = deg_bin(k) + 1
    Next i
    For k = deg_max To 1 Step -1
        deg_bin(k) = deg_bin(k - 1)
    Next k
    deg_bin(0) = 1

    deg = pnode_deg
    For ii = 1 To pSize
        i = vert(ii)
        neighbor_list = pAdjLists(i).GetList
        For jj = 1 To UBound(neighbor_list)
            j = pKey2Idx(neighbor_list(jj))
            If deg(j) > deg(i) Then
                deg_u = deg(j): pos_u = vpos(j)
                pos_w = deg_bin(deg_u): w = vert(pos_w)
                If j <> w Then
                    vpos(j) = pos_w: vert(pos_u) = w
                    vpos(w) = pos_u: vert(pos_w) = j
                End If
                deg_bin(deg_u) = deg_bin(deg_u) + 1
                deg(j) = deg(j) - 1
            End If
        Next jj
    Next ii
    
    pnode_core = deg
End Sub


Sub Calc_Node_Onion()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim n_core As Long, n_layer As Long, n_process As Long
Dim isProcess() As Boolean, node_deg() As Long, node_deg_prev() As Long, iArr() As Long
Dim node_list() As Long, node_list_inv() As Long
Dim min_deg As Long
    Call Calc_Key2Idx
    Call Calc_Node_Deg
    node_deg = pnode_deg
    ReDim pnode_core(1 To pSize)
    ReDim pnode_onion(1 To pSize)
    ReDim isProcess(1 To pSize)

    'Sort node by degree in descending order
    Call modMath.Sort_Quick_A(node_deg, 1, UBound(node_deg), node_list, 1)
    Call Reverse_vec(node_list)
    Call Reverse_vec(node_deg)
    ReDim node_list_inv(1 To pSize)
    For i = 1 To UBound(node_list)
        node_list_inv(node_list(i)) = i
    Next i

    n_core = 1: n_layer = 1: n_process = 0
    Do While n_process < pSize
        DoEvents
        node_deg_prev = node_deg
        For ii = UBound(node_list) To 1 Step -1
            i = node_list(ii)
            If node_deg_prev(ii) > n_core Then
                ReDim Preserve node_deg(1 To ii)
                ReDim Preserve node_list(1 To ii)
                Exit For
            End If
            pnode_core(i) = n_core
            pnode_onion(i) = n_layer
            iArr = pAdjLists(i).GetList
            For k = 1 To UBound(iArr)
                j = pKey2Idx(iArr(k))
                If isProcess(j) = False Then node_deg(node_list_inv(j)) = node_deg(node_list_inv(j)) - 1
            Next k
            isProcess(i) = True
            n_process = n_process + 1
        Next ii

        'Sort decremented degrees
        Call modMath.Sort_Quick_A(node_deg, 1, UBound(node_deg), node_list, 0)
        Call Reverse_vec(node_list)
        Call Reverse_vec(node_deg)
        ReDim node_list_inv(1 To pSize)
        For i = 1 To UBound(node_list)
            node_list_inv(node_list(i)) = i
        Next i
        
        'Increment layer and core number
        n_layer = n_layer + 1
        min_deg = node_deg(UBound(node_deg))
        If min_deg >= (n_core + 1) Then n_core = min_deg
    Loop
    Erase node_deg, node_deg_prev, isProcess, iArr
End Sub


Sub Calc_Node_Cluster_Coeff()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long, jj As Long
Dim iArr() As Long, tmp_x As Double
Dim vList As cAdjList
Dim AdjMatrix() As Long
    
    If pn_edge > ((pSize * (pSize - 1) / 2) * 0.02) Then
        'Graph is not sparse, use matrix calculation
        ReDim AdjMatrix(1 To pSize, 1 To pSize)
        Call Calc_EdgeList
        For k = 1 To pn_edge
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            If i <> j Then
                AdjMatrix(i, j) = 1
                AdjMatrix(j, i) = 1
            End If
        Next k
        Call MMULT_Cube_Diag(AdjMatrix, pnode_cluster_coeff)
        Erase AdjMatrix
        pcluster_coeff = 0
        For i = 1 To pSize
            pcluster_coeff = pcluster_coeff + pnode_cluster_coeff(i)
        Next i
        pcluster_coeff = pcluster_coeff / pSize
        Exit Sub
    End If
    
    'Directly count edges if graph is sparse enough
    Call Calc_Key2Idx
    pcluster_coeff = 0
    ReDim pnode_cluster_coeff(1 To pSize)
    For i = 1 To pSize
        If i Mod 100 = 0 Then
            DoEvents
            Application.StatusBar = "Calc_Node_Cluster_Coeff: " & i & "/" & pSize
        End If
        iArr = pAdjLists(i).GetList
        k = UBound(iArr)
        If k > 1 Then
            m = 0
            For ii = 1 To k - 1
                Set vList = pAdjLists(pKey2Idx(iArr(ii)))
                For jj = ii + 1 To k
                    If vList.isConnect(iArr(jj), tmp_x) > 0 Then m = m + 1
                Next jj
            Next ii
            pnode_cluster_coeff(i) = m * 2# / (k * (k - 1))
        End If
        pcluster_coeff = pcluster_coeff + pnode_cluster_coeff(i)
    Next i
    pcluster_coeff = pcluster_coeff / pSize
    Application.StatusBar = False
End Sub


Sub Calc_Node_Cluster_Coeff_Wgt()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, ii As Long, jj As Long, kk As Long
Dim iii As Long, jjj As Long, iterate As Long, n_cut As Long
Dim iArr() As Long, xArr() As Double, yArr() As Double, sort_idx() As Long, edge_wgt_cut As Double, dw As Double
Dim coeff() As Double
Dim edge_values() As Double, edge_values_count() As Long
Dim tmp_x As Double
Dim g1 As cGraph_Undir
    Call Calc_EdgeList
    xArr = pEdgeWgt
    Call modMath.Sort_Quick_A(xArr, 1, pn_edge, sort_idx)
    
    'Find unique values of edge weights in ascending order
    n_cut = 1
    ReDim edge_values(1 To pn_edge)
    ReDim edge_values_count(1 To pn_edge)
    edge_values(1) = xArr(1)
    edge_values_count(1) = 1
    For k = 2 To pn_edge
        If xArr(k) > edge_values(n_cut) Then
            n_cut = n_cut + 1
            edge_values(n_cut) = xArr(k)
            edge_values_count(n_cut) = 1
        Else
            edge_values_count(n_cut) = edge_values_count(n_cut) + 1
        End If
    Next k
    ReDim Preserve edge_values(1 To n_cut)
    ReDim Preserve edge_values_count(1 To n_cut)
    
    Set g1 = New cGraph_Undir
    Call g1.Copy_from(Me)
    
    k = 0
    ReDim pnode_cluster_coeff_wgt(1 To pSize)
    For iterate = 1 To n_cut
        DoEvents
        Application.StatusBar = "Calc_Node_Cluster_Coeff_Wgt: " & iterate & "/" & n_cut
        edge_wgt_cut = edge_values(iterate)
        tmp_x = edge_values_count(iterate) * 1# / pn_edge
        
        With g1
            If iterate > 1 Then
                For i = k + 1 To k + edge_values_count(iterate)
                    Call .EdgeDelete(pEdgeList(sort_idx(i), 1), pEdgeList(sort_idx(i), 2))
                Next i
                k = k + edge_values_count(iterate)
            End If
            Call .Calc_Node_Cluster_Coeff
            coeff = .node_cluster_coeff
        End With

        For i = 1 To pSize
            pnode_cluster_coeff_wgt(i) = pnode_cluster_coeff_wgt(i) + coeff(i) * tmp_x
        Next i

    Next iterate
    Call g1.Reset
    Set g1 = Nothing
    
    pcluster_coeff_wgt = 0
    For i = 1 To pSize
        pcluster_coeff_wgt = pcluster_coeff_wgt + pnode_cluster_coeff_wgt(i)
    Next i
    pcluster_coeff_wgt = pcluster_coeff_wgt / pSize
    Application.StatusBar = False
End Sub

'========================================================
'Graph Filter
'========================================================

'Construct Minimum Spanning Tree from graph. If isMax is
'set to TRUE then Maximum Spanning Tree is made instead.
Sub MST_Build(g1 As cGraph_Undir, Optional isMax As Boolean = False)
Dim i As Long, j As Long, k As Long, kk As Long, m As Long, n As Long
Dim n_component As Long, u As Long, v As Long, ii As Long
Dim tmp_x As Double
Dim sort_idx() As Long, node_list() As Long, uparent() As Long
    n_component = component_count(sort_idx)
    node_list = pNodeList.GetList
    Call Calc_EdgeList
    Call modMath.Sort_Quick_A(pEdgeWgt, 1, pn_edge, sort_idx, 1)
    Set g1 = New cGraph_Undir
    Call g1.Init(pSize, Me.NodeWgts, True, node_list, True)
    n = 0
    ReDim uparent(1 To pSize)
    For kk = 1 To pn_edge
        If n Mod 500 = 0 Then
            DoEvents
            Application.StatusBar = "cGraph_Undir:MST_Build: " & n & "/" & pSize - n_component
        End If
        If isMax = False Then
            k = sort_idx(kk)
            tmp_x = pEdgeWgt(kk)
        Else
            k = sort_idx(pn_edge - kk + 1)
            tmp_x = pEdgeWgt(pn_edge - kk + 1)
        End If
        i = pEdgeList(k, 1)
        j = pEdgeList(k, 2)
        If uparent(i) <> uparent(j) Or (uparent(i) = 0 And uparent(j) = 0) Then
            Call g1.EdgeAdd(node_list(i), node_list(j), tmp_x)
            'Union the sets that contain i & j, use i as new parent key
            n = n + 1
            ii = uparent(i): uparent(i) = i
            If ii > 0 And ii <> i Then
                For m = 1 To pSize
                    If uparent(m) = ii Then uparent(m) = i
                Next m
            End If
            ii = uparent(j): uparent(j) = i
            If ii > 0 Then
                For m = 1 To pSize
                    If uparent(m) = ii Then uparent(m) = i
                Next m
            End If
        End If
        If n = pSize - n_component Then Exit For
    Next kk
    Erase sort_idx, node_list, uparent
    Erase pEdgeList, pEdgeWgt
    Application.StatusBar = False
End Sub


'Build a directed DFS graph rooted at u
Sub DFS_Graph_Build(u As Long, g1 As cGraph_Dir, Optional check_full As Boolean = False)
Dim i As Long, j As Long, k As Long, v As Long, m As Long, n As Long, v_prev As Long, v2 As Long
Dim u_idx As Long
Dim tmp_x As Double
Dim node_list() As Long
Dim DFS_Order() As Long, DFS_parent() As Long, vPath() As Long, n_visited As Long
Dim n_bkwd As Long, EdgeList_Bkwd() As Long, EdgeWgt_Bkwd() As Double
Dim n_fwd As Long, EdgeList_Fwd() As Long, EdgeWgt_Fwd() As Double

    Call Calc_Key2Idx
    u_idx = pKey2Idx(u)
    
    If u_idx = 0 Then  'no path can be shown
        Debug.Print "DFS_Graph_Build: " & u & " does not exist in graph."
        Exit Sub
    End If
    
    ReDim DFS_Order(1 To pSize)
    ReDim DFS_parent(1 To pSize)
    ReDim vPath(1 To pSize)
    n_visited = 0
    n_fwd = 0: ReDim EdgeList_Fwd(1 To 2, 0 To 0): ReDim EdgeWgt_Fwd(0 To 0)
    n_bkwd = 0: ReDim EdgeList_Bkwd(1 To 2, 0 To 0): ReDim EdgeWgt_Bkwd(0 To 0)
    
    DFS_parent(u_idx) = -1
    Call DFS_Graph_Build_Recursive(u_idx, DFS_Order, n_visited, DFS_parent, vPath, _
                n_fwd, EdgeList_Fwd, EdgeWgt_Fwd, n_bkwd, EdgeList_Bkwd, EdgeWgt_Bkwd)
    If check_full = True Then
        For i = 1 To pSize
            If DFS_Order(i) = 0 Then
                DFS_parent(i) = -1
                Call DFS_Graph_Build_Recursive(i, DFS_Order, n_visited, DFS_parent, vPath, _
                        n_fwd, EdgeList_Fwd, EdgeWgt_Fwd, n_bkwd, EdgeList_Bkwd, EdgeWgt_Bkwd)
            End If
        Next i
    End If
    If n_fwd > 0 Then Call TransposeTrim(EdgeList_Fwd)
    If n_bkwd > 0 Then Call TransposeTrim(EdgeList_Bkwd)
    
    node_list = pNodeList.GetList
    For k = 1 To n_fwd
        EdgeList_Fwd(k, 1) = node_list(EdgeList_Fwd(k, 1))
        EdgeList_Fwd(k, 2) = node_list(EdgeList_Fwd(k, 2))
    Next k
    For k = 1 To n_bkwd
        EdgeList_Bkwd(k, 1) = node_list(EdgeList_Bkwd(k, 1))
        EdgeList_Bkwd(k, 2) = node_list(EdgeList_Bkwd(k, 2))
    Next k
    
    Set g1 = New cGraph_Dir
    With g1
        Call .Init(pSize, Me.NodeWgts, True, pNodeList.GetList, True)
        If n_fwd > 0 Then Call .EdgeAddList(EdgeList_Fwd, EdgeWgt_Fwd, True)
        If n_bkwd > 0 Then Call .EdgeAddList(EdgeList_Bkwd, EdgeWgt_Bkwd, True)
    End With
 
End Sub


Private Sub DFS_Graph_Build_Recursive(u As Long, _
    DFS_Order() As Long, n_visited As Long, DFS_parent() As Long, vPath() As Long, _
    n_fwd As Long, EdgeList_Fwd() As Long, EdgeWgt_Fwd() As Double, _
    n_bkwd As Long, EdgeList_Bkwd() As Long, EdgeWgt_Bkwd() As Double)
Dim k As Long, v As Long
Dim iArr() As Long, xArr() As Double
Dim uList As cAdjList
    n_visited = n_visited + 1
    DFS_Order(u) = n_visited
    vPath(n_visited) = u
    Set uList = pAdjLists(u)
    If uList.Size > 0 Then
        Call uList.GetListnWgt(iArr, xArr)
        For k = 1 To UBound(iArr)
            v = pKey2Idx(iArr(k))
            If DFS_Order(v) = 0 And DFS_parent(v) = 0 Then
                DFS_parent(v) = u
                n_fwd = n_fwd + 1
                ReDim Preserve EdgeList_Fwd(1 To 2, 0 To n_fwd)
                ReDim Preserve EdgeWgt_Fwd(0 To n_fwd)
                EdgeList_Fwd(1, n_fwd) = u
                EdgeList_Fwd(2, n_fwd) = v
                EdgeWgt_Fwd(n_fwd) = xArr(k)
                Call DFS_Graph_Build_Recursive(v, DFS_Order, n_visited, DFS_parent, vPath, _
                    n_fwd, EdgeList_Fwd, EdgeWgt_Fwd, n_bkwd, EdgeList_Bkwd, EdgeWgt_Bkwd)
            ElseIf DFS_Order(v) > 0 And DFS_parent(u) <> v And DFS_Order(v) < DFS_Order(u) Then
                n_bkwd = n_bkwd + 1
                ReDim Preserve EdgeList_Bkwd(1 To 2, 0 To n_bkwd)
                ReDim Preserve EdgeWgt_Bkwd(0 To n_bkwd)
                EdgeList_Bkwd(1, n_bkwd) = u
                EdgeList_Bkwd(2, n_bkwd) = v
                EdgeWgt_Bkwd(n_bkwd) = xArr(k)
            End If
        Next k
    End If
    Set uList = Nothing
    Erase iArr
End Sub


Private Sub TransposeTrim(iArr() As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim jArr() As Long
    ReDim jArr(1 To UBound(iArr, 2), 1 To UBound(iArr, 1))
    For i = 1 To UBound(iArr, 1)
        For j = 1 To UBound(iArr, 2)
            jArr(j, i) = iArr(i, j)
        Next j
    Next i
    iArr = jArr
End Sub


'===================
'Extracting the multiscale backbone of xomplex weighted networks, M. Angeles Serrano (2008)
Sub Filter_Disparity(g1 As cGraph_Undir, Optional significance As Double = 0.1, _
        Optional InvertWgt As Boolean = False)
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double, tmp_y As Double
Dim node_list() As Long
    node_list = NodeList
    Call Calc_Node_Deg
    Call Calc_Node_Deg_Wgt
    Call Calc_EdgeList
    
    If InvertWgt = True Then
        ReDim pnode_deg_wgt(1 To pSize)
        For k = 1 To pn_edge
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            pEdgeWgt(k) = 1 / pEdgeWgt(k)
            pnode_deg_wgt(i) = pnode_deg_wgt(i) + pEdgeWgt(k)
            pnode_deg_wgt(j) = pnode_deg_wgt(j) + pEdgeWgt(k)
        Next k
    End If

    Set g1 = New cGraph_Undir
    With g1
        Call .Init(pSize, NodeWgts, True, NodeList, True)
        For k = 1 To pn_edge
            If k Mod 500 = 0 Then
                DoEvents: Application.StatusBar = "Disparity Filter " & k & "/" & pn_edge
            End If
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            tmp_y = pEdgeWgt(k)
            If pnode_deg(i) = 1 And pnode_deg(j) > 1 Then
                tmp_x = (1 - tmp_y / pnode_deg_wgt(j)) ^ (pnode_deg(j) - 1)
                If tmp_x < significance Then Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
            ElseIf pnode_deg(j) = 1 And pnode_deg(i) > 1 Then
                tmp_x = (1 - tmp_y / pnode_deg_wgt(i)) ^ (pnode_deg(i) - 1)
                If tmp_x < significance Then Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
            ElseIf pnode_deg(j) = 1 And pnode_deg(i) = 1 Then
                Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
            Else
                tmp_x = (1 - tmp_y / pnode_deg_wgt(i)) ^ (pnode_deg(i) - 1)
                If tmp_x < significance Then
                    Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
                Else
                    tmp_x = (1 - tmp_y / pnode_deg_wgt(j)) ^ (pnode_deg(j) - 1)
                    If tmp_x < significance Then
                        Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
                    End If
                End If
            End If
        Next k
    End With
    Application.StatusBar = False
End Sub



Sub Filter_EdgeCount(g1 As cGraph_Undir, Optional n_edge As Long = -1)
Dim i As Long, j As Long, k As Long, kk As Long
Dim n_edge_new As Long
Dim sort_idx() As Long, xArr() As Double
Dim EdgeList_new() As Long, EdgeWgt_new() As Double

    If n_edge > 0 Then
        n_edge_new = n_edge
    Else
        n_edge_new = pSize * 3
    End If
    If n_edge_new > pn_edge Then n_edge_new = pn_edge

    Call Calc_EdgeList
    xArr = pEdgeWgt
    Call modMath.Sort_Quick_A(xArr, 1, UBound(xArr, 1), sort_idx)
    
    ReDim EdgeList_new(1 To n_edge_new, 1 To 2)
    ReDim EdgeWgt_new(1 To n_edge_new)
    For kk = 1 To n_edge_new
        k = sort_idx(pn_edge - kk + 1)
        EdgeList_new(kk, 1) = pEdgeList(k, 1)
        EdgeList_new(kk, 2) = pEdgeList(k, 2)
        EdgeWgt_new(kk) = pEdgeWgt(k)
    Next kk
    
    Erase sort_idx, xArr
    
    Set g1 = New cGraph_Undir
    With g1
        Call .Init(pSize, Me.NodeWgts, True, Me.NodeList, True)
        Call .EdgeAddList(EdgeList_new, EdgeWgt_new, True)
    End With
End Sub


Private Function isAllocated(vArr As Variant) As Boolean
Dim i As Long
    isAllocated = False
    If VBA.IsArray(vArr) = False Then Exit Function
    On Error Resume Next
    i = UBound(vArr, 1)
    If Err.Number = 0 Then
        If LBound(vArr) <= UBound(vArr) Then
            isAllocated = True
        Else
            isAllocated = False
        End If
    Else
        isAllocated = False
    End If
End Function




'Find clusters in graph based on edge strength
'"Software component capture usign graph clustering", Yves Chiricota
'Output: function itself returns an integer giving number of clusters discovered.
'        cluster_idx(1:N), cluster membership index of each node
Function Cluster_by_EdgeStrength(cluster_idx() As Long) As Long
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_cluster As Long
Dim node_list() As Long, iArr() As Long, comp_idx() As Long, sort_idx() As Long, x() As Double
Dim tmp_x As Double, tmp_y As Double, quality_max As Double
Dim g1 As cGraph_Undir
    node_list = pNodeList.GetList
    Call Calc_Key2Idx
    
    'Calculate edge strength of each edge
    Call Calc_EdgeList
    ReDim x(1 To UBound(pEdgeList, 1))
    For k = 1 To UBound(pEdgeList, 1)
        x(k) = Edge_Strength_uv(node_list(pEdgeList(k, 1)), node_list(pEdgeList(k, 2)))
    Next k
    
    'Create clusters by removing weak edges until maximum
    'cluster quality is achieved.
    Call modMath.Sort_Quick_A(x, 1, pn_edge, sort_idx, 1)
    Set g1 = New cGraph_Undir
    Call g1.Copy_from(Me)
    quality_max = -Exp(70)
    k = 1
    Do
        tmp_x = x(k)
        With g1
            Do
                If x(k) <= tmp_x Then
                    Call .EdgeDelete(pEdgeList(sort_idx(k), 1), pEdgeList(sort_idx(k), 2))
                Else
                    Exit Do
                End If
                k = k + 1
            Loop While k < pn_edge
            n_cluster = .component_count(comp_idx)
        End With
        tmp_y = Cluster_Quality_MQ(comp_idx, n_cluster)
        If tmp_y >= quality_max Then
            quality_max = tmp_y
            cluster_idx = comp_idx
            Cluster_by_EdgeStrength = n_cluster
        End If
    Loop While g1.n_edge > 0
    Call g1.Reset
    Set g1 = Nothing
    Erase comp_idx
    
    'Rename cluster so largest cluster is #1, second largest is #2 etc.
    If n_cluster = 1 Then Exit Function
    ReDim iArr(1 To n_cluster)
    For i = 1 To pSize
        iArr(cluster_idx(i)) = iArr(cluster_idx(i)) + 1
    Next i
    Call modMath.Sort_Quick_A(iArr, 1, n_cluster, sort_idx)
    iArr = cluster_idx
    For i = 1 To pSize
        For j = 1 To n_cluster
            If iArr(i) = sort_idx(j) Then
                cluster_idx(i) = n_cluster - j + 1
                Exit For
            End If
        Next j
    Next i
    
End Function


'Filter graph by edge strength
'"Software component capture usign graph clustering", Yves Chiricota
Sub Filter_EdgeStrength(g1 As cGraph_Undir, _
        Optional threshold As Double = -1, Optional usethreshold As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_cluster As Long
Dim node_list() As Long, comp_idx() As Long, sort_idx() As Long, x() As Double, x_orig() As Double
Dim tmp_x As Double, tmp_y As Double, quality_max As Double, x_threshold As Double
Dim g2 As cGraph_Undir

    node_list = pNodeList.GetList
    Call Calc_Key2Idx
    
    'Calculate edge strength of each edge
    Call Calc_EdgeList
    ReDim x(1 To UBound(pEdgeList, 1))
    For k = 1 To UBound(pEdgeList, 1)
        x(k) = Edge_Strength_uv(node_list(pEdgeList(k, 1)), node_list(pEdgeList(k, 2)))
    Next k
    
    If usethreshold = True Then
        'Remove edges weaker than user supplied threshold
        x_threshold = threshold
    Else
        'Remove weak edges until maximum cluster quality is achieved.
        x_orig = x
        Call modMath.Sort_Quick_A(x, 1, UBound(x), sort_idx, 1)
        Set g2 = New cGraph_Undir
        Call g2.Copy_from(Me)
        quality_max = -Exp(70)
        k = 1
        Do
            tmp_x = x(k)
            With g2
                Do
                    If x(k) <= tmp_x Then
                        Call .EdgeDelete(pEdgeList(sort_idx(k), 1), pEdgeList(sort_idx(k), 2))
                    Else
                        Exit Do
                    End If
                    k = k + 1
                Loop While k < pn_edge
                n_cluster = .component_count(comp_idx)
            End With
            tmp_y = Cluster_Quality_MQ(comp_idx, n_cluster)
            If tmp_y > quality_max Then
                quality_max = tmp_y
                x_threshold = tmp_x
            End If
        Loop While g2.n_edge > 0
        Call g2.Reset
        Set g2 = Nothing
        Erase comp_idx
    End If

    'Remove edges
    Set g1 = New cGraph_Undir
    With g1
        Call .Copy_from(Me)
        For k = 1 To UBound(pEdgeList, 1)
            If x_orig(k) <= x_threshold Then
                Call .EdgeDelete(node_list(pEdgeList(k, 1)), node_list(pEdgeList(k, 2)))
            End If
        Next k
    End With
End Sub


Private Function Edge_Strength_uv(u As Long, v As Long) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long, w As Long
Dim n_u As Long, n_v As Long, n_w As Long, u_deg As Long, v_deg As Long
Dim count As Long, tmp_x As Double
Dim uArr() As Long, vArr() As Long, wArr() As Long, iArr() As Long, jArr() As Long
Dim isFound As Boolean

    'neighbors of u & v
    iArr = pAdjLists(pKey2Idx(u)).GetList: u_deg = UBound(iArr)
    jArr = pAdjLists(pKey2Idx(v)).GetList: v_deg = UBound(jArr)
    
    'Partition neighbors to common and non-common sets
    n_u = 0: n_v = 0: n_w = 0
    ReDim uArr(1 To u_deg)
    ReDim vArr(1 To v_deg)
    ReDim wArr(1 To u_deg)
    For i = 1 To u_deg
        m = iArr(i)
        If m <> v Then
            isFound = False
            For j = 1 To v_deg
                If jArr(j) = m Then
                    n_w = n_w + 1: wArr(n_w) = m
                    isFound = True
                    Exit For
                End If
            Next j
            If isFound = False Then
                n_u = n_u + 1: uArr(n_u) = m
            End If
        End If
    Next i
    For i = 1 To v_deg
        m = jArr(i)
        If m <> u Then
            isFound = False
            For j = 1 To n_w
                If wArr(j) = m Then
                    isFound = True
                    Exit For
                End If
            Next j
            If isFound = False Then
                n_v = n_v + 1: vArr(n_v) = m
            End If
        End If
    Next i
    If n_u > 0 Then ReDim Preserve uArr(1 To n_u)
    If n_v > 0 Then ReDim Preserve vArr(1 To n_v)
    If n_w > 0 Then ReDim Preserve wArr(1 To n_w)
    
    
    Edge_Strength_uv = 0
    If (n_u * n_v) > 0 Then
        count = 0
        For i = 1 To n_u
            m = uArr(i)
            For j = 1 To n_v
                If EdgeExists(m, vArr(j)) = True Then count = count + 1
            Next j
        Next i
        Edge_Strength_uv = Edge_Strength_uv + count * 1# / (n_u * n_v)
    End If
    
    If n_w > 0 Then
        If n_u > 0 Then
            count = 0
            For i = 1 To n_w
                m = wArr(i)
                For j = 1 To n_u
                    If EdgeExists(m, uArr(j)) = True Then count = count + 1
                Next j
            Next i
            Edge_Strength_uv = Edge_Strength_uv + count * 1# / (n_u * n_w)
        End If
        If n_v > 0 Then
            count = 0
            For i = 1 To n_w
                m = wArr(i)
                For j = 1 To n_v
                    If EdgeExists(m, vArr(j)) = True Then count = count + 1
                Next j
            Next i
            Edge_Strength_uv = Edge_Strength_uv + count * 1# / (n_v * n_w)
        End If
    End If
    
    If n_w > 1 Then
        count = 0
        For i = 1 To n_w - 1
            m = wArr(i)
            For j = i + 1 To n_w
                If EdgeExists(m, wArr(j)) = True Then count = count + 1
            Next j
        Next i
        Edge_Strength_uv = Edge_Strength_uv + count * 2# / (n_w * (n_w - 1))
    End If
    
    Edge_Strength_uv = Edge_Strength_uv + n_w / (n_w + n_u + n_v)

End Function


'Evaluate quality of a clustering by counting number of intra-cluster edges
'relative to inter-cluster edges.
'"Software component capture using graph clustering", Yves Chiricota
Function Cluster_Quality_MQ(cluster_idx() As Long, n_cluster As Long) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim ii As Long, jj As Long, u As Long, v As Long
Dim tmp_x As Double, tmp_y As Double
Dim node_list() As Long, cluster_size() As Long, cluster_edge() As Double
    If n_cluster = 1 Then
        Cluster_Quality_MQ = pn_edge * 2# / (pSize * (pSize - 1))
        Exit Function
    End If
    node_list = pNodeList.GetList
    ReDim cluster_size(1 To n_cluster)
    ReDim cluster_edge(1 To n_cluster, 1 To n_cluster)
    For i = 1 To pSize
        cluster_size(cluster_idx(i)) = cluster_size(cluster_idx(i)) + 1
    Next i
    For i = 1 To pSize - 1
        u = node_list(i): ii = cluster_idx(i)
        For j = i + 1 To pSize
            v = node_list(j): jj = cluster_idx(j)
            If EdgeExists(u, v) = True Then
                cluster_edge(ii, jj) = cluster_edge(ii, jj) + 1
                If ii <> jj Then cluster_edge(jj, ii) = cluster_edge(jj, ii) + 1
            End If
        Next j
    Next i
    tmp_x = 0: tmp_y = 0
    For i = 1 To n_cluster
        If cluster_size(i) > 1 Then
            cluster_edge(i, i) = cluster_edge(i, i) * 2# / (cluster_size(i) * (cluster_size(i) - 1))
        End If
        tmp_x = tmp_x + cluster_edge(i, i)
    Next i
    For i = 1 To n_cluster - 1
        For j = i + 1 To n_cluster
            cluster_edge(i, j) = cluster_edge(i, j) / (cluster_size(i) * cluster_size(j))
            tmp_y = tmp_y + cluster_edge(i, j)
        Next j
    Next i
    Cluster_Quality_MQ = tmp_x / n_cluster - tmp_y * 2# / (n_cluster * (n_cluster - 1))
    Erase cluster_size, cluster_edge, node_list
End Function


Private Sub MMULT_Cube_Diag(A() As Long, x() As Double)
Dim i As Long, j As Long, k As Long, n As Long
Dim tmp_x As Long, tmp_y As Long, tmp_z As Long
    n = UBound(A, 1)
    ReDim x(1 To n)
    For i = 1 To n
        If i Mod 100 = 0 Then DoEvents
        tmp_x = 0: tmp_z = 0
        For j = 1 To n
            tmp_y = A(j, i)
            If tmp_y > 0 Then
                For k = 1 To j - 1
                    tmp_x = tmp_x + A(k, j) * A(k, i) * tmp_y
                    tmp_z = tmp_z + A(k, i) * tmp_y
                Next k
                For k = j + 1 To n
                    tmp_x = tmp_x + A(k, j) * A(k, i) * tmp_y
                    tmp_z = tmp_z + A(k, i) * tmp_y
                Next k
            End If
        Next j
        If tmp_z > 0 Then x(i) = tmp_x * 1# / tmp_z
    Next i
End Sub





Private Sub Reverse_vec(iArr() As Long)
Dim i As Long, k As Long, n As Long, m As Long
    n = UBound(iArr)
    m = n \ 2
    For i = 1 To m
        k = iArr(i)
        iArr(i) = iArr(n - i + 1)
        iArr(n - i + 1) = k
    Next i
End Sub

Private Sub MinHeap_Add_Tuple(q As cHeap, x As Double, idx As Long, u As Long, v As Long, iArr() As Long)
    idx = idx + 1
    Call q.Add(x, idx)
    ReDim Preserve iArr(1 To 2, 0 To idx)
    iArr(1, idx) = u
    iArr(2, idx) = v
End Sub

Private Sub MinHeap_Pop_Tuple(q As cHeap, x As Double, idx As Long, u As Long, v As Long, iArr() As Long)
    Call q.Pop_Min(x, idx)
    u = iArr(1, idx)
    v = iArr(2, idx)
End Sub


'Check if graph is bipartite and return integer vector node_color(1:N) which lable each node as 0 or 1.
Function isBipartite(node_color() As Long, Optional src_node_key As Long = 1) As Boolean
Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long, w As Long, u_parent As Long
Dim neigh_list() As Long
Dim n_head As Long, n_tail As Long, iQueue() As Long
    isBipartite = True
    Call Calc_Key2Idx
    ReDim node_color(1 To pSize)
    For i = 1 To pSize
        node_color(i) = -1
    Next i
    ReDim iQueue(1 To 3 * pSize)
    u = pKey2Idx(src_node_key)
    node_color(u) = 0                       'Source node is colored as 0
    n_head = 1: n_tail = 1: iQueue(1) = u   'Add source node to queue
    Do While n_head <= n_tail
        u = iQueue(n_head): n_head = n_head + 1 'pop u from queue
        neigh_list = pAdjLists(u).GetList       'List of u's neighbors
        'Color each of u's neighbors
        For i = 1 To UBound(neigh_list)
            w = pKey2Idx(neigh_list(i))
            'node w was already assigned same color as u, graph is not bipartite
            If node_color(w) = node_color(u) Then
                Debug.Print "isBipartite: graph is not bipartite. Cycle: " & w & " - " & u & " - " & u_parent
                isBipartite = False
                Exit Function
            End If
            If node_color(w) = -1 Then
                node_color(w) = 1 - node_color(u)
                n_tail = n_tail + 1: iQueue(n_tail) = w
            End If
        Next i
        u_parent = u
    Loop
    Erase iQueue, neigh_list
End Function
