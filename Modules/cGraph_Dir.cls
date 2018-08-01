VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGraph_Dir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSize As Long, pn_edge As Long
Private pnode_pos() As Double
Private pEdgeList() As Long, pEdgeWgt() As Double
Private pnode_eigen() As Double, pnode_pagerank() As Double, pnode_katz() As Double
Private pnode_hub() As Double, pnode_authority() As Double
Private pnode_reach() As Long, pnode_reach_in() As Long
Private pnode_deg() As Long, pnode_deg_out() As Long, pnode_deg_in() As Long
Private pnode_core() As Long, pnode_core_out() As Long, pnode_core_in() As Long
Private pnode_deg_wgt() As Double, pnode_deg_out_wgt() As Double, pnode_deg_in_wgt() As Double
Private pnode_closeness_out() As Double, pnode_closeness_in() As Double, pnode_betweenness() As Double
Private pKey2Idx() As Long
Private pNodeList As cAdjList
Private pAdjLists As Collection, pAdjLists_in As Collection


'========================================================
'Access Properties
'========================================================
Public Property Get Size() As Long
    Size = pSize
End Property

Public Property Get n_edge() As Long
    n_edge = pn_edge
End Property

Public Property Get node_deg(Optional strType As String = "OUT") As Long()
    If VBA.UCase$(strType) = "OUT" Then
        node_deg = pnode_deg_out
    ElseIf VBA.UCase$(strType) = "IN" Then
        node_deg = pnode_deg_in
    ElseIf VBA.UCase$(strType) = "TOTAL" Then
        node_deg = pnode_deg
    End If
End Property

Public Property Get node_core(Optional strType As String = "OUT") As Long()
    If VBA.UCase$(strType) = "OUT" Then
        node_core = pnode_core_out
    ElseIf VBA.UCase$(strType) = "IN" Then
        node_core = pnode_core_in
    ElseIf VBA.UCase$(strType) = "TOTAL" Then
        node_core = pnode_core
    End If
End Property

Public Property Get node_deg_wgt(Optional strType As String = "OUT") As Double()
    If VBA.UCase$(strType) = "OUT" Then
        node_deg_wgt = pnode_deg_out_wgt
    ElseIf VBA.UCase$(strType) = "IN" Then
        node_deg_wgt = pnode_deg_in_wgt
    ElseIf VBA.UCase$(strType) = "TOTAL" Then
        node_deg_wgt = pnode_deg_wgt
    End If
End Property

Public Property Get node_reach(Optional strType As String = "OUT") As Long()
    If VBA.UCase$(strType) = "OUT" Then
        node_reach = pnode_reach
    ElseIf VBA.UCase$(strType) = "IN" Then
        node_reach = pnode_reach_in
    End If
End Property

Public Property Get node_eigen() As Double()
    node_eigen = pnode_eigen
End Property

Public Property Get node_katz() As Double()
    node_katz = pnode_katz
End Property

Public Property Get node_hub() As Double()
    node_hub = pnode_hub
End Property

Public Property Get node_authority() As Double()
    node_authority = pnode_authority
End Property

Public Property Get node_pagerank() As Double()
    node_pagerank = pnode_pagerank
End Property

Public Property Get node_closeness(Optional strType As String = "OUT") As Double()
    If strType = "OUT" Then
        node_closeness = pnode_closeness_out
    ElseIf strType = "IN" Then
        node_closeness = pnode_closeness_in
    End If
End Property

Public Property Get node_betweenness() As Double()
    node_betweenness = pnode_betweenness
End Property


Public Property Let node_pos(x() As Double)
    pnode_pos = x
End Property

Public Property Get node_pos() As Double()
    node_pos = pnode_pos
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
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            iArr(k, 1) = jArr(i)
            iArr(k, 2) = jArr(j)
        Next k
        EdgeList = iArr
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

Sub SubGraph(g1 As cGraph_Dir, node_list() As Long, Optional isKey As Boolean = True, _
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
                neigh_list = pAdjLists_in(node_list(i)).GetList
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
                neigh_list = pAdjLists_in(node_list(i)).GetList
                For j = 1 To UBound(neigh_list)
                    isIncl(pKey2Idx(neigh_list(j))) = 1
                Next j
            End If
        Next i
    End If
    Set g1 = New cGraph_Dir
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


Function node_neighbors(v As Long, Optional strType As String = "OUT", _
    Optional wgts As Variant, Optional output_wgt As Boolean = False) As Long()
Dim i As Long, j As Long, k As Long, n As Long, v_idx As Long
Dim iArr() As Long, jArr() As Long, isDuplicate As Boolean
Dim xArr() As Double, yArr() As Double
    v_idx = Key2Idx(v)
    If strType = "OUT" Then
        node_neighbors = pAdjLists(v_idx).GetList
        If output_wgt = True Then wgts = pAdjLists(v_idx).GetListWgt
    ElseIf strType = "IN" Then
        node_neighbors = pAdjLists_in(v_idx).GetList
        If output_wgt = True Then wgts = pAdjLists_in(v_idx).GetListWgt
    ElseIf strType = "TOTAL" Then
        If output_wgt = False Then
            iArr = pAdjLists(v_idx).GetList
            jArr = pAdjLists_in(v_idx).GetList
        Else
            Call pAdjLists(v_idx).GetListnWgt(iArr, xArr)
            Call pAdjLists_in(v_idx).GetListnWgt(jArr, yArr)
        End If
        k = UBound(iArr, 1)
        If k > 0 Then
            For j = 1 To UBound(jArr, 1)
                isDuplicate = False
                For i = 1 To k
                    If jArr(j) = iArr(i) Then
                        isDuplicate = True
                        xArr(i) = xArr(i) + yArr(j)
                        Exit For
                    End If
                Next i
                If isDuplicate = False Then
                    n = UBound(iArr, 1) + 1
                    ReDim Preserve iArr(LBound(iArr, 1) To n)
                    iArr(n) = jArr(j)
                    If output_wgt = True Then
                        ReDim Preserve xArr(LBound(xArr, 1) To n)
                        xArr(n) = yArr(j)
                    End If
                End If
            Next j
        Else
            iArr = jArr
            xArr = yArr
        End If
        node_neighbors = iArr
        If output_wgt = True Then wgts = xArr
'        If UBound(iArr, 1) > 0 Then
'            ReDim jArr(1 To UBound(iArr, 1))
'            For i = 1 To UBound(iArr)
'                jArr(i) = iArr(i)
'            Next i
'        Else
'            ReDim jArr(0 To 0)
'        End If
'        node_neighbors = jArr
    End If
End Function

Function node_neighbors_wgt(v As Long, Optional strType As String = "OUT") As Double()
    If strType = "OUT" Then
        node_neighbors_wgt = pAdjLists(Key2Idx(v)).GetListWgt
    ElseIf strType = "IN" Then
        node_neighbors_wgt = pAdjLists_in(Key2Idx(v)).GetListWgt
    End If
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
    Set pAdjLists_in = New Collection
    For i = 1 To pSize
        Set v = New cAdjList
        pAdjLists.Add v
        Set v = New cAdjList
        pAdjLists_in.Add v
    Next i
    Set pNodeList = New cAdjList
    If custom_key = False Then
        If isWgt = True Then
            For i = 1 To pSize
                j = pNodeList.Append(i, VBA.CDbl(node_wgt(i)))
            Next i
        Else
            For i = 1 To pSize
                j = pNodeList.Append(i, 1)
            Next i
        End If
    Else
        If isWgt = True Then
            For i = 1 To pSize
                j = pNodeList.Append(Int(node_key(i)), VBA.CDbl(node_wgt(i)))
            Next i
        Else
            For i = 1 To pSize
                j = pNodeList.Append(Int(node_key(i)), 1)
            Next i
        End If
    End If
End Sub

Sub Reset()
Dim i As Long, n As Long
    Erase pKey2Idx
    Erase pnode_deg, pnode_deg_out, pnode_deg_in
    Erase pnode_core, pnode_core_out, pnode_core_in
    Erase pnode_deg_wgt, pnode_deg_out_wgt, pnode_deg_in_wgt
    Erase pnode_eigen, pnode_pagerank, pnode_reach, pnode_reach_in
    Erase pEdgeList, pEdgeWgt, pnode_pos
    Erase pnode_betweenness, pnode_closeness_in, pnode_closeness_out
    n = pAdjLists.count
    For i = n To 1 Step -1
        pAdjLists.Remove (i)
        pAdjLists_in.Remove (i)
    Next i
    Set pAdjLists = Nothing
    Set pAdjLists_in = Nothing
    Call pNodeList.Reset
    Set pNodeList = Nothing
    pSize = 0
    pn_edge = 0
End Sub


'Make current graph into an identical copy of g1
Sub Copy_from(g1 As cGraph_Dir)
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
        iArr = g1.node_neighbors(idx(i), "OUT", xArr, True)
        j = pAdjLists(i).ReplaceList(iArr, xArr, True)
        iArr = g1.node_neighbors(idx(i), "IN", xArr, True)
        j = pAdjLists_in(i).ReplaceList(iArr, xArr, True)
    Next i
    Erase idx, iArr, xArr
    Application.StatusBar = False
End Sub


Sub AdjM_Add(AdjMatrix As Variant, Optional isWgt As Boolean = True, Optional isTranspose As Boolean = False, _
                Optional ignore_selfloop As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long
Dim xArr() As Double, iArr() As Long
Dim yArr() As Double, jArr() As Long
Dim node_keys() As Long
    If pSize <> UBound(AdjMatrix, 1) Or pSize <> UBound(AdjMatrix, 2) Then
        Debug.Print "cGraph_Dir: AdjM_Add: size of matrix inconsistent with number of nodes."
        Exit Sub
    End If
    node_keys = NodeList
    'Clear all existing edges
    pn_edge = 0
    For i = 1 To pSize
        Call pAdjLists(i).Init
        Call pAdjLists_in(i).Init
    Next i
    For i = 1 To pSize
        u = node_keys(i)
        If i Mod 100 = 0 Then
            DoEvents: Application.StatusBar = "Importing Adjacency matrix: " & i & "/" & pSize
        End If
        'Add outbound edges
        k = 0: ReDim xArr(1 To pSize): ReDim iArr(1 To pSize)
        For j = 1 To pSize
            If j <> i And AdjMatrix(i, j) <> 0 Then
                k = k + 1: iArr(k) = node_keys(j): xArr(k) = AdjMatrix(i, j)
            End If
        Next j
        If k > 0 Then
            ReDim Preserve iArr(1 To k)
            If isWgt = True Then
                ReDim Preserve xArr(1 To k)
                If isTranspose = False Then
                    m = pAdjLists(i).ReplaceList(iArr, xArr, True)
                Else
                    m = pAdjLists_in(i).ReplaceList(iArr, xArr, True)
                End If
            Else
                If isTranspose = False Then
                    m = pAdjLists(i).ReplaceList(iArr)
                Else
                    m = pAdjLists_in(i).ReplaceList(iArr)
                End If
            End If
            pn_edge = pn_edge + k
        End If
        
        'Add inbound edges
        k = 0: ReDim xArr(1 To pSize): ReDim iArr(1 To pSize)
        For j = 1 To pSize
            If j <> i And AdjMatrix(j, i) <> 0 Then
                k = k + 1: iArr(k) = node_keys(j): xArr(k) = AdjMatrix(j, i)
            End If
        Next j
        If k > 0 Then
            ReDim Preserve iArr(1 To k)
            If isWgt = True Then
                ReDim Preserve xArr(1 To k)
                If isTranspose = False Then
                    m = pAdjLists_in(i).ReplaceList(iArr, xArr, True)
                Else
                    m = pAdjLists(i).ReplaceList(iArr, xArr, True)
                End If
            Else
                If isTranspose = False Then
                    m = pAdjLists_in(i).ReplaceList(iArr)
                Else
                    m = pAdjLists(i).ReplaceList(iArr)
                End If
            End If
        End If
    Next i
    Erase xArr, iArr
    
    'Add self-loop
    If ignore_selfloop = False Then
        For i = 1 To pSize
            If AdjMatrix(i, i) <> 0 Then
                u = node_keys(i)
                Call EdgeAdd(u, u, VBA.CDbl(AdjMatrix(i, i)))
            End If
        Next i
    End If
    Application.StatusBar = False
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
            For j = 1 To n_node - 1
                If i <> j Then
                    tmp_x = VBA.Rnd
                    If tmp_x < x_prob Then Call EdgeAdd(i, j)
                End If
            Next j
        Next i
    Else
        For i = 1 To n_node
            For j = 1 To n_node
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

'Add a node to graph using v as key, if node already exist no action is performed
Sub NodeAdd(v As Long, Optional Wgt As Double = 1)
Dim i As Long
Dim vList As cAdjList
    i = pNodeList.Append(v, Wgt)
    pSize = pSize + i
    If i > 0 Then 'Node is indeed new
        Set vList = New cAdjList
        pAdjLists.Add vList
        Set vList = New cAdjList
        pAdjLists_in.Add vList
    End If
End Sub

'Delete a node from graph by its key
Sub NodeDelete(v As Long, Optional adjpos As Boolean = False)
Dim i As Long, j As Long, k As Long, idx As Long
Dim iArr() As Long, jArr() As Long, xArr() As Double
    idx = Key2Idx(v)
    iArr = pAdjLists(idx).GetList
    jArr = pAdjLists_in(idx).GetList
    If UBound(iArr) > 0 Then
        For j = 1 To UBound(iArr)
            Call EdgeDelete(v, iArr(j))
        Next j
    End If
    If UBound(jArr) > 0 Then
        For j = 1 To UBound(jArr)
            Call EdgeDelete(jArr(j), v)
        Next j
    End If
    Erase iArr, jArr
    i = pNodeList.Remove(v)
    pSize = pSize - i
    If i > 0 Then 'Node is indeed removed
        pAdjLists.Remove idx
        pAdjLists_in.Remove idx
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
    Call EdgeDelete(u, v)
    Call EdgeDelete(v, u)

    If adjpos = True Then
        tmp_x = NodeWgt(u)
        tmp_y = NodeWgt(v)
        For k = 1 To UBound(pnode_pos, 2)
            pnode_pos(i, k) = (pnode_pos(i, k) * tmp_x + pnode_pos(j, k) * tmp_y) / (tmp_x + tmp_y)
        Next k
    End If
    
    'Outbound edges from v are added to those of u
    Call pAdjLists(j).GetListnWgt(iArr, xArr)
    For k = 1 To UBound(iArr)
        w = iArr(k)
        Call EdgeDelete(v, w)
        Call EdgeAdd(u, w, xArr(k), 2)
    Next k

    'Inbound edges to v are added to u
    Call pAdjLists_in(j).GetListnWgt(iArr, xArr)
    For k = 1 To UBound(iArr)
        w = iArr(k)
        Call EdgeDelete(w, v)
        Call EdgeAdd(w, u, xArr(k), 2)
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
Dim new_AdjLists As Collection, new_AdjLists_in As Collection, xArr() As Double
    'Reorder node list
    Call pNodeList.Reorder(iorder, swap_idx_tmp, True)
    'Reorder adjacency list
    Set new_AdjLists = New Collection
    Set new_AdjLists_in = New Collection
    For i = 1 To pSize
        new_AdjLists.Add pAdjLists(swap_idx_tmp(i))
        new_AdjLists_in.Add pAdjLists_in(swap_idx_tmp(i))
    Next i
    For i = pSize To 1 Step -1
        pAdjLists.Remove (i)
        pAdjLists_in.Remove (i)
    Next i
    Set pAdjLists = new_AdjLists
    Set pAdjLists_in = new_AdjLists_in
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

'Add directed edge from u to v
Sub EdgeAdd(u As Long, v As Long, Optional Wgt As Double = 1, Optional replace_wgt As Long = 0, _
                Optional check_duplicate As Boolean = True)
Dim i As Long
    If check_duplicate = True Then
        i = pAdjLists(Key2Idx(u)).Append(v, Wgt, replace_wgt)
        pn_edge = pn_edge + i
        'Also add incoming edges to v
        i = pAdjLists_in(Key2Idx(v)).Append(u, Wgt, replace_wgt)
    Else
        i = pAdjLists(Key2Idx(u)).Prepend(v, Wgt, replace_wgt, False)
        pn_edge = pn_edge + i
        'Also add incoming edges to v
        i = pAdjLists_in(Key2Idx(v)).Prepend(u, Wgt, replace_wgt, False)
    End If
End Sub

'remove directed edge from u to v
Sub EdgeDelete(u As Long, v As Long)
Dim i As Long
    i = pAdjLists(Key2Idx(u)).Remove(v)
    pn_edge = pn_edge - i
    'Also delete from v's incident list
    i = pAdjLists_in(Key2Idx(v)).Remove(u)
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
End Sub

'remove all outward edges from u
Sub EdgeDelete_Out(u As Long)
Dim uList As cAdjList
Dim iArr() As Long, i As Long
    Set uList = pAdjLists(Key2Idx(u))
    iArr = uList.GetList
    For i = 1 To UBound(iArr)
        Call EdgeDelete(u, iArr(i))
    Next i
    Call uList.Init
End Sub

'remove all incoming edges to u
Sub EdgeDelete_In(u As Long)
Dim uList As cAdjList
Dim i As Long, iArr() As Long
    Set uList = pAdjLists_in(Key2Idx(u))
    iArr = uList.GetList
    For i = 1 To UBound(iArr)
        Call EdgeDelete(iArr(i), u)
    Next i
    Call uList.Init
End Sub

Sub Calc_EdgeList()
Dim i As Long, j As Long, k As Long, n As Long
Dim iArr() As Long, xArr() As Double
    n = 0
    Call Calc_Key2Idx
    ReDim pEdgeList(1 To pn_edge, 1 To 2)
    ReDim pEdgeWgt(1 To pn_edge)
    For j = 1 To pSize
        Call pAdjLists(j).GetListnWgt(iArr, xArr)
        For k = 1 To UBound(iArr)
            n = n + 1
            pEdgeList(n, 1) = j
            pEdgeList(n, 2) = pKey2Idx(iArr(k))
            pEdgeWgt(n) = xArr(k)
        Next k
    Next j
    Erase iArr, xArr
End Sub

'Remove duplicate records from edge list, each record is then assumed to be undirected
Sub Symmetrize_EdgeList(EdgeList() As Long, EdgeWgt() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim iArr() As Long, jArr() As Long
Dim isDuplicate As Boolean
Dim xArr() As Double, yArr() As Double
    n = 0
    Call Calc_Key2Idx
    ReDim EdgeList(1 To 2, 1 To pn_edge)
    ReDim EdgeWgt(1 To pn_edge)
    For i = 1 To pSize
        ReDim jArr(1 To pSize)
        Call pAdjLists(i).GetListnWgt(iArr, xArr)
        For k = 1 To UBound(iArr)
            j = pKey2Idx(iArr(k))
            If j > i Then
                n = n + 1
                jArr(j) = 1
                EdgeList(1, n) = i
                EdgeList(2, n) = j
                EdgeWgt(n) = xArr(k)
            End If
        Next k
        Call pAdjLists_in(i).GetListnWgt(iArr, xArr)
        For k = 1 To UBound(iArr)
            j = pKey2Idx(iArr(k))
            If j > i And jArr(j) = 0 Then
                n = n + 1
                jArr(j) = 1
                EdgeList(1, n) = i
                EdgeList(2, n) = j
                EdgeWgt(n) = xArr(k)
            End If
        Next k
    Next i
    Erase iArr, xArr
    ReDim Preserve EdgeWgt(1 To n)
    ReDim Preserve EdgeList(1 To 2, 1 To n)
    iArr = EdgeList
    ReDim EdgeList(1 To n, 1 To 2)
    For k = 1 To n
        EdgeList(k, 1) = iArr(1, k)
        EdgeList(k, 2) = iArr(2, k)
    Next k
    Erase iArr
End Sub

Sub EdgeReverseAll()
Dim i As Long
Dim uList As cAdjList, vList As cAdjList
For i = 1 To pSize
    Set uList = pAdjLists(i)
    Set vList = pAdjLists_in(i)
    pAdjLists.Remove i
    pAdjLists_in.Remove i
    If i < pSize Then
        pAdjLists.Add vList, Before:=i
        pAdjLists_in.Add uList, Before:=i
    Else
        pAdjLists.Add vList
        pAdjLists_in.Add uList
    End If
Next i
End Sub

Sub EdgeReverse(u As Long, v As Long)
Dim tmp_x As Double
    If EdgeExists(u, v) = False Then Exit Sub
    tmp_x = EdgeWgt(u, v)
    Call EdgeDelete(u, v)
    Call EdgeAdd(v, u, tmp_x, 1)
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

'Convert to undirected graph and calculate layout
Sub Calc_Layout(Optional strMethod As String = "FORCE_MULTI", _
            Optional n_dimension As Long = 2, Optional isWgt As Boolean = True, _
            Optional tol As Double = 0.00000001, Optional iter_max As Long = 5000, _
            Optional isDegNorm As Boolean = True, Optional WgtType As String = "NONE", _
            Optional c_spring As Double = 1, Optional c_repulsive As Double = 1, Optional c_power As Double = 2)
Dim i As Long
Dim g1 As cGraph_Undir
    Set g1 = New cGraph_Undir
    With g1
        Call .Convert_from_Directed(Me)
        Call .Calc_Layout(strMethod, n_dimension, isWgt, tol, iter_max, _
                isDegNorm, WgtType, c_spring, c_repulsive, c_power)
        pnode_pos = g1.node_pos
        Call .Reset
    End With
    Set g1 = Nothing
End Sub

Function Print_Edges() As Variant
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long, m As Long
Dim tmp_x As Double, tmp_y As Double, tmp_vec() As Double
Dim iArr() As Long
Dim uList As cAdjList
Dim vArr As Variant
    n = 0
    n_dimension = UBound(pnode_pos, 2)
    Call Calc_Key2Idx
    ReDim vArr(1 To pn_edge * 3 - 1, 1 To n_dimension)
    ReDim tmp_vec(1 To n_dimension)
    For i = 1 To pSize
        Set uList = pAdjLists(i)
        If uList.Size > 0 Then
            For m = 1 To n_dimension
                tmp_vec(m) = pnode_pos(i, m)
            Next m
            iArr = uList.GetList
            For k = 1 To UBound(iArr)
                j = pKey2Idx(iArr(k))
                n = n + 3
                For m = 1 To n_dimension
                    vArr(n - 2, m) = tmp_vec(m)
                    vArr(n - 1, m) = pnode_pos(j, m)
                Next m
            Next k
        End If
    Next i
    Print_Edges = vArr
    Erase iArr, vArr, tmp_vec
    Set uList = Nothing
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
            If DFS_Order(v) = 0 Then
                DFS_parent(v) = u
                Call DFS_Recursive(v, DFS_Order, n_visited, DFS_parent, vPath)
            End If
        Next k
    End If
    Set uList = Nothing
    Erase iArr
End Sub

Function DFS_2Ways(u As Long, Optional plot_Path As Variant, Optional show_path As Boolean = False, _
            Optional check_full As Boolean = False) As Long()
Dim i As Long, j As Long, k As Long, v As Long, m As Long, n As Long, v_prev As Long, v2 As Long
Dim u_idx As Long
Dim tmp_x As Double
Dim DFS_Order() As Long, DFS_parent() As Long, vPath() As Long, n_visited As Long
    
    Call Calc_Key2Idx
    u_idx = pKey2Idx(u)
    
    If u_idx = 0 Then  'no path can be shown
        Debug.Print "DFS_2Ways Failed: " & u & " does not exist in graph."
        ReDim DFS_Order(1 To pSize)
        DFS_2Ways = DFS_Order
        If show_path = True Then ReDim plot_Path(1 To 1, 1 To 2)
        Exit Function
    End If
    
    ReDim DFS_Order(1 To pSize)
    ReDim DFS_parent(1 To pSize)
    ReDim vPath(1 To pSize)
    n_visited = 0
    DFS_parent(u_idx) = -1
    Call DFS_Recursive_2Ways(u_idx, DFS_Order, n_visited, DFS_parent, vPath)
    If check_full = True Then
        For i = 1 To pSize
            If DFS_Order(i) = 0 Then
                DFS_parent(i) = -1
                Call DFS_Recursive_2Ways(i, DFS_Order, n_visited, DFS_parent, vPath)
            End If
        Next i
    End If
    DFS_2Ways = DFS_Order

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

Private Sub DFS_Recursive_2Ways(u As Long, _
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
            If DFS_Order(v) = 0 Then
                DFS_parent(v) = u
                Call DFS_Recursive_2Ways(v, DFS_Order, n_visited, DFS_parent, vPath)
            End If
        Next k
    End If
    Set uList = Nothing
    Set uList = pAdjLists_in(u)
    If uList.Size > 0 Then
        iArr = uList.GetList
        For k = 1 To UBound(iArr)
            v = pKey2Idx(iArr(k))
            If DFS_Order(v) = 0 Then
                DFS_parent(v) = u
                Call DFS_Recursive_2Ways(v, DFS_Order, n_visited, DFS_parent, vPath)
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

'Return number of strongly connected components, and an
'integer array giving component membership of each node.
Function StrongConnect(comp_idx() As Long) As Long
Dim i As Long, j As Long, u As Long, v As Long
Dim DFS_Order() As Long, vStack() As Long, n_stack As Long, n_visited As Long
Dim LowPt() As Long, onStack() As Boolean, sort_idx() As Long, n_component As Long
Dim iArr() As Long
    Call Calc_Key2Idx
    ReDim DFS_Order(1 To pSize): n_visited = 0
    ReDim vStack(1 To 3 * pSize): n_stack = 0
    ReDim onStack(1 To pSize)
    ReDim LowPt(1 To pSize)
    ReDim comp_idx(1 To pSize)
    n_component = 0
    For i = 1 To pSize
        If DFS_Order(i) = 0 Then
            Call StrongConnect_Step(i, DFS_Order, n_visited, LowPt, vStack, n_stack, onStack, comp_idx, n_component)
        End If
    Next i
    StrongConnect = n_component
    
    'Label largest component as #1
    ReDim comp_size(1 To n_component)
    For i = 1 To pSize
        j = comp_idx(i)
        comp_size(j) = comp_size(j) + 1
    Next i
    Call modMath.Sort_Quick_A(comp_size, 1, n_component, sort_idx, 1)
    iArr = comp_idx
    For i = 1 To n_component
        j = sort_idx(i)
        For u = 1 To pSize
            If iArr(u) = j Then comp_idx(u) = n_component - i + 1
        Next u
    Next i
    Erase iArr
End Function

Private Sub StrongConnect_Step(v As Long, DFS_Order() As Long, n_visited As Long, LowPt() As Long, _
vStack() As Long, n_stack As Long, onStack() As Boolean, comp_idx() As Long, n_component As Long)
Dim k As Long, u As Long
Dim iArr() As Long
Dim uList As cAdjList
    n_visited = n_visited + 1
    DFS_Order(v) = n_visited
    LowPt(v) = n_visited
    n_stack = n_stack + 1
    vStack(n_stack) = v
    onStack(v) = True
    Set uList = pAdjLists(v)
    If uList.Size > 0 Then
         iArr = uList.GetList
         For k = 1 To UBound(iArr)
            u = pKey2Idx(iArr(k))
            If DFS_Order(u) = 0 Then
                Call StrongConnect_Step(u, DFS_Order, n_visited, LowPt, _
                    vStack, n_stack, onStack, comp_idx, n_component)
                If LowPt(u) < LowPt(v) And LowPt(u) > 0 Then LowPt(v) = LowPt(u)
            ElseIf onStack(u) = True Then
                If DFS_Order(u) < LowPt(v) Then LowPt(v) = DFS_Order(u)
            End If
         Next k
    End If
    If LowPt(v) = DFS_Order(v) Then
        n_component = n_component + 1
        Do While n_stack > 0
            u = vStack(n_stack): n_stack = n_stack - 1
            onStack(u) = False
            comp_idx(u) = n_component
            If u = v Then Exit Do
        Loop
    End If
End Sub


'Return number of weakly connected components, and an
'integer array giving component membership of each node.
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
            iArr = DFS_2Ways(node_list(i), , , False)
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


'Make a new graph of each weakly connected component and put into a collection
Sub Disconnected_Components(g_components As Collection, comp_idx() As Long, n_component As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long
Dim g1 As cGraph_Dir
Dim node_keys() As Long, node_wgts() As Double
Dim iArr() As Long, jArr() As Long, xArr() As Double, yArr() As Double, comp_size() As Long

    n_component = component_count(comp_idx)
    ReDim comp_size(1 To n_component)
    For i = 1 To pSize
        j = comp_idx(i)
        comp_size(j) = comp_size(j) + 1
    Next i
    
    Set g_components = New Collection
    If n_component = 1 Then
        g_components.Add Me
        Exit Sub
    End If
    
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
        
        Set g1 = New cGraph_Dir
        With g1
            Call .Init(comp_size(k), xArr, True, iArr, True)
            For m = 1 To comp_size(k)
                i = iArr(m)
                jArr = node_neighbors(i, "OUT", yArr, True)
                For n = 1 To UBound(jArr)
                    Call .EdgeAdd(i, jArr(n), yArr(n), 0)
                Next n
                jArr = node_neighbors(i, "IN", yArr, True)
                For n = 1 To UBound(jArr)
                    Call .EdgeAdd(jArr(n), i, yArr(n), 0)
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


'========================================================
'Calculate network attributes
'========================================================

Sub Calc_Node_Deg()
Dim i As Long
    ReDim pnode_deg(1 To pSize)
    Call Calc_Node_Deg_In
    Call Calc_Node_Deg_Out
    For i = 1 To pSize
        pnode_deg(i) = pnode_deg_in(i) + pnode_deg_out(i)
    Next i
End Sub

Sub Calc_Node_Deg_Out()
Dim i As Long
    ReDim pnode_deg_out(1 To pSize)
    For i = 1 To pSize
        pnode_deg_out(i) = pAdjLists(i).Size
    Next i
End Sub

Sub Calc_Node_Deg_In()
Dim i As Long
    ReDim pnode_deg_in(1 To pSize)
    For i = 1 To pSize
        pnode_deg_in(i) = pAdjLists_in(i).Size
    Next i
End Sub

Sub Calc_Node_Deg_Out_Wgt()
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double, tmp_y As Double, xArr() As Double
    ReDim pnode_deg_out_wgt(1 To pSize)
    For i = 1 To pSize
        tmp_x = 0
        xArr = pAdjLists(i).GetListWgt
        For k = 1 To UBound(xArr)
            tmp_x = tmp_x + xArr(k)
        Next k
        pnode_deg_out_wgt(i) = tmp_x
    Next i
End Sub

Sub Calc_Node_Deg_In_Wgt()
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double, tmp_y As Double, xArr() As Double
    ReDim pnode_deg_in_wgt(1 To pSize)
    For i = 1 To pSize
        tmp_x = 0
        xArr = pAdjLists_in(i).GetListWgt
        For k = 1 To UBound(xArr)
            tmp_x = tmp_x + xArr(k)
        Next k
        pnode_deg_in_wgt(i) = tmp_x
    Next i
End Sub

Sub Calc_Node_Deg_Wgt()
Dim i As Long
    ReDim pnode_deg_wgt(1 To pSize)
    Call Calc_Node_Deg_In_Wgt
    Call Calc_Node_Deg_Out_Wgt
    For i = 1 To pSize
        pnode_deg_wgt(i) = pnode_deg_in_wgt(i) + pnode_deg_out_wgt(i)
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

Sub Calc_Node_Reach_in()
Dim i As Long, n_visited As Long, j As Long
Dim DFS_Order() As Long, DFS_parent() As Long, vPath() As Long
    Call Calc_Key2Idx
    Call EdgeReverseAll
    ReDim pnode_reach_in(1 To pSize)
    For i = 1 To pSize
        ReDim DFS_Order(1 To pSize)
        ReDim DFS_parent(1 To pSize)
        ReDim vPath(1 To pSize)
        n_visited = 0
        DFS_parent(i) = -1
        Call DFS_Recursive(i, DFS_Order, n_visited, DFS_parent, vPath)
        pnode_reach_in(i) = n_visited - 1
    Next i
    Erase DFS_Order, DFS_parent, vPath
    Call EdgeReverseAll
End Sub


Sub Calc_Node_Eigen(Optional isWgt As Boolean = False, Optional strLR As String = "RIGHT", Optional eigenval As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim iterate As Long, iter_max As Long
Dim tmp_x As Double, tmp_y As Double, tol As Double
Dim x() As Double, x_new() As Double, iArr() As Long, jArr() As Long, kArr() As Long
Dim xArr() As Double, yArr() As Double

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
    k = 0
    Call Calc_Key2Idx
    ReDim kArr(1 To pn_edge, 1 To 2)
    ReDim yArr(1 To pn_edge)
    For i = 1 To pSize
        Call pAdjLists(i).GetListnWgt(jArr, xArr)
        For j = 1 To UBound(jArr)
            k = k + 1
            kArr(k, 1) = i
            kArr(k, 2) = pKey2Idx(jArr(j))
            If isWgt = True Then yArr(k) = xArr(j)
        Next j
    Next i

    If isWgt = False Then
        For k = 1 To pn_edge
            yArr(k) = 1
        Next k
    End If

    For iterate = 1 To iter_max

        ReDim x_new(1 To pSize)
        If strLR = "RIGHT" Then
            For k = 1 To pn_edge
                i = kArr(k, 1)
                x_new(i) = x_new(i) + x(kArr(k, 2)) * yArr(k)
            Next k
        ElseIf strLR = "LEFT" Then
            For k = 1 To pn_edge
                i = kArr(k, 2)
                x_new(i) = x_new(i) + x(kArr(k, 1)) * yArr(k)
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
    If iterate >= iter_max Then Debug.Print "Calc_Node_Eigen: Failed to converge."
    pnode_eigen = x
    eigenval = tmp_y
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


Sub Calc_Node_PageRank(Optional damping As Double = 0.85, _
        Optional isWgt As Boolean = False, Optional strType As String = "IN")
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim iterate As Long, iter_max As Long
Dim tmp_x As Double, tmp_y As Double, tol As Double, x_sum As Double, tmp_const As Double
Dim x() As Double, x_new() As Double, iArr() As Long, jArr() As Long, kArr() As Long
Dim xArr() As Double, yArr() As Double
Dim empty_list() As Long
    Randomize
    tol = 0.000000000001
    iter_max = 1000
    tmp_const = (1 - damping) / pSize
    ReDim x(1 To pSize)
    For i = 1 To pSize
        x(i) = 1# / pSize
    Next i
    
    If strType = "OUT" Then Call EdgeReverseAll
    
    'Create transition probability matrix
    Call Calc_Key2Idx
    k = 0
    ReDim empty_list(0 To 0)
    ReDim kArr(1 To pn_edge, 1 To 2)
    ReDim yArr(1 To pn_edge)
    If isWgt = True Then Call Calc_Node_Deg_Out_Wgt
    For i = 1 To pSize
        m = pAdjLists(i).Size
        If m > 0 Then
            tmp_x = 1# / m
            Call pAdjLists(i).GetListnWgt(jArr, xArr)
            For j = 1 To UBound(jArr)
                k = k + 1
                kArr(k, 1) = i
                kArr(k, 2) = pKey2Idx(jArr(j))
                If isWgt = False Then
                    yArr(k) = tmp_x
                Else
                    yArr(k) = xArr(j) / pnode_deg_out_wgt(i)
                End If
            Next j
        Else
            j = UBound(empty_list) + 1
            ReDim Preserve empty_list(0 To j)
            empty_list(j) = i
        End If
    Next i
    DoEvents: Application.StatusBar = "Calc_Node_PageRank: 0/" & iter_max
    For iterate = 1 To iter_max
        ReDim x_new(1 To pSize)
        For k = 1 To pn_edge
            j = kArr(k, 1)
            i = kArr(k, 2) 'edge directed from j to i
            x_new(i) = x_new(i) + x(j) * yArr(k)
        Next k
        
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
    
    If strType = "OUT" Then Call EdgeReverseAll
    
End Sub


Sub Calc_Node_Closeness(Optional isWgt As Boolean = True, Optional strType As String = "OUT")
Dim i As Long, j As Long, k As Long, n As Long
Dim iArr() As Long, isVisited() As Long
Dim dist() As Double, tmp_x As Double
    Call Calc_Key2Idx
    If strType = "OUT" Then
        ReDim pnode_closeness_out(1 To pSize)
    ElseIf strType = "IN" Then
        ReDim pnode_closeness_in(1 To pSize)
        Call EdgeReverseAll
    End If
    iArr = pNodeList.GetList
    For i = 1 To pSize
        If i Mod 10 = 0 Then DoEvents: Application.StatusBar = "cGraph_Dir:Calc_Node_Closeness: " & i & "/" & pSize
        Call Dijkstra_Algorithm(iArr(i), dist, isVisited, isWgt)
        n = 0
        tmp_x = 0
        For j = 1 To pSize
            If j <> i And isVisited(j) > 0 Then
                n = n + 1
                tmp_x = tmp_x + dist(j)
            End If
        Next j
        If strType = "OUT" Then
            If n > 0 Then pnode_closeness_out(i) = ((n / pSize) ^ 2) * (pSize - 1) / tmp_x
        ElseIf strType = "IN" Then
            If n > 0 Then pnode_closeness_in(i) = ((n / pSize) ^ 2) * (pSize - 1) / tmp_x
        End If
    Next i
    If strType = "IN" Then Call EdgeReverseAll
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


'=================================================================================
'"An O(m) algorithm for Cores Decomposition of Networks", Vladimir Batagelj (2003)
'=================================================================================
Sub Calc_Node_Cores(Optional strType As String = "OUT")
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long, jj As Long
Dim w As Long, deg_max As Long, deg_u As Long, pos_u As Long, pos_w As Long
Dim deg_bin() As Long, vpos() As Long, vert() As Long, neighbor_list() As Long, iArr() As Long
Dim deg() As Long, node_list() As Long
    node_list = pNodeList.GetList
    Call Calc_Key2Idx
    If strType = "OUT" Then
        Call Calc_Node_Deg_Out: deg = pnode_deg_out
    ElseIf strType = "IN" Then
        Call Calc_Node_Deg_In: deg = pnode_deg_in
    ElseIf strType = "TOTAL" Then
        Call Calc_Node_Deg: deg = pnode_deg
    End If
    
    deg_max = -1
    For i = 1 To pSize
        If deg(i) > deg_max Then deg_max = deg(i)
    Next i
    
    ReDim deg_bin(0 To deg_max)
    For i = 1 To pSize
        deg_bin(deg(i)) = deg_bin(deg(i)) + 1
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
        k = deg(i)
        vpos(i) = deg_bin(k)
        vert(vpos(i)) = i
        deg_bin(k) = deg_bin(k) + 1
    Next i
    For k = deg_max To 1 Step -1
        deg_bin(k) = deg_bin(k - 1)
    Next k
    deg_bin(0) = 1

    For ii = 1 To pSize
        i = vert(ii)
        If strType = "OUT" Then
            neighbor_list = pAdjLists(i).GetList
        ElseIf strType = "IN" Then
            neighbor_list = pAdjLists_in(i).GetList
        ElseIf strType = "TOTAL" Then
            neighbor_list = node_neighbors(node_list(i), "TOTAL")
        End If
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
    
    If strType = "OUT" Then
        pnode_core_out = deg
    ElseIf strType = "IN" Then
        pnode_core_in = deg
    ElseIf strType = "TOTAL" Then
        pnode_core = deg
    End If
End Sub


Sub Calc_Node_Katz(Optional alpha As Double = 0.1, Optional beta As Double = 1, _
    Optional isWgt As Boolean = False, Optional strLR As String = "RIGHT")
Dim i As Long, j As Long, k As Long, m As Long, n As Long, ii As Long
Dim iterate As Long, iter_max As Long
Dim tmp_x As Double, tmp_y As Double, tol As Double
Dim x() As Double, x_new() As Double, iArr() As Long, jArr() As Long, kArr() As Long
Dim xArr() As Double, yArr() As Double, lambda_max As Double, alpha1 As Double, beta1 As Double
Dim eigen_old() As Double, eigen_allocated As Boolean
    eigen_allocated = isAllocated(pnode_eigen)
    If eigen_allocated = True Then eigen_old = pnode_eigen
    Call Calc_Node_Eigen(isWgt, strLR, lambda_max): Erase pnode_eigen
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
    k = 0
    Call Calc_Key2Idx
    ReDim kArr(1 To pn_edge, 1 To 2)
    ReDim yArr(1 To pn_edge)
    For i = 1 To pSize
        Call pAdjLists(i).GetListnWgt(jArr, xArr)
        For j = 1 To UBound(jArr)
            k = k + 1
            kArr(k, 1) = i
            kArr(k, 2) = pKey2Idx(jArr(j))
            If isWgt = True Then yArr(k) = xArr(j)
        Next j
    Next i
    If isWgt = False Then
        For k = 1 To pn_edge
            yArr(k) = 1
        Next k
    End If

    For iterate = 1 To iter_max

        ReDim x_new(1 To pSize)
        If strLR = "RIGHT" Then
            For k = 1 To pn_edge
                i = kArr(k, 1)
                x_new(i) = x_new(i) + x(kArr(k, 2)) * yArr(k)
            Next k
        ElseIf strLR = "LEFT" Then
            For k = 1 To pn_edge
                i = kArr(k, 2)
                x_new(i) = x_new(i) + x(kArr(k, 1)) * yArr(k)
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


Sub Calc_Node_HubAuthority(Optional isWgt As Boolean = False, Optional iter_max As Long = 1000)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_component As Long
Dim g_components As Collection, g1 As cGraph_Dir, comp_idx() As Long
Dim xArr() As Double, yArr() As Double, tmp_x As Double, iArr() As Long
    Call Disconnected_Components(g_components, comp_idx, n_component)
    If n_component = 1 Then
        Call Calc_Node_HubAuthority_Single(isWgt, iter_max)
        Exit Sub
    End If
    
    Call Calc_Key2Idx
    ReDim pnode_hub(1 To pSize)
    ReDim pnode_authority(1 To pSize)
    For m = 1 To n_component
        Set g1 = g_components(m)
        With g1
            Call .Calc_Node_HubAuthority_Single(isWgt, iter_max)
            iArr = .NodeList
            tmp_x = Sqr(pSize * 1# / .Size)
            xArr = .node_hub
            yArr = .node_authority
            For i = 1 To .Size
                j = pKey2Idx(iArr(i))
                pnode_hub(j) = xArr(i) / tmp_x
                pnode_authority(j) = yArr(i) / tmp_x
            Next i
            Call .Reset
        End With
    Next m
    For m = n_component To 1 Step -1
        g_components.Remove (m)
    Next m
    Set g_components = Nothing
End Sub


Sub Calc_Node_HubAuthority_Single(Optional isWgt As Boolean = False, Optional iter_max As Long = 1000)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, tol As Double
Dim xArr() As Double, authority_chg As Double, hub_chg As Double
    
    If pSize = 1 Then
        ReDim pnode_hub(1 To 1)
        ReDim pnode_authority(1 To 1)
        pnode_hub(1) = 1
        pnode_authority(1) = 1
        Exit Sub
    End If
    
    tol = 0.000000001
    Call Calc_EdgeList
    
    ReDim pnode_hub(1 To pSize)
    ReDim pnode_authority(1 To pSize)
    tmp_x = 1# / Sqr(pSize)
    For i = 1 To pSize
        pnode_hub(i) = tmp_x
        pnode_authority(i) = tmp_x
    Next i
    DoEvents: Application.StatusBar = "Calc_Node_HubAuthority_Single: 0/" & iter_max
    For iterate = 1 To iter_max
        If iterate Mod 50 = 0 Then
            DoEvents: Application.StatusBar = "Calc_Node_HubAuthority_Single: " & iterate & "/" & iter_max
        End If
        xArr = pnode_authority
        ReDim pnode_authority(1 To pSize)
        If isWgt = False Then
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                pnode_authority(j) = pnode_authority(j) + pnode_hub(i)
            Next k
        Else
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                pnode_authority(j) = pnode_authority(j) + pnode_hub(i) * pEdgeWgt(k)
            Next k
        End If
        authority_chg = 0
        tmp_x = 0
        For i = 1 To pSize
            tmp_x = tmp_x + pnode_authority(i) ^ 2
        Next i
        tmp_x = Sqr(tmp_x)
        For i = 1 To pSize
            pnode_authority(i) = pnode_authority(i) / tmp_x
            authority_chg = authority_chg + xArr(i) * pnode_authority(i)
        Next i
        
        xArr = pnode_hub
        ReDim pnode_hub(1 To pSize)
        If isWgt = False Then
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                pnode_hub(i) = pnode_hub(i) + pnode_authority(j)
            Next k
        Else
            For k = 1 To pn_edge
                i = pEdgeList(k, 1)
                j = pEdgeList(k, 2)
                pnode_hub(i) = pnode_hub(i) + pnode_authority(j) * pEdgeWgt(k)
            Next k
        End If
        
        hub_chg = 0
        tmp_x = 0
        For i = 1 To pSize
            tmp_x = tmp_x + pnode_hub(i) ^ 2
        Next i
        tmp_x = Sqr(tmp_x)
        For i = 1 To pSize
            pnode_hub(i) = pnode_hub(i) / tmp_x
            hub_chg = hub_chg + xArr(i) * pnode_hub(i)
        Next i
        
        If Abs(authority_chg - 1) < tol And Abs(hub_chg - 1) < tol Then Exit For
        
    Next iterate
    If iterate >= iter_max Then
        Debug.Print "Calc_Node_HubAuthority: Failed to converge in " & iter_max & " iterations."
    End If
    Application.StatusBar = False
End Sub


Function Sort_Topological(u As Long) As Long()
Dim i As Long, j As Long, k As Long, n_visited As Long
Dim sort_order() As Long, iArr() As Long
Dim x As cAdjList
Dim isFix() As Boolean, isTemp() As Boolean
Dim tmp_x As Double
    iArr = pNodeList.GetList
    Call Calc_Key2Idx
    n_visited = 0
    ReDim sort_order(1 To pSize)
    ReDim isFix(1 To pSize)
    ReDim isTemp(1 To pSize)
    Set x = New cAdjList
    Call Sort_Topological_Recursive(pKey2Idx(u), sort_order, isFix, isTemp, n_visited)
    For i = 1 To pSize
        If x.isConnect(iArr(i), tmp_x) = 0 Then
            Call Sort_Topological_Recursive(i, sort_order, isFix, isTemp, n_visited)
        End If
    Next i
    iArr = sort_order
    For i = pSize To 1 Step -1
        sort_order(iArr(i)) = pSize - i + 1
    Next i
    Sort_Topological = sort_order
End Function

Private Sub Sort_Topological_Recursive(i As Long, sort_order() As Long, _
        isFix() As Boolean, isTemp() As Boolean, n_visited As Long)
Dim j As Long, k As Long
Dim iArr() As Long
    If isFix(i) = True Then Exit Sub
    If isTemp(i) = True Then
        Debug.Print "cGraph_Dir: Sort_Topological: graph is not DAG."
        Exit Sub
    End If
    isTemp(i) = True
    iArr = pAdjLists(i).GetList
    For k = 1 To UBound(iArr)
        j = iArr(k)
        Call Sort_Topological_Recursive(pKey2Idx(j), sort_order, isFix, isTemp, n_visited)
    Next k
    isFix(i) = True
    n_visited = n_visited + 1
    sort_order(n_visited) = i
End Sub


Sub Filter_Disparity(g1 As cGraph_Dir, Optional significance As Double = 0.1, _
        Optional InvertWgt As Boolean = False)
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double, tmp_y As Double
Dim node_list() As Long, s_out As Double, s_in As Double
    node_list = NodeList
    Call Calc_Node_Deg
    Call Calc_Node_Deg_Wgt
    Call Calc_EdgeList
    
    If InvertWgt = True Then
        ReDim pnode_deg_out_wgt(1 To pSize)
        ReDim pnode_deg_in_wgt(1 To pSize)
        For k = 1 To pn_edge
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            pEdgeWgt(k) = 1 / pEdgeWgt(k)
            pnode_deg_out_wgt(i) = pnode_deg_out_wgt(i) + pEdgeWgt(k)
            pnode_deg_in_wgt(j) = pnode_deg_in_wgt(j) + pEdgeWgt(k)
        Next k
    End If
    
    Set g1 = New cGraph_Dir
    With g1
        Call .Init(pSize, NodeWgts, True, NodeList, True)
        For k = 1 To pn_edge
            i = pEdgeList(k, 1)
            j = pEdgeList(k, 2)
            tmp_y = pEdgeWgt(k)
            s_out = 1: s_in = 1
            If pnode_deg_out(i) > 1 Then s_out = (1 - tmp_y / pnode_deg_out_wgt(i)) ^ (pnode_deg_out(i) - 1)
            If pnode_deg_in(j) > 1 Then s_in = (1 - tmp_y / pnode_deg_in_wgt(j)) ^ (pnode_deg_in(j) - 1)
            If pnode_deg_out(i) = 1 And pnode_deg_in(j) = 1 And pnode_deg_in(i) > 1 And pnode_deg_out(j) > 1 Then
                Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
            ElseIf pnode_deg_out(i) = 1 And pnode_deg_in(j) > 1 Then
                If s_in < significance Then Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
            ElseIf pnode_deg_in(i) = 1 And pnode_deg_out(j) > 1 Then
                If s_out < significance Then Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
            ElseIf s_out < significance Or s_in < significance Then
                Call g1.EdgeAdd(node_list(i), node_list(j), tmp_y)
            End If
        Next k
    End With
End Sub


'Create a new directed graph where each node is a strongly connected component of
'the input graph g1. Integer array newKey() maps the new node to the old ones.
'newKey(i)=u means that the i-th node in g1 is now grouped under node with key u in the new graph.
Sub StrongConnect_Condense(g1 As cGraph_Dir, newKey() As Long, Optional adjpos As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long, n_scc As Long
Dim comp_idx() As Long, node_list() As Long, edge_list_old() As Long, node_list_old() As Long
Dim isMerge() As Long
Dim tmp_x As Double
    Call Me.Copy_from(g1)
    If adjpos = True Then pnode_pos = g1.node_pos
    Call Calc_EdgeList
    edge_list_old = pEdgeList
    node_list_old = NodeList
    n_scc = StrongConnect(comp_idx)
    ReDim isMerge(1 To pSize)
    ReDim newKey(1 To pSize)
    For i = 1 To pSize
        newKey(i) = i
    Next i
    For k = 1 To UBound(edge_list_old)
        i = edge_list_old(k, 1)
        j = edge_list_old(k, 2)
        If isMerge(i) = 0 And isMerge(j) = 0 Then
            If comp_idx(i) = comp_idx(j) Then
                u = node_list_old(i)
                v = node_list_old(j)
                Call NodeMerge(u, v, adjpos)
                isMerge(j) = 1
                newKey(j) = Merge_Parent(i, newKey)
            End If
        End If
    Next k
    For i = 1 To UBound(newKey)
        newKey(i) = Merge_Parent(newKey(i), newKey)
    Next i
    For i = 1 To UBound(newKey)
        newKey(i) = node_list_old(newKey(i))
    Next i
End Sub

Private Function Merge_Parent(u As Long, iParent() As Long) As Long
    If u <> iParent(u) Then
        iParent(u) = Merge_Parent(iParent(u), iParent)
    End If
    Merge_Parent = iParent(u)
End Function




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
Dim neigh_list() As Long, neigh_list_in() As Long
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
        'pop u from queue
        u = iQueue(n_head): n_head = n_head + 1
        'List of u's neighbors
        neigh_list = pAdjLists(u).GetList
        neigh_list_in = pAdjLists_in(u).GetList
        m = UBound(neigh_list): n = UBound(neigh_list_in)
        ReDim Preserve neigh_list(LBound(neigh_list) To m + n)
        For i = 1 To n
            neigh_list(m + i) = neigh_list_in(i)
        Next i
        Erase neigh_list_in
        'Color each of u's neighbors
        For i = 1 To UBound(neigh_list)
            w = pKey2Idx(neigh_list(i))
            'node w was already assigned same color as u, graph is not bipartite
            If node_color(w) = node_color(u) Then
                Debug.Print "isBipartite: graph is not bipartite. Cycle: " & u & " - " & w & " - " & u_parent
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
