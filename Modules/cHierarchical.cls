VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHierarchical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Main Reference:
'1. "Modern Hierarchical, agglomerative clustering algorithms", Daniel Mullner (2011)
'2. "dendsort: modular leaf ordering methods for dendogram representaions in R", Ryo Sakai (2014)
'3. "Fast optimal leaf ordering for hierarchical clustering", Ziv Bar-Joseph (2001)

Private pn_raw As Long          'size of original data
Private pZ() As Long            'First column is index of internal node, 2nd and 3rd columns are its children
Private pZ_height() As Double   'height of internal node
Private plabel As Variant       'label of leaf node
Private pleaf_order() As Long   'pleaf_order(i)=k means that leaf k is placed at the i-th position
Private psizes() As Long        'sizes of leaf+internal node
Private pxy_plot() As Double    'layout position of leaf+internal node
Private pparents() As Long      'parents of leaf+internal node
Private plink_type As String
Private pswapped_Z() As Long
Private pswapped_Z_height() As Double
Private pDistance() As Double

Sub Reset()
    Erase pZ, pZ_height, plabel, pleaf_order, psizes, pxy_plot
    Erase pparents, pDistance, pswapped_Z, pswapped_Z_height
End Sub

Public Property Get z() As Long()
    z = pZ
End Property

Public Property Get Z_height() As Double()
    Z_height = pZ_height
End Property

Public Property Get tree_height() As Double
    tree_height = pZ_height(pn_raw - 1)
End Property

Public Property Get leaf_order() As Long()
    leaf_order = pleaf_order
End Property

Public Property Get sizes() As Long()
    sizes = psizes
End Property

Public Property Get Size() As Long
    Size = pn_raw
End Property

Public Property Get xy_plot() As Double()
    xy_plot = pxy_plot
End Property

Public Property Get link_type() As String
    link_type = plink_type
End Property

Public Property Get leaves_list(i As Long) As Long()
Dim child_list() As Long
    If i > pn_raw Then
        Call Find_Children(i, child_list, 1)
    Else
        ReDim child_list(1 To 1)
        child_list(1) = i
    End If
    leaves_list = child_list
End Property


'==========================================================
'Create dendrogram for distance matrix x(), specific
'alogirthm is chosen for each linkage method.
'==========================================================
Sub Build(x() As Double, Optional link_type As String = "AVERAGE", Optional label As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long

    pn_raw = UBound(x, 1)
    pDistance = x
    plink_type = UCase(link_type)
    ReDim plabel(1 To pn_raw)
    If IsMissing(label) = True Then
        For i = 1 To pn_raw
            plabel(i) = i
        Next i
    Else
        For i = 1 To pn_raw
            plabel(i) = label(i)
        Next i
    End If
    ReDim pZ(1 To pn_raw - 1, 1 To 3)
    ReDim pZ_height(1 To pn_raw - 1)
    
    If plink_type = "MEDIAN" Or plink_type = "CENTROID" Then
        Call Build_Generic(x)
        'Call Build_Direct(x)
    ElseIf plink_type = "AVERAGE" Or plink_type = "WARD" _
            Or plink_type = "COMPLETE" Or plink_type = "WEIGHTED" Then
        Call Build_NNChain(x)
    ElseIf plink_type = "SINGLE" Then
        Call Build_MSTLinkage(x)
    End If

    Call Find_Parents       'find parents of each node
    Call Find_Sizes         'find sizes of each node
    Call Calc_Leaf_Order    'get leafs ordering
End Sub


'Create a new node w that parents over u & v
Private Sub Merge_uv(u As Long, v As Long, w As Long, iterate As Long, _
            w_height As Double, parent() As Long, height() As Double, node_size() As Long)
    parent(u) = w
    parent(v) = w
    height(w) = w_height
    node_size(w) = node_size(u) + node_size(v)
    pZ_height(iterate) = w_height
    pZ(iterate, 1) = w
    If height(u) < height(v) Then
        pZ(iterate, 2) = v
        pZ(iterate, 3) = u
    Else
        pZ(iterate, 2) = u
        pZ(iterate, 3) = v
    End If
End Sub


'=== Direct Implementation of agglomoerative clustering
Private Sub Build_Direct(x() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim ii As Long, jj As Long
Dim u As Long, v As Long, w As Long, new_node As Long
Dim distance() As Double
Dim tmp_min As Double, temp As Double
Dim tmp_x As Double, tmp_y As Double, tmp_vec() As Double
Dim height() As Double
Dim node_size() As Long, parent() As Long
Dim node_ptr() As Long, u_addr As Long, v_addr As Long, INFINITY As Double
    INFINITY = Exp(70)

    distance = x                        'Reuse index in distance matrix...
    ReDim node_ptr(1 To pn_raw)         '...node_ptr(i)=k means row/column i refers to node k
    ReDim parent(1 To 2 * pn_raw - 1)
    ReDim node_size(1 To 2 * pn_raw - 1)
    ReDim height(1 To 2 * pn_raw - 1)   'Keep track of height of each node
    For i = 1 To pn_raw
        node_size(i) = 1
        node_ptr(i) = i
    Next i
    
    '==========================
    'Start iteration
    '==========================
    new_node = pn_raw
    For iterate = 1 To pn_raw - 1

        'Find closest pair of nodes
        tmp_min = INFINITY
        For i = 1 To pn_raw - 1
            m = node_ptr(i)
            If m > 0 Then
                For j = i + 1 To pn_raw
                    n = node_ptr(j)
                    If n > 0 Then
                        If distance(i, j) < tmp_min Then
                            tmp_min = distance(i, j)
                            u = m: u_addr = i
                            v = n: v_addr = j
                        End If
                    End If
                Next j
            End If
        Next i

        '=== Create new node to join the pair
        new_node = new_node + 1
        Call Merge_uv(u, v, new_node, iterate, tmp_min, parent, height, node_size)
        node_ptr(u_addr) = -1       'Destroy u and point v to new node
        node_ptr(v_addr) = new_node

        '=== Calculate distance of the new node to other nodes
        ReDim tmp_vec(1 To pn_raw)
        Select Case UCase(plink_type)
        
        Case "CENTROID" 'Centroid method
            m = node_size(u)
            n = node_size(v)
            temp = m * n * (distance(u_addr, v_addr) / (m + n)) ^ 2
            For w = 1 To pn_raw
                If w <> u_addr And w <> v_addr And node_ptr(w) > 0 Then
                    tmp_x = distance(u_addr, w) ^ 2
                    tmp_y = distance(v_addr, w) ^ 2
                    tmp_vec(w) = Sqr((m * tmp_x + n * tmp_y) / (m + n) - temp)
                End If
            Next w

        Case "MEDIAN" 'Median method
            temp = (distance(u_addr, v_addr) ^ 2) / 4
            For w = 1 To pn_raw
                If w <> u_addr And w <> v_addr And node_ptr(w) > 0 Then
                    tmp_x = distance(u_addr, w) ^ 2
                    tmp_y = distance(v_addr, w) ^ 2
                    tmp_vec(w) = Sqr((tmp_x + tmp_y) / 2 - temp)
                End If
            Next w
            
        Case "AVERAGE" 'Average method
            m = node_size(u)
            n = node_size(v)
            For w = 1 To pn_raw
                If w <> u_addr And w <> v_addr And node_ptr(w) > 0 Then
                    tmp_x = distance(u_addr, w)
                    tmp_y = distance(v_addr, w)
                    tmp_vec(w) = (m * tmp_x + n * tmp_y) / (m + n)
                End If
            Next w

        Case "WARD" 'Ward's Method
            m = node_size(u)
            n = node_size(v)
            temp = distance(u_addr, v_addr) ^ 2
            For w = 1 To pn_raw
                If w <> u_addr And w <> v_addr And node_ptr(w) > 0 Then
                    k = node_size(node_ptr(w))
                    tmp_x = distance(u_addr, w) ^ 2
                    tmp_y = distance(v_addr, w) ^ 2
                    tmp_vec(w) = Sqr(((m + k) * tmp_x + (n + k) * tmp_y - k * temp) / (m + n + k))
                End If
            Next w

        End Select
        
        For i = 1 To pn_raw
            distance(v_addr, i) = tmp_vec(i)
            distance(i, v_addr) = tmp_vec(i)
        Next i
        '===========================================
        
        If iterate Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = pn_raw - new_node & " nodes remaining..."
        End If
    Next iterate

    Erase distance, height, node_size, parent, node_ptr
    Application.StatusBar = False
End Sub


'=== Generic linkage method from Daniel Mullner
Private Sub Build_Generic(x() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim u As Long, v As Long, w As Long, new_node As Long
Dim distance() As Double
Dim tmp_x As Double, tmp_y As Double, tmp_min As Double, temp As Double
Dim node_id() As Long, node_size() As Long, iStack() As Long, n_nghbr() As Long
Dim height() As Double, mindist() As Double, tmp_vec() As Double
Dim qHeap As cHeap
Dim INFINITY As Double

    INFINITY = Exp(70)
    distance = x
    ReDim iStack(0 To pn_raw)
    ReDim node_size(1 To pn_raw)
    ReDim node_id(1 To pn_raw)
    ReDim height(1 To pn_raw)
    For i = 1 To pn_raw
        node_size(i) = 1
        iStack(i) = i
        node_id(i) = i
    Next i
    
    'Initialize candidate list of nearest pairs
    ReDim n_nghbr(1 To pn_raw - 1)
    ReDim mindist(1 To pn_raw - 1)
    For i = 1 To pn_raw - 1
        tmp_min = INFINITY: k = -1
        For j = i + 1 To pn_raw
            If distance(i, j) < tmp_min Then
                tmp_min = distance(i, j)
                k = j
            End If
        Next j
        n_nghbr(i) = k
        mindist(i) = tmp_min
    Next i
    
    'Put mindist into a min-Heap
    Set qHeap = New cHeap
    Call qHeap.Build(mindist, "MIN")
    
    'Start iteration
    new_node = pn_raw
    For iterate = 1 To pn_raw - 1
        
        'Peep closest pair (u,v)
        Call qHeap.Top(tmp_min, u): v = n_nghbr(u)
        
        'Verify that (u,v) is indeed the closest pair
        Do While tmp_min <> distance(u, v)
            tmp_min = INFINITY: k = -1
            For j = 1 To UBound(iStack)
                w = iStack(j)
                If w > u Then
                    If distance(u, w) < tmp_min Then
                        tmp_min = distance(u, w)
                        k = w
                    End If
                End If
            Next j
            n_nghbr(u) = k
            mindist(u) = tmp_min
            Call qHeap.Update(tmp_min, u)
            Call qHeap.Top(tmp_min, u): v = n_nghbr(u)
        Loop

        'Pop closest pair from heap
        Call qHeap.Pop_Min(tmp_min, u): v = n_nghbr(u)
        
        'Merge u & v
        new_node = new_node + 1
        pZ_height(iterate) = tmp_min
        pZ(iterate, 1) = new_node
        If height(u) < height(v) Then
            pZ(iterate, 2) = node_id(v)
            pZ(iterate, 3) = node_id(u)
        Else
            pZ(iterate, 2) = node_id(u)
            pZ(iterate, 3) = node_id(v)
        End If
        height(v) = tmp_min             'height of new node
        node_id(v) = new_node           'reuse index v for new node...
        Call Array_Remove(iStack, u)    '...and destroy u
        
        'Update distance of the new node to other nodes
        ReDim tmp_vec(1 To UBound(iStack))
        Select Case UCase(plink_type)
        
        Case "CENTROID" 'Centroid method
            m = node_size(u)
            n = node_size(v)
            temp = m * n * (distance(u, v) / (m + n)) ^ 2
            For i = 1 To UBound(iStack)
                w = iStack(i)
                If w <> v Then
                    tmp_x = distance(u, w) ^ 2
                    tmp_y = distance(v, w) ^ 2
                    tmp_vec(i) = Sqr((m * tmp_x + n * tmp_y) / (m + n) - temp)
                End If
            Next i
        
        Case "MEDIAN" 'Median method
            temp = (distance(u, v) ^ 2) / 4
            For i = 1 To UBound(iStack)
                w = iStack(i)
                If w <> v Then
                    tmp_x = distance(u, w) ^ 2
                    tmp_y = distance(v, w) ^ 2
                    tmp_vec(i) = Sqr((tmp_x + tmp_y) / 2 - temp)
                End If
            Next i

        Case "AVERAGE" 'Average method
            m = node_size(u)
            n = node_size(v)
            For i = 1 To UBound(iStack)
                w = iStack(i)
                If w <> v Then
                    tmp_x = distance(u, w)
                    tmp_y = distance(v, w)
                    tmp_vec(i) = (m * tmp_x + n * tmp_y) / (m + n)
                End If
            Next i
 
        Case "WARD" 'Average method
            m = node_size(u)
            n = node_size(v)
            temp = distance(u, v) ^ 2
            For i = 1 To UBound(iStack)
                w = iStack(i)
                If w <> v Then
                    k = node_size(w)
                    tmp_x = distance(u, w) ^ 2
                    tmp_y = distance(v, w) ^ 2
                    tmp_vec(i) = Sqr(((m + k) * tmp_x + (n + k) * tmp_y - k * temp) / (m + n + k))
                End If
            Next i
 
        End Select
        
        For i = 1 To UBound(iStack)
            w = iStack(i)
            distance(w, v) = tmp_vec(i)
            distance(v, w) = tmp_vec(i)
        Next i
        Erase tmp_vec
        
        'Update size of new node
        node_size(v) = node_size(u) + node_size(v)

        'Reference to u should now point to v
        For i = 1 To UBound(iStack)
            w = iStack(i)
            If w < u Then
                If n_nghbr(w) = u Then
                    n_nghbr(w) = v
                End If
            End If
        Next i
        
        'For each node check if new node is a nearer neighbor
        For i = 1 To UBound(iStack)
            w = iStack(i)
            If w < v Then
                tmp_x = distance(w, v)
                If tmp_x < mindist(w) Then
                    n_nghbr(w) = v
                    mindist(w) = tmp_x
                    Call qHeap.Update(tmp_x, w)
                End If
            End If
        Next i
        
        'Find nearest neighbor candidate of new node
        If v < pn_raw Then
            tmp_min = INFINITY: k = -1
            For i = 1 To UBound(iStack)
                w = iStack(i)
                If w > v Then
                    If distance(w, v) < tmp_min Then
                        tmp_min = distance(w, v)
                        k = w
                    End If
                End If
            Next i
            n_nghbr(v) = k
            mindist(v) = tmp_min
            Call qHeap.Update(tmp_min, v)
        End If
        
        If iterate Mod 50 = 0 Then
            DoEvents
            Application.StatusBar = pn_raw - iterate & " nodes remaining..."
        End If
    Next iterate
    Call qHeap.Reset
    Set qHeap = Nothing
    Erase distance, height, node_size, iStack, n_nghbr, mindist
    Application.StatusBar = False
End Sub


'=== Using Nearest Neighbour Chain algorithm to speed up clustering
Private Sub Build_NNChain(x() As Double)
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim iterate As Long
Dim distance() As Double
Dim tmp_min As Double, tmp_max As Double, temp As Double
Dim tmp_x As Double, tmp_y As Double
Dim u As Long, v As Long, w As Long, new_node As Long
Dim tmp_u As Long, tmp_v As Long
Dim parent() As Long
Dim height() As Double
Dim node_size() As Long
Dim iStack() As Long, iChain() As Long

    ReDim distance(1 To 2 * pn_raw - 1, 1 To 2 * pn_raw - 1)
    For i = 1 To pn_raw - 1
        For j = i + 1 To pn_raw
            distance(i, j) = x(i, j)
            distance(j, i) = x(i, j)
        Next j
    Next i

    ReDim node_size(1 To 2 * pn_raw - 1)
    ReDim parent(1 To 2 * pn_raw - 1)
    ReDim height(1 To 2 * pn_raw - 1)
    For i = 1 To pn_raw
        node_size(i) = 1
    Next i
    
    '=== Using Nearest Neighbour Chain algorithm to speed up clustering
    '=== Start Adding internal nodes for most similar pairs
    ReDim iStack(0 To 0)
    ReDim iChain(0 To 0)
    For i = 1 To pn_raw
        Call Array_Push(iStack, i)
    Next i
    
    iterate = 0
    new_node = pn_raw
    Do While UBound(iStack) > 1

        iterate = iterate + 1
    
        If UBound(iChain) <= 3 Then
            ReDim iChain(0 To 0)
            u = iStack(1)
            v = iStack(2)
            Call Array_Push(iChain, u)
        Else
            u = iChain(UBound(iChain) - 3)
            v = iChain(UBound(iChain) - 2)
            For i = 1 To 3
                Call Array_Pop(iChain)
            Next i
        End If
    
        k = 0
        Do
            '=== Find pair of minimum dissimilarity
            tmp_min = 99999999
            If parent(v) = 0 Then
                w = v
                tmp_min = distance(u, v)
            End If
            For i = 1 To UBound(iStack)
                tmp_u = iStack(i)
                If tmp_u <> u Then
                    If distance(u, tmp_u) < tmp_min Then
                        tmp_min = distance(u, tmp_u)
                        w = tmp_u
                    End If
                End If
            Next i
            '==============================================
            
            v = u
            u = w
            Call Array_Push(iChain, u)
            
            If UBound(iChain) >= 3 Then
                If u = iChain(UBound(iChain) - 2) Then k = 1
            End If
        Loop Until k = 1
    
        '=== Attributes of the new vertex
        new_node = new_node + 1
        Call Merge_uv(u, v, new_node, iterate, tmp_min, parent, height, node_size)
        '===========================================
        
        Call Array_Remove(iStack, u)
        Call Array_Remove(iStack, v)
        
        '=== Calculate the distance of the new vertex to other vertices
        Select Case UCase(plink_type)
        Case "AVERAGE"  'Average Linkage
            m = node_size(u)
            n = node_size(v)
            For i = 1 To UBound(iStack)
                w = iStack(i)
                tmp_x = distance(u, w)
                tmp_y = distance(v, w)
                distance(new_node, w) = (m * tmp_x + n * tmp_y) / (m + n)
                distance(w, new_node) = distance(new_node, w)
            Next i

        Case "WEIGHTED"  'Average Linkage
            For i = 1 To UBound(iStack)
                w = iStack(i)
                tmp_x = distance(u, w)
                tmp_y = distance(v, w)
                distance(new_node, w) = (tmp_x + tmp_y) / 2
                distance(w, new_node) = distance(new_node, w)
            Next i

        Case "COMPLETE" 'Complete Linkage
            For i = 1 To UBound(iStack)
                w = iStack(i)
                distance(new_node, w) = distance(u, w)
                tmp_x = distance(v, w)
                If tmp_x > distance(new_node, w) Then distance(new_node, w) = tmp_x
                distance(w, new_node) = distance(new_node, w)
            Next i
            
        Case "SINGLE" 'Single Linkage
            For i = 1 To UBound(iStack)
                w = iStack(i)
                distance(new_node, w) = distance(u, w)
                tmp_x = distance(v, w)
                If tmp_x < distance(new_node, w) Then distance(new_node, w) = tmp_x
                distance(w, new_node) = distance(new_node, w)
            Next i
        
        Case "WARD" 'Ward's Method
            m = node_size(u)
            n = node_size(v)
            temp = distance(u, v) ^ 2
            For i = 1 To UBound(iStack)
                w = iStack(i)
                k = node_size(w)
                tmp_x = distance(u, w) ^ 2
                tmp_y = distance(v, w) ^ 2
                distance(w, new_node) = Sqr(((m + k) * tmp_x + (n + k) * tmp_y - k * temp) / (m + n + k))
                distance(new_node, w) = distance(w, new_node)
            Next i
    
        End Select
        '===========================================
    
        Call Array_Push(iStack, new_node)
    
        If iterate Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = UBound(iStack) & " nodes remaining..."
        End If
    Loop
    Erase distance, node_size, height, parent

    Call sort_tree          're-index Z() in order of increasing height
    Application.StatusBar = False
End Sub


'=== MST alogrithm used only for single linkage
Private Sub Build_MSTLinkage(x() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim u As Long, v As Long, w As Long, new_node As Long
Dim tmp_min As Double, temp As Double
Dim tmp_x As Double, tmp_y As Double
Dim height() As Double
Dim dist() As Double, INFINITY As Double
Dim isProcess() As Long

    INFINITY = Exp(70)
    ReDim height(1 To 2 * pn_raw - 1)   'Keep track of height of each node
    new_node = pn_raw
    
    u = 1
    ReDim dist(1 To pn_raw)
    ReDim isProcess(1 To pn_raw)
    For i = 1 To pn_raw
        dist(i) = INFINITY
    Next i
    
    For iterate = 1 To pn_raw - 1
        
        isProcess(u) = 1
        For i = 1 To pn_raw
            If isProcess(i) = 0 Then
                If x(i, u) < dist(i) Then dist(i) = x(i, u)
            End If
        Next i
        
        tmp_min = INFINITY: v = -1
        For i = 1 To pn_raw
            If isProcess(i) = 0 Then
                If dist(i) < tmp_min Then
                    tmp_min = dist(i)
                    v = i
                End If
            End If
        Next i
        
        pZ_height(iterate) = tmp_min
        pZ(iterate, 1) = pn_raw + iterate
        pZ(iterate, 2) = u
        pZ(iterate, 3) = v
        
        u = v
        
        If iterate Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = pn_raw - new_node & " nodes remaining..."
        End If
        
    Next iterate
    Erase dist
    Call Relabel_Z
    Application.StatusBar = False
End Sub

Private Sub Relabel_Z()
Dim i As Long, j As Long, n As Long, u As Long, v As Long, w As Long
Dim sort_index() As Long, Z_sorted() As Long, parent() As Long, nxt_label As Long
Dim tmp_x As Double

    Call modMath.Sort_Bubble_A(pZ_height, sort_index)
    ReDim Z_sorted(1 To pn_raw - 1, 1 To 2)
    For i = 1 To pn_raw - 1
        j = sort_index(i)
        Z_sorted(i, 1) = pZ(j, 2)
        Z_sorted(i, 2) = pZ(j, 3)
    Next i

    ReDim pZ(1 To pn_raw - 1, 1 To 3)
    For i = 1 To pn_raw - 1
        pZ(i, 1) = pn_raw + i
    Next i

    ReDim parent(1 To 2 * pn_raw - 1)
    nxt_label = pn_raw + 1

    For i = 1 To pn_raw - 1

        u = Z_sorted(i, 1)
        v = Z_sorted(i, 2)
        
        'U.EFFICIENT_FIND(a)
        w = u
        Do While parent(u) <> 0
            u = parent(u)
        Loop
        Do While parent(w) <> u And parent(w) <> 0
            j = w
            w = parent(w)
            parent(j) = u
        Loop
        
        'U.EFFICIENT_FIND(b)
        w = v
        Do While parent(v) <> 0
            v = parent(v)
        Loop
        Do While parent(w) <> v And parent(w) <> 0
            j = w
            w = parent(w)
            parent(j) = v
        Loop
        
        'Append(U.EFFICIENT_FIND(a),U.EFFICIENT_FIND(b)) to L'
        pZ(i, 2) = u
        pZ(i, 3) = v
        
        'U.UNION(a,b)
        parent(u) = nxt_label
        parent(v) = nxt_label
        nxt_label = nxt_label + 1
    Next i
    Erase Z_sorted, parent, sort_index
End Sub


Private Sub Init(x() As Double, link_type As String, Optional label As Variant)
Dim i As Long
    pn_raw = UBound(x, 1)
    pDistance = x
    plink_type = UCase(link_type)
    ReDim plabel(1 To pn_raw)
    If IsMissing(label) = True Then
        For i = 1 To pn_raw
            plabel(i) = i
        Next i
    Else
        For i = 1 To pn_raw
            plabel(i) = label(i)
        Next i
    End If
    ReDim pZ(1 To pn_raw - 1, 1 To 3)
    ReDim pZ_height(1 To pn_raw - 1)
End Sub


'=== Direct Implementation of agglomoerative clustering
'Input: label() is the string labels of the N data points
'Input: x() is the distance matrix(N by N), symmetric and zero diagonals
'Input: linkage can be "AVERAGE","COMPLETE","SINGLE","WARD"
Sub linkage(x() As Double, Optional link_type As String = "AVERAGE", Optional label As Variant)
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim iterate As Long
Dim distance() As Double
Dim tmp_min As Double, temp As Double
Dim tmp_x As Double, tmp_y As Double
Dim u As Long, v As Long, w As Long, new_node As Long
Dim parent() As Long
Dim height() As Double
Dim node_size() As Long

    If IsMissing(label) = True Then
        Call Init(x, link_type)
    Else
        Call Init(x, link_type, label)
    End If
    
    ReDim distance(1 To 2 * pn_raw - 1, 1 To 2 * pn_raw - 1)
    For i = 1 To pn_raw - 1
        For j = i + 1 To pn_raw
            distance(i, j) = x(i, j)
            distance(j, i) = x(i, j)
        Next j
    Next i

    ReDim node_size(1 To 2 * pn_raw - 1)
    ReDim parent(1 To 2 * pn_raw - 1)
    ReDim height(1 To 2 * pn_raw - 1)   'Keep track of height of each node
    For i = 1 To pn_raw
        node_size(i) = 1
    Next i
    
    '=== Direct Implementation of agglomoerative clustering
    new_node = pn_raw
    For iterate = 1 To pn_raw - 1
        
        'Find closest pair of nodes
        tmp_min = 999999999
        For i = 1 To new_node - 1
            For j = i + 1 To new_node
                If parent(i) = 0 And parent(j) = 0 Then
                    If distance(i, j) < tmp_min Then
                        tmp_min = distance(i, j)
                        u = i
                        v = j
                    End If
                End If
            Next j
        Next i
    
        '=== Create new node to join the pair
        new_node = new_node + 1
        parent(u) = new_node
        parent(v) = new_node
    
        height(new_node) = tmp_min
        node_size(new_node) = node_size(u) + node_size(v)
        pZ_height(iterate) = tmp_min
        pZ(iterate, 1) = new_node
        If height(u) < height(v) Then
            pZ(iterate, 2) = v
            pZ(iterate, 3) = u
        Else
            pZ(iterate, 2) = u
            pZ(iterate, 3) = v
        End If
        '===========================================
        
        '=== Calculate the distance of the new node to other nodes
        Select Case UCase(link_type)
        Case "AVERAGE"  'Average Linkage
            m = node_size(u)
            n = node_size(v)
            For w = 1 To new_node - 1
                tmp_x = distance(u, w)
                tmp_y = distance(v, w)
                distance(new_node, w) = (m * tmp_x + n * tmp_y) / (m + n)
                distance(w, new_node) = distance(new_node, w)
            Next w
        
        Case "WEIGHTED"  'Average Linkage
            For w = 1 To new_node - 1
                tmp_x = distance(u, w)
                tmp_y = distance(v, w)
                distance(new_node, w) = (tmp_x + tmp_y) / 2
                distance(w, new_node) = distance(new_node, w)
            Next w
        
        Case "COMPLETE" 'Complete Linkage
            For w = 1 To new_node - 1
                distance(new_node, w) = distance(u, w)
                tmp_x = distance(v, w)
                If tmp_x > distance(new_node, w) Then distance(new_node, w) = tmp_x
                distance(w, new_node) = distance(new_node, w)
            Next w
            
        Case "SINGLE" 'Single Linkage
            For w = 1 To new_node - 1
                distance(new_node, w) = distance(u, w)
                tmp_x = distance(v, w)
                If tmp_x < distance(new_node, w) Then distance(new_node, w) = tmp_x
                distance(w, new_node) = distance(new_node, w)
            Next w
        
        Case "WARD" 'Ward's Method
            m = node_size(u)
            n = node_size(v)
            temp = distance(u, v) ^ 2
            For w = 1 To new_node - 1
                k = node_size(w)
                tmp_x = distance(u, w) ^ 2
                tmp_y = distance(v, w) ^ 2
                distance(w, new_node) = Sqr(((m + k) * tmp_x + (n + k) * tmp_y - k * temp) / (m + n + k))
                distance(new_node, w) = distance(w, new_node)
            Next w
            
        Case "CENTROID" 'Centroid method
            m = node_size(u)
            n = node_size(v)
            temp = m * n * (distance(u, v) / (m + n)) ^ 2
            For w = 1 To new_node - 1
                tmp_x = distance(u, w) ^ 2
                tmp_y = distance(v, w) ^ 2
                distance(w, new_node) = Sqr((m * tmp_x + n * tmp_y) / (m + n) - temp)
                distance(new_node, w) = distance(w, new_node)
            Next w
        
        Case "MEDIAN" 'Centroid method
            temp = (distance(u, v) ^ 2) / 4
            For w = 1 To new_node - 1
                tmp_x = distance(u, w) ^ 2
                tmp_y = distance(v, w) ^ 2
                distance(w, new_node) = Sqr((tmp_x + tmp_y) / 2 - temp)
                distance(new_node, w) = distance(w, new_node)
            Next w
            
        End Select
        '===========================================
    
        If iterate Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = pn_raw - new_node & " nodes remaining..."
        End If
        
    Next iterate
    Erase distance, height, node_size, parent
    
    Call Find_Parents           'find parents of each node
    Call Find_Sizes             'find sizes of each node
    Call Calc_Leaf_Order        'get leafs ordering
    Application.StatusBar = False
End Sub

'=== Using Nearest Neighbour Chain algorithm to speed up clustering
'Input: label() is the string labels of the N data points
'Input: x() is the distance matrix(N by N), symmetric and zero diagonals
'Input: linkage can be "AVERAGE","COMPLETE","SINGLE","WARD"
Sub NNChainLinkage(x() As Double, Optional link_type As String = "AVERAGE", Optional label As Variant)
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim iterate As Long
Dim distance() As Double
Dim tmp_min As Double, tmp_max As Double, temp As Double
Dim tmp_x As Double, tmp_y As Double
Dim u As Long, v As Long, w As Long, new_node As Long
Dim tmp_u As Long, tmp_v As Long
Dim parent() As Long
Dim height() As Double
Dim node_size() As Long

    If IsMissing(label) = True Then
        Call Init(x, link_type)
    Else
        Call Init(x, link_type, label)
    End If

    ReDim distance(1 To 2 * pn_raw - 1, 1 To 2 * pn_raw - 1)
    For i = 1 To pn_raw - 1
        For j = i + 1 To pn_raw
            distance(i, j) = x(i, j)
            distance(j, i) = x(i, j)
        Next j
    Next i

    ReDim node_size(1 To 2 * pn_raw - 1)
    ReDim parent(1 To 2 * pn_raw - 1)
    ReDim height(1 To 2 * pn_raw - 1)
    
    '=== Using Nearest Neighbour Chain algorithm to speed up clustering
    '=== Start Adding internal nodes for most similar pairs
    Dim iStack() As Long, iChain() As Long
    ReDim iStack(0 To 0)
    ReDim iChain(0 To 0)
    For i = 1 To pn_raw
        Call Array_Push(iStack, i)
        node_size(i) = 1
    Next i
    
    iterate = 0
    new_node = pn_raw
    Do While UBound(iStack) > 1

        iterate = iterate + 1
    
        If UBound(iChain) <= 3 Then
            ReDim iChain(0 To 0)
            u = iStack(1)
            v = iStack(2)
            Call Array_Push(iChain, u)
        Else
            u = iChain(UBound(iChain) - 3)
            v = iChain(UBound(iChain) - 2)
            For i = 1 To 3
                Call Array_Pop(iChain)
            Next i
        End If
    
        k = 0
        Do
            '=== Find pair of minimum dissimilarity
            tmp_min = 99999999
            If parent(v) = 0 Then
                w = v
                tmp_min = distance(u, v)
            End If
            For i = 1 To UBound(iStack)
                tmp_u = iStack(i)
                If tmp_u <> u Then
                    If distance(u, tmp_u) < tmp_min Then
                        tmp_min = distance(u, tmp_u)
                        w = tmp_u
                    End If
                End If
            Next i
            '==============================================
            
            v = u
            u = w
            Call Array_Push(iChain, u)
            
            If UBound(iChain) >= 3 Then
                If u = iChain(UBound(iChain) - 2) Then k = 1
            End If
        Loop Until k = 1
    
        '=== Attributes of the new vertex
        new_node = new_node + 1
        parent(u) = new_node
        parent(v) = new_node
        
        height(new_node) = tmp_min
        node_size(new_node) = node_size(u) + node_size(v)
        pZ_height(iterate) = tmp_min
        pZ(iterate, 1) = new_node
        If height(u) < height(v) Then
            pZ(iterate, 2) = v
            pZ(iterate, 3) = u
        Else
            pZ(iterate, 2) = u
            pZ(iterate, 3) = v
        End If
        '===========================================
        
        Call Array_Remove(iStack, u)
        Call Array_Remove(iStack, v)
        
        '=== Calculate the distance of the new vertex to other vertices
        Select Case UCase(link_type)
        Case "AVERAGE"  'Average Linkage
            m = node_size(u)
            n = node_size(v)
            For i = 1 To UBound(iStack)
                w = iStack(i)
                tmp_x = distance(u, w)
                tmp_y = distance(v, w)
                distance(new_node, w) = (m * tmp_x + n * tmp_y) / (m + n)
                distance(w, new_node) = distance(new_node, w)
            Next i

        Case "WEIGHTED"  'Average Linkage
            For i = 1 To UBound(iStack)
                w = iStack(i)
                tmp_x = distance(u, w)
                tmp_y = distance(v, w)
                distance(new_node, w) = (tmp_x + tmp_y) / 2
                distance(w, new_node) = distance(new_node, w)
            Next i

        Case "COMPLETE" 'Complete Linkage
            For i = 1 To UBound(iStack)
                w = iStack(i)
                distance(new_node, w) = distance(u, w)
                tmp_x = distance(v, w)
                If tmp_x > distance(new_node, w) Then distance(new_node, w) = tmp_x
                distance(w, new_node) = distance(new_node, w)
            Next i
            
        Case "SINGLE" 'Single Linkage
            For i = 1 To UBound(iStack)
                w = iStack(i)
                distance(new_node, w) = distance(u, w)
                tmp_x = distance(v, w)
                If tmp_x < distance(new_node, w) Then distance(new_node, w) = tmp_x
                distance(w, new_node) = distance(new_node, w)
            Next i
        
        Case "WARD" 'Ward's Method
            m = node_size(u)
            n = node_size(v)
            temp = distance(u, v) ^ 2
            For i = 1 To UBound(iStack)
                w = iStack(i)
                k = node_size(w)
                tmp_x = distance(u, w) ^ 2
                tmp_y = distance(v, w) ^ 2
                distance(w, new_node) = Sqr(((m + k) * tmp_x + (n + k) * tmp_y - k * temp) / (m + n + k))
                distance(new_node, w) = distance(w, new_node)
            Next i
    
        End Select
        '===========================================
    
        Call Array_Push(iStack, new_node)
    
        If iterate Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = UBound(iStack) & " nodes remaining..."
        End If
    Loop
    Erase distance, node_size, height, parent
    
    Call sort_tree          're-index Z() in order of increasing height
    Call Find_Parents       'find parents of each node
    Call Find_Sizes         'find sizes of each node
    Call Calc_Leaf_Order    'get leafs ordering
    Application.StatusBar = False
End Sub

're-index Z() in order of increasing height
Private Sub sort_tree()
Dim i As Long, j As Long, n As Long, u As Long, v As Long
Dim sort_index() As Long, Z_sorted() As Long, parent() As Long
    Call modMath.Sort_Bubble_A(pZ_height, sort_index)
    ReDim Z_sorted(1 To pn_raw - 1, 1 To 3)
    For i = 1 To pn_raw - 1
        j = sort_index(i)
        Z_sorted(i, 1) = pZ(j, 1)
        Z_sorted(i, 2) = pZ(j, 2)
        Z_sorted(i, 3) = pZ(j, 3)
    Next i
    
    ReDim parent(1 To 2 * pn_raw - 1)
    For i = 1 To pn_raw - 1
        Z_sorted(i, 1) = pn_raw + i
        u = Z_sorted(i, 2)
        v = Z_sorted(i, 3)
        If u > pn_raw Then
            Z_sorted(i, 2) = parent(pZ(u - pn_raw, 2))
        Else
            Z_sorted(i, 2) = u
        End If
        If v > pn_raw Then
            Z_sorted(i, 3) = parent(pZ(v - pn_raw, 2))
        Else
            Z_sorted(i, 3) = v
        End If
        parent(u) = pn_raw + i
        parent(v) = pn_raw + i
    Next i
    pZ = Z_sorted
End Sub

'=== Print Matrix
'Input: d(), NxN matrix to be visualize
'Input: when reverse is FALSE, higher value is blue, lower value is red
Sub Print_Matrix(vRng As Range, d() As Double, Optional reverse As Boolean = False)
Dim i As Long, j As Long, m As Long, n As Long, n_raw As Long
Dim vR() As Long, vG() As Long, vB() As Long
Dim temp As Double, tmp_min As Double, tmp_max As Double

    n_raw = UBound(d, 1)
    If n_raw <> pn_raw Then vRng.Value = "Matrix size does not match tree"
    
    ReDim vR(1 To n_raw, 1 To n_raw)
    ReDim vG(1 To n_raw, 1 To n_raw)
    ReDim vB(1 To n_raw, 1 To n_raw)
    tmp_min = 999999
    tmp_max = -999999
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            If reverse = False Then
                temp = d(i, j)
            Else
                temp = -d(i, j)
            End If
            If temp < tmp_min Then tmp_min = temp
            If temp > tmp_max Then tmp_max = temp
        Next j
    Next i
    
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            If reverse = False Then
                temp = (d(i, j) - tmp_min) / (tmp_max - tmp_min)
            Else
                temp = (-d(i, j) - tmp_min) / (tmp_max - tmp_min)
            End If
            Call Color_Scale(temp, vR(i, j), vG(i, j), vB(i, j))
            vR(j, i) = vR(i, j)
            vG(j, i) = vG(i, j)
            vB(j, i) = vB(i, j)
        Next j
    Next i
    
    With vRng
        For i = 1 To n_raw
            If i Mod 20 = 0 Then
                DoEvents
                Application.StatusBar = "Printing matrix..." & i & "/" & n_raw
            End If
            m = pleaf_order(i)
            .Offset(m, 0) = plabel(i)
            .Offset(0, m) = plabel(i)
            For j = 1 To n_raw
                n = pleaf_order(j)
                If i <> j Then
                    .Offset(m, n).Interior.Color = RGB(vR(i, j), vG(i, j), vB(i, j))
                End If
            Next j
        Next i
    End With
    Erase vR, vG, vB
    Application.StatusBar = False
End Sub

'Input: x is a real number between 0 and 1
'Output: vR,vG,vB are integers from 0 to 255
Private Sub Color_Scale(x As Double, vR As Long, vG As Long, vB As Long)
'    Dim i As Long
'    i = 255 - 200 * x
'    vR = i
'    vG = i
'    vB = i
    If x <= 0.5 Then
        vR = 255
        vG = Round(510 * x, 0)
        vB = 0
    Else
        vR = Round(-510 * (x - 1), 0)
        vG = vR
        vB = Round(510 * x - 255, 0)
    End If
End Sub

'=== Print tree
Sub Print_Tree(vRng As Range, Optional skew As Double = 1, Optional circular As Boolean = False, _
            Optional cut_level As Double = 1)
Dim i As Long, j As Long, u As Long, v As Long, w As Long, n_max As Long
Dim tree_height As Double
    Call get_xy_plot(skew, circular)   'get (x,y)-coordinates for visualization
    
    'Truncate tree at n_max
    n_max = pn_raw - 1
    tree_height = pZ_height(pn_raw - 1)
    If cut_level < 1 Then
        tree_height = tree_height * cut_level
        If tree_height < pZ_height(1) Then
            n_max = 0
        Else
            For i = pn_raw - 1 To 1 Step -1
                If pZ_height(i) <= tree_height Then
                    n_max = i
                    Exit For
                End If
            Next i
        End If
    End If

    With vRng
        For i = 1 To pn_raw
            j = pleaf_order(i)
            .Offset(j - 1, 0).Value = plabel(i)
            .Offset(j - 1, 1).Value = pxy_plot(i, 1)
            .Offset(j - 1, 2).Value = pxy_plot(i, 2)
        Next i
        If circular = False Then
            j = 0
            For i = 1 To n_max
                w = pn_raw + i
                u = pZ(i, 2)
                v = pZ(i, 3)
                .Offset(j, 3).Value = pxy_plot(u, 1)
                .Offset(j, 4).Value = pxy_plot(u, 2)
                .Offset(j + 1, 3).Value = pxy_plot(w, 1)
                .Offset(j + 1, 4).Value = pxy_plot(u, 2)
                .Offset(j + 2, 3).Value = pxy_plot(w, 1)
                .Offset(j + 2, 4).Value = pxy_plot(v, 2)
                .Offset(j + 3, 3).Value = pxy_plot(v, 1)
                .Offset(j + 3, 4).Value = pxy_plot(v, 2)
                j = j + 5
            Next i
        Else
            j = 0
            For i = 1 To n_max
                w = pn_raw + i
                u = pZ(i, 2)
                v = pZ(i, 3)
                .Offset(j, 3).Value = pxy_plot(u, 1)
                .Offset(j, 4).Value = pxy_plot(u, 2)
                .Offset(j + 1, 3).Value = pxy_plot(w, 1)
                .Offset(j + 1, 4).Value = pxy_plot(w, 2)
                .Offset(j + 2, 3).Value = pxy_plot(v, 1)
                .Offset(j + 2, 4).Value = pxy_plot(v, 2)
                j = j + 4
            Next i
        End If
    End With
End Sub

Private Sub get_xy_plot(skew As Double, Optional circular As Boolean = False)
Dim i As Long
Dim theta() As Double, tmp_x As Double
ReDim pxy_plot(1 To 2 * pn_raw - 1, 1 To 2)
If circular = False Then
    For i = 1 To pn_raw
        pxy_plot(i, 1) = 1
        pxy_plot(i, 2) = pleaf_order(i)
    Next i
    For i = 1 To pn_raw - 1
        pxy_plot(i + pn_raw, 1) = 1 + pZ_height(i) ^ skew
        pxy_plot(i + pn_raw, 2) = (pxy_plot(pZ(i, 2), 2) + pxy_plot(pZ(i, 3), 2)) / 2
    Next i
Else
    ReDim theta(1 To 2 * pn_raw - 1)
    ReDim radius(1 To pn_raw - 1)
    For i = 1 To pn_raw
        theta(i) = 6.28318530717959 * pleaf_order(i) / pn_raw
        pxy_plot(i, 1) = Cos(theta(i))
        pxy_plot(i, 2) = Sin(theta(i))
    Next i
    For i = 1 To pn_raw - 1
        tmp_x = (1 - (pZ_height(i) / pZ_height(pn_raw - 1))) ^ skew
        theta(i + pn_raw) = (theta(pZ(i, 2)) + theta(pZ(i, 3))) / 2
        pxy_plot(i + pn_raw, 1) = tmp_x * Cos(theta(i + pn_raw))
        pxy_plot(i + pn_raw, 2) = tmp_x * Sin(theta(i + pn_raw))
    Next i
End If
End Sub


Private Sub Find_Parents()
Dim i As Long
    ReDim pparents(1 To 2 * pn_raw - 1)
    For i = 1 To pn_raw - 1
        pparents(pZ(i, 2)) = pZ(i, 1)
        pparents(pZ(i, 3)) = pZ(i, 1)
    Next i
End Sub

Private Sub Find_Sizes()
Dim i As Long
    ReDim psizes(1 To 2 * pn_raw - 1)
    For i = 1 To pn_raw
        psizes(i) = 1
    Next i
    For i = 1 To pn_raw - 1
        psizes(pn_raw + i) = psizes(pZ(i, 2)) + psizes(pZ(i, 3))
    Next i
End Sub

'Get leaves order based on current Z()
Private Sub Calc_Leaf_Order()
Dim u As Long, v As Long, k As Long
Dim visited() As Long
    ReDim pleaf_order(1 To pn_raw)
    ReDim visited(1 To 2 * pn_raw - 1)
    k = 0
    u = 2 * pn_raw - 1
    visited(u) = 1
    Do While k < pn_raw
        If k Mod 50 = 0 Then Application.StatusBar = "Ordering leaves..." & k & "/" & pn_raw
        If visited(pZ(u - pn_raw, 2)) = 0 Then
            u = pZ(u - pn_raw, 2)
            visited(u) = 1
            If u <= pn_raw Then
                k = k + 1
                pleaf_order(u) = k
                u = pparents(u)
            End If
        ElseIf visited(pZ(u - pn_raw, 3)) = 0 Then
            u = pZ(u - pn_raw, 3)
            visited(u) = 1
            If u <= pn_raw Then
                k = k + 1
                pleaf_order(u) = k
                u = pparents(u)
            End If
        Else
            u = pparents(u)
        End If
    Loop
    Application.StatusBar = False
End Sub

'get original distance and cophenetic distance side by side in condensed form
Function CopheneticDist() As Double()
Dim i As Long, j As Long, k As Long
Dim cd() As Double, x() As Double
    cd = Cophenetic_Matrix()
    ReDim x(1 To pn_raw * (pn_raw - 1) / 2, 1 To 2)
    For i = 1 To pn_raw - 1
        For j = i + 1 To pn_raw
            k = k + 1
            x(k, 1) = pDistance(i, j)
            x(k, 2) = cd(i, j)
        Next j
    Next i
    CopheneticDist = x
End Function

'Calculate pairwise cophenetic distance
Private Function Cophenetic_Matrix() As Double()
Dim i As Long, j As Long, w As Long
Dim y() As Double
    ReDim y(1 To pn_raw, 1 To pn_raw)
    For i = 1 To pn_raw - 1
        DoEvents
        If i Mod 10 = 0 Then Application.StatusBar = "calculating cophenetic distance..." & i & "/" & pn_raw
        For j = i + 1 To pn_raw
            w = LeastCommonAncestor(i, j)
            y(i, j) = pZ_height(w - pn_raw)
            y(j, i) = y(i, j)
        Next j
    Next i
    Cophenetic_Matrix = y
    Application.StatusBar = False
End Function

Private Function LeastCommonAncestor(u As Long, v As Long) As Long
Dim path_u() As Long, path_v() As Long
Dim i As Long, j As Long, k As Long
    ReDim path_u(0 To 0)
    ReDim path_v(0 To 0)
    i = u
    Do
        i = pparents(i)
        Call Enqueue(path_u, i)
    Loop Until pparents(i) = 0
    i = v
    Do
        i = pparents(i)
        Call Enqueue(path_v, i)
    Loop Until pparents(i) = 0
    
    i = 1
    Do While i <= UBound(path_u) And i <= UBound(path_v)
        If path_u(i) <> path_v(i) Then Exit Do
        i = i + 1
    Loop
    LeastCommonAncestor = path_u(i - 1)
End Function


'=== Re-order pZ() and pleaf_order() by optimization  ==========================================
'Ziv Bar-Joseph "Fast optimal leaf ordering for hierarchical clustering", 2001
'Note: Simiarlity matrix is used in the paper, here we use distances, so the logic is reversed
'they maximize the cost function M, but we minimize it here
'======================================================================
Sub optOrdering()
Dim i As Long, j As Long, u As Long, v As Long, w As Long, v_L As Long, v_R As Long
Dim m As Long, n As Long
Dim sorted_Z() As Long, sorted_D() As Double, sorted_leaves() As Long
Dim cluster_ranges() As Long
Dim must_swap() As Long

Application.StatusBar = "Optimizing leaf ordering..."

're-index leaves according to their order in original tree
ReDim sorted_leaves(1 To pn_raw)
ReDim sorted_Z(1 To pn_raw - 1, 1 To 2)
ReDim sorted_D(1 To pn_raw, 1 To pn_raw)

For i = 1 To pn_raw
    sorted_leaves(pleaf_order(i)) = i
Next i

For i = 1 To pn_raw - 1
    v_L = pZ(i, 2)
    v_R = pZ(i, 3)
    If v_L <= pn_raw Then v_L = pleaf_order(v_L)
    If v_R <= pn_raw Then v_R = pleaf_order(v_R)
    sorted_Z(i, 1) = v_L
    sorted_Z(i, 2) = v_R
Next i

For i = 1 To pn_raw - 1
    For j = i + 1 To pn_raw
        m = pleaf_order(i)
        n = pleaf_order(j)
        sorted_D(m, n) = pDistance(i, j)
        sorted_D(n, m) = sorted_D(m, n)
    Next j
Next i

ReDim cluster_ranges(1 To 2 * pn_raw - 1, 1 To 2)
For i = 1 To pn_raw
    cluster_ranges(i, 1) = i
    cluster_ranges(i, 2) = i
Next i
For i = 1 To pn_raw - 1
    cluster_ranges(pn_raw + i, 1) = cluster_ranges(sorted_Z(i, 1), 1)
    cluster_ranges(pn_raw + i, 2) = cluster_ranges(sorted_Z(i, 2), 2)
Next i

Dim M_order As Variant
Dim M_inner() As Double
ReDim M_order(1 To 2 * pn_raw - 1)
For i = 1 To 2 * pn_raw - 1
    ReDim M_inner(cluster_ranges(i, 1) To cluster_ranges(i, 2), cluster_ranges(i, 1) To cluster_ranges(i, 2))
    M_order(i) = M_inner
Next i

Call optOrdering_FindOrder(2 * pn_raw - 1, M_order, sorted_Z, sorted_D, cluster_ranges)

Call optOrdering_backtrack(2 * pn_raw - 1, M_order, sorted_Z, sorted_D, cluster_ranges)

For v = 1 To pn_raw - 1
    i = sorted_Z(v, 1)
    j = sorted_Z(v, 2)
    If i <= pn_raw Then i = sorted_leaves(i)
    If j <= pn_raw Then j = sorted_leaves(j)
    pZ(v, 2) = i
    pZ(v, 3) = j
Next v
Erase sorted_Z, sorted_D, cluster_ranges, M_order

Call Calc_Leaf_Order
Application.StatusBar = False

End Sub

Private Sub optOrdering_FindOrder(v As Long, M_order As Variant, sorted_Z() As Long, sorted_D() As Double, cluster_ranges() As Long)
Dim i As Long, j As Long, k As Long, m As Long, u As Long, w As Long
Dim v_L As Long, v_R As Long, v_LL As Long, v_LR As Long, v_RL As Long, v_RR As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_min As Double

If v > pn_raw Then
    DoEvents
    
    v_L = sorted_Z(v - pn_raw, 1)
    v_R = sorted_Z(v - pn_raw, 2)
    
    If v_L > pn_raw Then
        v_LL = sorted_Z(v_L - pn_raw, 1)
        v_LR = sorted_Z(v_L - pn_raw, 2)
    Else
        v_LL = v_L
        v_LR = v_L
    End If
    If v_R > pn_raw Then
        v_RL = sorted_Z(v_R - pn_raw, 1)
        v_RR = sorted_Z(v_R - pn_raw, 2)
    Else
        v_RL = v_R
        v_RR = v_R
    End If

    Call optOrdering_FindOrder(v_L, M_order, sorted_Z, sorted_D, cluster_ranges)
    Call optOrdering_FindOrder(v_R, M_order, sorted_Z, sorted_D, cluster_ranges)

    For u = cluster_ranges(v_L, 1) To cluster_ranges(v_L, 2)
        If u <= cluster_ranges(v_LL, 2) Then
            i = v_LR
        Else
            i = v_LL
        End If
        For w = cluster_ranges(v_R, 1) To cluster_ranges(v_R, 2)
            If w <= cluster_ranges(v_RL, 2) Then
                j = v_RR
            Else
                j = v_RL
            End If
            tmp_min = 1073741824
            For m = cluster_ranges(i, 1) To cluster_ranges(i, 2)
                For k = cluster_ranges(j, 1) To cluster_ranges(j, 2)
                    tmp_x = M_order(v_L)(u, m) + M_order(v_R)(w, k) + sorted_D(m, k)
                    If tmp_x < tmp_min Then tmp_min = tmp_x
                Next k
            Next m
            M_order(v)(u, w) = tmp_min
            M_order(v)(w, u) = tmp_min
        Next w
    Next u
    
    If (v - pn_raw) Mod 50 = 0 Then Application.StatusBar = "Optimizing leaf ordering..." & v - pn_raw & "/" & pn_raw
End If

End Sub


Private Sub optOrdering_backtrack(v As Long, M_order As Variant, sorted_Z() As Long, sorted_D() As Double, cluster_ranges() As Long)
Dim i As Long, j As Long, k As Long, m As Long, u As Long, w As Long
Dim v_L As Long, v_R As Long, v_LL As Long, v_LR As Long, v_RL As Long, v_RR As Long
Dim tmp_x As Double, tmp_min As Double

If v > pn_raw Then

    v_L = sorted_Z(v - pn_raw, 1)
    v_R = sorted_Z(v - pn_raw, 2)

    If v_L > pn_raw Then
        v_LL = sorted_Z(v_L - pn_raw, 1)
        v_LR = sorted_Z(v_L - pn_raw, 2)
    Else
        v_LL = v_L
        v_LR = v_L
    End If
    If v_R > pn_raw Then
        v_RL = sorted_Z(v_R - pn_raw, 1)
        v_RR = sorted_Z(v_R - pn_raw, 2)
    Else
        v_RL = v_R
        v_RR = v_R
    End If

    tmp_min = 1073741824
    For u = cluster_ranges(v_L, 1) To cluster_ranges(v_L, 2)
        For w = cluster_ranges(v_R, 1) To cluster_ranges(v_R, 2)
            If M_order(v)(u, w) < tmp_min Then
                tmp_min = M_order(v)(u, w)
                i = u
                j = w
            End If
        Next w
    Next u
    
    If i > cluster_ranges(v_LL, 2) And v_L > pn_raw Then
        sorted_Z(v_L - pn_raw, 1) = v_LR
        sorted_Z(v_L - pn_raw, 2) = v_LL
    End If
    
    If j <= cluster_ranges(v_RL, 2) And v_R > pn_raw Then
        sorted_Z(v_R - pn_raw, 1) = v_RR
        sorted_Z(v_R - pn_raw, 2) = v_RL
    End If
    
    Call optOrdering_backtrack(v_L, M_order, sorted_Z, sorted_D, cluster_ranges)
    Call optOrdering_backtrack(v_R, M_order, sorted_Z, sorted_D, cluster_ranges)
End If


End Sub



'=== Re-order pZ() and pleaf_order() by optimization  ==========================================
'Ziv Bar-Joseph "Fast optimal leaf ordering for hierarchical clustering", 2001
'Note: Simiarlity matrix is used in the paper, here we use distances, so the logic is reversed
'they maximize the cost function M, but we minimize it here
'======================================================================
Sub Optimal_leaf_ordering()
Dim i As Long, j As Long, u As Long, v As Long, v_L As Long, v_R As Long
Dim m As Long, n As Long
Dim sorted_Z() As Long, sorted_D() As Double, sorted_leaves() As Long
Dim cluster_ranges() As Long
Dim must_swap() As Long

Application.StatusBar = "Optimizing leaf ordering..."

ReDim sorted_leaves(1 To pn_raw)
For i = 1 To pn_raw
    sorted_leaves(pleaf_order(i)) = i
Next i

ReDim sorted_Z(1 To pn_raw - 1, 1 To 2)
For i = 1 To pn_raw - 1
    v_L = pZ(i, 2)
    v_R = pZ(i, 3)
    If v_L <= pn_raw Then v_L = pleaf_order(v_L)
    If v_R <= pn_raw Then v_R = pleaf_order(v_R)
    sorted_Z(i, 1) = v_L
    sorted_Z(i, 2) = v_R
Next i

ReDim sorted_D(1 To pn_raw, 1 To pn_raw)
For i = 1 To pn_raw - 1
    For j = i + 1 To pn_raw
        m = pleaf_order(i)
        n = pleaf_order(j)
        sorted_D(m, n) = pDistance(i, j)
        sorted_D(n, m) = sorted_D(m, n)
    Next j
Next i

ReDim cluster_ranges(1 To 2 * pn_raw - 1, 1 To 2)
For i = 1 To pn_raw
    cluster_ranges(i, 1) = i
    cluster_ranges(i, 2) = i
Next i
For i = 1 To pn_raw - 1
    cluster_ranges(pn_raw + i, 1) = cluster_ranges(sorted_Z(i, 1), 1)
    cluster_ranges(pn_raw + i, 2) = cluster_ranges(sorted_Z(i, 2), 2)
Next i

must_swap = identify_swaps(sorted_Z, sorted_D, cluster_ranges)

Dim is_cluster_descendant() As Long
Dim applied_swap() As Long
Dim final_swap() As Long
Dim swapped_Z() As Long
Dim swapped_Z_height() As Double
Dim out_L As Long, out_R As Long, in_L As Long, in_R As Long

ReDim is_cluster_descendant(1 To pn_raw - 1, 1 To pn_raw - 1)
For i = 1 To pn_raw - 1
    is_cluster_descendant(i, i) = 1
    v_L = sorted_Z(i, 1)
    v_R = sorted_Z(i, 2)
    If v_L > pn_raw Then
        is_cluster_descendant(i, v_L - pn_raw) = 1
        For j = 1 To pn_raw - 1
            is_cluster_descendant(i, j) = is_cluster_descendant(i, j) + is_cluster_descendant(v_L - pn_raw, j)
        Next j
    End If
    If v_R > pn_raw Then
        is_cluster_descendant(i, v_R - pn_raw) = 1
        For j = 1 To pn_raw - 1
            is_cluster_descendant(i, j) = is_cluster_descendant(i, j) + is_cluster_descendant(v_R - pn_raw, j)
        Next j
    End If
Next i

ReDim applied_swap(1 To pn_raw - 1, 1 To pn_raw - 1)
For i = 1 To pn_raw - 1
    For j = 1 To pn_raw - 1
        If is_cluster_descendant(i, j) > 0 Then applied_swap(i, j) = must_swap(j)
    Next j
Next i

ReDim final_swap(1 To pn_raw - 1)
For i = 1 To pn_raw - 1
    For j = 1 To pn_raw - 1
        final_swap(i) = final_swap(i) + applied_swap(i, j)
    Next j
    final_swap(i) = final_swap(i) Mod 2
Next i

ReDim swapped_Z(1 To pn_raw - 1, 1 To 3)
For i = 1 To pn_raw - 1
    in_L = pZ(i, 2)
    in_R = pZ(i, 3)
    If final_swap(i) > 0 Then
        out_L = in_R
        out_R = in_L
    Else
        out_R = in_R
        out_L = in_L
    End If
    swapped_Z(i, 1) = pn_raw + i
    swapped_Z(i, 2) = out_L
    swapped_Z(i, 3) = out_R
Next i

pZ = swapped_Z
Call Calc_Leaf_Order

Application.StatusBar = False
End Sub

Private Function identify_swaps(sorted_Z() As Long, sorted_D() As Double, cluster_ranges() As Long) As Long()
Dim M_order() As Double
Dim swap_status() As Long, must_swap() As Long
Dim n_raw As Long
Dim i As Long, v_L As Long, v_R As Long, v_size As Long
Dim v_L_min As Long, v_L_max As Long, v_R_min As Long, v_R_max As Long
Dim u_clusters() As Long, m_clusters() As Long, w_clusters() As Long, k_clusters() As Long
Dim total_u_clusters As Long, total_w_clusters As Long
Dim u As Long, w As Long, m As Long, k As Long
Dim u_min As Long, u_max As Long, m_min As Long, m_max As Long, w_min As Long, w_max As Long, k_min As Long, k_max As Long
Dim swap_L As Long, swap_R As Long
Dim m_vals() As Double
Dim m_idx() As Long
Dim k_vals() As Double
Dim k_idx() As Long
Dim MI As Long, ki As Long
Dim min_km_dist As Double
Dim cur_min_M As Double, current_M As Double
Dim best_m As Long, best_k As Long
Dim best_u As Long, best_w As Long

ReDim M_order(1 To pn_raw, 1 To pn_raw)
ReDim swap_status(1 To pn_raw, 1 To pn_raw, 1 To 2)
ReDim must_swap(1 To pn_raw - 1)

For i = 1 To pn_raw - 1
    
    v_L = sorted_Z(i, 1)
    v_R = sorted_Z(i, 2)
    
    v_L_min = cluster_ranges(v_L, 1)
    v_L_max = cluster_ranges(v_L, 2)
    v_R_min = cluster_ranges(v_R, 1)
    v_R_max = cluster_ranges(v_R, 2)
    
    ReDim u_clusters(1 To 2)
    ReDim m_clusters(1 To 2)
    ReDim w_clusters(1 To 2)
    ReDim k_clusters(1 To 2)
    
    If v_L <= pn_raw Then
        total_u_clusters = 1
        u_clusters(1) = v_L
        m_clusters(1) = v_L
        ReDim Preserve u_clusters(1 To 1)
        ReDim Preserve m_clusters(1 To 1)
    Else
        total_u_clusters = 2
        u_clusters(1) = sorted_Z(v_L - pn_raw, 1)
        m_clusters(1) = sorted_Z(v_L - pn_raw, 2)
        u_clusters(2) = sorted_Z(v_L - pn_raw, 2)
        m_clusters(2) = sorted_Z(v_L - pn_raw, 1)
    End If

    If v_R <= pn_raw Then
        total_w_clusters = 1
        w_clusters(1) = v_R
        k_clusters(1) = v_R
        ReDim Preserve w_clusters(1 To 1)
        ReDim Preserve k_clusters(1 To 1)
    Else
        total_w_clusters = 2
        w_clusters(1) = sorted_Z(v_R - pn_raw, 2)
        w_clusters(2) = sorted_Z(v_R - pn_raw, 1)
        k_clusters(1) = sorted_Z(v_R - pn_raw, 1)
        k_clusters(2) = sorted_Z(v_R - pn_raw, 2)
    End If

    For swap_L = 1 To total_u_clusters
        For swap_R = 1 To total_w_clusters
            
            u_min = cluster_ranges(u_clusters(swap_L), 1)
            u_max = cluster_ranges(u_clusters(swap_L), 2)
            m_min = cluster_ranges(m_clusters(swap_L), 1)
            m_max = cluster_ranges(m_clusters(swap_L), 2)
            w_min = cluster_ranges(w_clusters(swap_R), 1)
            w_max = cluster_ranges(w_clusters(swap_R), 2)
            k_min = cluster_ranges(k_clusters(swap_R), 1)
            k_max = cluster_ranges(k_clusters(swap_R), 2)
            
            min_km_dist = 1073741824
            For m = m_min To m_max
                For k = k_min To k_max
                    If sorted_D(m, k) < min_km_dist Then min_km_dist = sorted_D(m, k)
                Next k
            Next m
            
            For u = u_min To u_max
            
                ReDim m_vals(1 To m_max - m_min + 1)
                ReDim m_idx(1 To m_max - m_min + 1)
                
                Call sort_M_slice(M_order, m_vals, m_idx, m_min, m_max, u)
                
                For w = w_min To w_max
                    ReDim k_vals(1 To k_max - k_min + 1)
                    ReDim k_idx(1 To k_max - k_min + 1)
                    
                    Call sort_M_slice(M_order, k_vals, k_idx, k_min, k_max, w)
                    
                    cur_min_M = 1073741824
                    For MI = 1 To m_max - m_min + 1
                        m = m_idx(MI)
                        If M_order(u, m) + M_order(w, k_idx(1)) + min_km_dist >= cur_min_M Then
                            Exit For
                        End If
                        For ki = 1 To k_max - k_min + 1
                            k = k_idx(ki)
                            
                            If M_order(u, m) + M_order(w, k) + min_km_dist >= cur_min_M Then
                                Exit For
                            End If
                               
                            current_M = M_order(u, m) + M_order(w, k) + sorted_D(m, k)
                            If current_M < cur_min_M Then
                                cur_min_M = current_M
                                best_m = m
                                best_k = k
                            End If
                    
                        Next ki
                    Next MI
                    
                    M_order(u, w) = cur_min_M
                    M_order(w, u) = cur_min_M
                    swap_status(u, w, 1) = swap_L
                    swap_status(w, u, 1) = swap_L
                    swap_status(u, w, 2) = swap_R
                    swap_status(w, u, 2) = swap_R
                    
                Next w
            Next u
            
            cur_min_M = 1073741824
            For u = v_L_min To v_L_max
                For w = v_R_min To v_R_max
                    If M_order(u, w) < cur_min_M Then
                        cur_min_M = M_order(u, w)
                        best_u = u
                        best_w = w
                    End If
                Next w
            Next u

            If v_L > pn_raw Then must_swap(v_L - pn_raw) = swap_status(best_u, best_w, 1)
            If v_R > pn_raw Then must_swap(v_R - pn_raw) = swap_status(best_u, best_w, 2)

        Next swap_R
    Next swap_L
    
Next i

identify_swaps = must_swap

End Function


Private Sub sort_M_slice(M_order() As Double, vals() As Double, idx() As Long, _
                dim1_min As Long, dim1_max As Long, dim2_val As Long)
Dim i As Long
ReDim vals(1 To dim1_max - dim1_min + 1)
ReDim idx(1 To dim1_max - dim1_min + 1)
For i = 1 To dim1_max - dim1_min + 1
    vals(i) = M_order(dim1_min + i - 1, dim2_val)
    idx(i) = dim1_min + i - 1
Next i
Call modMath.Sort_Quick_A(vals, 1, dim1_max - dim1_min + 1, idx, 0)
End Sub




'=== Re-order pZ() and pleaf_order() by optimization  ================
'Ryo Sakai "dendsort: modular leaf ordering methods for dendogram representaions in R", 2014
'======================================================================
Sub MOLO_Ordering()
    Call MOLO_Sort_Smallest(2 * pn_raw - 1)
    Call Calc_Leaf_Order
End Sub

Private Sub MOLO_Sort_Smallest(u As Long)
Dim i As Long, v_L As Long, v_R As Long
Dim d_L As Double, d_R As Double
    i = u - pn_raw
    v_L = pZ(i, 2)
    v_R = pZ(i, 3)
    If v_L > pn_raw And v_R <= pn_raw Then
    
        Call MOLO_Sort_Smallest(v_L)
        pZ(i, 2) = v_R
        pZ(i, 3) = v_L
        
    ElseIf v_L <= pn_raw And v_R > pn_raw Then
    
        Call MOLO_Sort_Smallest(v_R)
        
    ElseIf v_L > pn_raw And v_R > pn_raw Then
    
        Call MOLO_Sort_Smallest(v_L)
        Call MOLO_Sort_Smallest(v_R)
        d_L = pZ_height(i)
        d_R = pZ_height(i)
        Call MOLO_find_min_height(v_L, d_L)
        Call MOLO_find_min_height(v_R, d_R)
        If d_R > d_L Then
            pZ(i, 2) = v_R
            pZ(i, 3) = v_L
        End If
        
    End If
End Sub

Private Sub MOLO_find_min_height(u As Long, min_height As Double)
Dim v_L As Long, v_R As Long
    If u > pn_raw Then
        If pZ_height(u - pn_raw) < min_height Then
            min_height = pZ_height(u - pn_raw)
        End If
        v_L = pZ(u - pn_raw, 2)
        v_R = pZ(u - pn_raw, 3)
        If v_L > pn_raw Then Call MOLO_find_min_height(v_L, min_height)
        If v_R > pn_raw Then Call MOLO_find_min_height(v_R, min_height)
    End If
End Sub








'=== Stack Operations
Private Sub Array_Push(iStack() As Long, k As Long)
Dim n As Long
    n = UBound(iStack) + 1
    ReDim Preserve iStack(0 To n)
    iStack(n) = k
End Sub

Private Sub Array_Pop(iStack() As Long)
Dim n As Long
    n = UBound(iStack)
    ReDim Preserve iStack(0 To n - 1)
End Sub

Private Sub Array_Remove(iStack() As Long, k As Long)
Dim i As Long, j As Long, n As Long
    n = UBound(iStack)
    For i = 1 To n
        If iStack(i) = k Then
            For j = i To n - 1
                iStack(j) = iStack(j + 1)
            Next j
            ReDim Preserve iStack(0 To n - 1)
            Exit For
        End If
    Next i
End Sub
'=============================================================

'=== Queue Operations
Private Sub Enqueue(iQueue() As Long, k As Long)
Dim n As Long, i As Long
n = UBound(iQueue)
ReDim Preserve iQueue(0 To n + 1)
For i = n + 1 To 2 Step -1
    iQueue(i) = iQueue(i - 1)
Next i
iQueue(1) = k
End Sub
'=============================================================


'Cut tree at branch that has longest life, but limit number of clusters
'to max_cluster if a positive integer is supplied.
'Output: Assign_Cluster(1:N), cluster membership index of each sample
'        n_cluster, number of clusters discovered.
Function Assign_Cluster(n_cluster As Long, Optional max_cluster As Long = -1) As Long()
Dim k As Long, m As Long, u As Long, v As Long, cut_pt As Long
Dim tmp_x As Double, tmp_y As Double, max_life As Double
Dim n_cluster_max As Long, cluster_idx() As Long
    If max_cluster > 0 Then
        n_cluster_max = max_cluster
    Else
        n_cluster_max = pn_raw
    End If
    m = 1
    n_cluster = 1
    max_life = -1
    cut_pt = -1
    For k = pn_raw - 1 To 1 Step -1
        m = m + 1: tmp_x = pZ_height(k): tmp_y = pZ_height(k)
        u = pZ(k, 2): If u > pn_raw Then tmp_x = tmp_x - pZ_height(u - pn_raw)
        v = pZ(k, 3): If v > pn_raw Then tmp_y = tmp_y - pZ_height(v - pn_raw)
        If tmp_y < tmp_x Then tmp_x = tmp_y
        If tmp_x > max_life Then
            max_life = tmp_x
            n_cluster = m
            cut_pt = pn_raw + k
        End If
        If m >= n_cluster_max Then Exit For
    Next k
    Call Find_cluster_idx(cluster_idx, cut_pt)
    Call Cluster_Idx_Rearrange(cluster_idx, n_cluster)
    Assign_Cluster = cluster_idx
End Function

'Cut tree to give a specific number of clusters
Function Assign_Cluster_fix_cluster_num(n_cluster As Long) As Long()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, cut_pt As Long
Dim cluster_idx() As Long
    If n_cluster > pn_raw Then
        Debug.Print "Assign_Cluster_fix_cluster_num: number of clusters larger than sample size."
        Exit Function
    ElseIf n_cluster = 1 Then
        ReDim cluster_idx(1 To pn_raw)
        For i = 1 To pn_raw
            cluster_idx(i) = 1
        Next i
        Assign_Cluster_fix_cluster_num = cluster_idx
        Exit Function
    End If
    cut_pt = pZ(pn_raw - n_cluster + 1, 1)
    Call Find_cluster_idx(cluster_idx, cut_pt)
    Call Cluster_Idx_Rearrange(cluster_idx, n_cluster)
    Assign_Cluster_fix_cluster_num = cluster_idx
End Function


'Returns cluster index base on cut point
Private Sub Find_cluster_idx(cluster_idx() As Long, cut_pt As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, u As Long, v As Long
Dim tmp_x As Double, tmp_y As Double, iArr() As Long, new_roots() As Long
    n = 0
    ReDim new_roots(1 To 1)
    For k = cut_pt - pn_raw To pn_raw - 1
        If pZ(k, 2) < cut_pt Then
            n = n + 1
            ReDim Preserve new_roots(1 To n)
            new_roots(n) = pZ(k, 2)
        End If
        If pZ(k, 3) < cut_pt Then
            n = n + 1
            ReDim Preserve new_roots(1 To n)
            new_roots(n) = pZ(k, 3)
        End If
    Next k
    ReDim cluster_idx(1 To pn_raw)
    For k = 1 To UBound(new_roots)
        If new_roots(k) > pn_raw Then
            Call Find_Children(new_roots(k), iArr)
            For i = 1 To UBound(iArr)
                cluster_idx(iArr(i)) = k
            Next i
        Else
            cluster_idx(new_roots(k)) = k
        End If
    Next k
End Sub

'Rename cluster so largest cluster is called #1, second largest is #2 etc.
Private Sub Cluster_Idx_Rearrange(cluster_idx() As Long, n_cluster As Long)
Dim i As Long, j As Long, k As Long, n As Long
Dim cluster_size() As Long, sort_idx() As Long, iArr() As Long
    ReDim cluster_size(1 To n_cluster)
    For i = 1 To pn_raw
        cluster_size(cluster_idx(i)) = cluster_size(cluster_idx(i)) + 1
    Next i
    Call modMath.Sort_Quick_A(cluster_size, 1, n_cluster, sort_idx)
    iArr = cluster_idx
    For i = 1 To pn_raw
        For j = 1 To n_cluster
            If iArr(i) = sort_idx(j) Then
                cluster_idx(i) = n_cluster - j + 1
                Exit For
            End If
        Next j
    Next i
End Sub

'Return all leaves under node u
Private Sub Find_Children(u As Long, child_list() As Long, Optional first_run As Long = 1)
Dim i As Long, j As Long, k As Long, n As Long, child_L As Long, child_R As Long
    If first_run = 1 Then
        ReDim child_list(0 To 0)
        If u <= pn_raw Then Exit Sub
    End If
    If u <= pn_raw Then
        n = UBound(child_list) + 1
        If n = 1 Then
            ReDim child_list(1 To n)
        Else
            ReDim Preserve child_list(1 To n)
        End If
        child_list(n) = u
        Exit Sub
    End If
    child_L = pZ(u - pn_raw, 2)
    child_R = pZ(u - pn_raw, 3)
    Call Find_Children(child_L, child_list, 0)
    Call Find_Children(child_R, child_list, 0)
End Sub



Function Print_Tree_as_variant(Optional skew As Double = 1, Optional circular As Boolean = False, _
        Optional cut_level As Double = 1, Optional print_mask As Boolean = False) As Variant
Dim i As Long, j As Long, u As Long, v As Long, w As Long, n_max As Long
Dim vArr As Variant, tree_height As Double, na_err As Variant

    na_err = VBA.CVErr(xlErrNA)

    Call get_xy_plot(skew, circular)   'get (x,y)-coordinates for visualization
    
    'Truncate tree at n_max
    n_max = pn_raw - 1
    tree_height = pZ_height(pn_raw - 1)
    If cut_level < 1 Then
        tree_height = tree_height * cut_level
        If tree_height < pZ_height(1) Then
            n_max = 0
        Else
            For i = pn_raw - 1 To 1 Step -1
                If pZ_height(i) <= tree_height Then
                    n_max = i
                    Exit For
                End If
            Next i
        End If
    End If

    i = 5 * n_max - 1
    If i < pn_raw Then i = pn_raw
    If print_mask = False Then
        ReDim vArr(1 To i, 0 To 6)
    Else
        ReDim vArr(1 To i, 0 To 8)
    End If
    
    'Print sweep line
    vArr(1, 5) = tree_height + 1
    vArr(2, 5) = tree_height + 1
    vArr(1, 6) = 0
    vArr(2, 6) = pn_raw
    For i = 3 To UBound(vArr, 1)
        vArr(i, 5) = na_err
        vArr(i, 6) = na_err
    Next i
    
    'Print leaves
    For i = 1 To pn_raw
        j = pleaf_order(i)
        vArr(j, 0) = plabel(i)
        vArr(j, 1) = pxy_plot(i, 1)
        vArr(j, 2) = pxy_plot(i, 2)
    Next i
    For i = pn_raw + 1 To UBound(vArr, 1)
        vArr(i, 0) = na_err
        vArr(i, 1) = na_err
        vArr(i, 2) = na_err
    Next i
    
    'Print tree
    j = 1
    For i = 1 To n_max
        w = pn_raw + i
        u = pZ(i, 2)
        v = pZ(i, 3)
        vArr(j, 3) = pxy_plot(u, 1)
        vArr(j, 4) = pxy_plot(u, 2)
        vArr(j + 1, 3) = pxy_plot(w, 1)
        vArr(j + 1, 4) = pxy_plot(u, 2)
        vArr(j + 2, 3) = pxy_plot(w, 1)
        vArr(j + 2, 4) = pxy_plot(v, 2)
        vArr(j + 3, 3) = pxy_plot(v, 1)
        vArr(j + 3, 4) = pxy_plot(v, 2)
'        If i < n_max Then
'            vArr(j + 4, 3) = VBA.CVErr(xlErrNA)
'            vArr(j + 4, 4) = VBA.CVErr(xlErrNA)
'        End If
        j = j + 5
    Next i
    
    'Print masking line
    If print_mask = True Then
        j = -1
        For i = 1 To n_max
            u = pZ(i, 2)
            v = pZ(i, 3)
            If i > 1 Then
                vArr(j, 7) = pxy_plot(u, 1)
                vArr(j, 8) = pxy_plot(u, 2)
            End If
            vArr(j + 2, 7) = pxy_plot(v, 1)
            vArr(j + 2, 8) = pxy_plot(v, 2)
            j = j + 4
        Next i
    End If

    Print_Tree_as_variant = vArr
    Erase vArr
End Function

