VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cqtree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=================================
'Quad-tree/Oct-tree data structure
'=================================

Private pn_dimension As Long    'dimension of problem, either 2 or 3
Private pcenter() As Double     'coordinates of cell's mid-point
Private pCoM() As Double        'coordinates of cell's center of mass
Private pwidth As Double        'width of cell which is designed to be square/cube
Private pSize As Long, pMass As Double  'size is number of points in cell, mass is total mass of these points
Private pchild() As cqtree
Private pisRoot As Long     'equals 1 if cell is root cell
Private pLeafIdx As Long    'if cell contains only one point then this is the integer pointer to that point
Private pdepth As Long      'depth of cell, root cell has depth 0
Private pisLeaf As Long     'equals 1 if cell is a leaf, 0 otherwise

Public Property Get center() As Double()
    center = pcenter
End Property

Public Property Get width() As Double
    width = pwidth
End Property

Public Property Get Size() As Long
    Size = pSize
End Property

Public Property Get Mass() As Double
    Mass = pMass
End Property

Public Property Get n_dimension() As Long
    n_dimension = pn_dimension
End Property

Sub Reset()
    Dim i As Long
    For i = 1 To UBound(pchild)
        If Not pchild(i) Is Nothing Then Call pchild(i).Reset
    Next i
    Erase pcenter, pCoM, pchild
    pSize = 0
    pMass = 0
End Sub

Sub Init(x_center() As Double, width As Double, CoM() As Double, _
            Size As Long, depth As Long, max_depth As Long, Optional LeafIdx As Long, _
            Optional Mass As Double = 1)
    pn_dimension = UBound(CoM, 1)
    pcenter = x_center
    pCoM = CoM
    pSize = Size
    pMass = Mass
    pdepth = depth
    pwidth = width
    ReDim pchild(-1 To -1)
    If pdepth = max_depth Then
        pisLeaf = 1
    Else
        pisLeaf = 0
    End If
    If pSize = 1 Then
        pLeafIdx = LeafIdx
        pisLeaf = 1
    Else
        pLeafIdx = -1
    End If
End Sub


Sub Build(x() As Double, Optional max_depth As Long = 8, _
            Optional PtWgts As Variant, Optional isWgt As Boolean = False)
Dim i As Long, j As Long, k As Long, n_raw As Long
Dim max_x As Double, min_x As Double
Dim width As Double, mid As Double, total_wgt As Double
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, INFINITY As Double
Dim CoM() As Double, x_pos() As Double, x_mid() As Double
Dim pts As Collection
    INFINITY = Exp(70)
    pn_dimension = UBound(x, 2)
    n_raw = UBound(x, 1)

    '=== Find bounding box of all points
    width = -INFINITY
    min_x = INFINITY: max_x = -INFINITY
    ReDim x_mid(1 To pn_dimension)
    For j = 1 To pn_dimension
        
        For i = 1 To n_raw
            If x(i, j) > max_x Then max_x = x(i, j)
            If x(i, j) < min_x Then min_x = x(i, j)
        Next i
'        x_mid(j) = (max_x + min_x) / 2
'        If (max_x - min_x) > width Then
'            width = max_x - min_x
'        End If
    Next j
    
    width = max_x - min_x
    For j = 1 To pn_dimension
        x_mid(j) = (max_x + min_x) / 2
    Next j
    
    '=== Assign each point to a quad tree node
    tmp_x = 0: tmp_y = 0
    Set pts = New Collection
    For i = 1 To n_raw
        pts.Add i
    Next i

    '=== Total mass and center of mass
    If isWgt = False Then
        CoM = CoM_Calc(x, pts, , False, total_wgt)
    Else
        CoM = CoM_Calc(x, pts, PtWgts, True, total_wgt)
    End If

    'max_depth = Int(Log(n_raw) / Log(4)) * 2
    '=== Recursively split points into cells
    pisRoot = 1
    If isWgt = False Then
        Call Init(x_mid, width, CoM, n_raw, 0, max_depth, , total_wgt)
        Call Split(x, pts, max_depth, , False)
    Else
        Call Init(x_mid, width, CoM, n_raw, 0, max_depth, , total_wgt)
        Call Split(x, pts, max_depth, PtWgts, True)
    End If
End Sub



Sub Split(x() As Double, pts As Collection, max_depth As Long, Optional PtWgts As Variant, Optional isWgt As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_raw As Long, n_child As Long
Dim idx As Long
Dim tmp_x As Double, tmp_y As Double
Dim new_node As cqtree
Dim sizes() As Long, masses() As Double, CoM() As Double, total_mass As Double
Dim centers() As Double, CoM_i() As Double, centers_i() As Double
Dim q As Collection, qs As Collection, qi As Collection

    n_child = 2 ^ (pn_dimension)
    ReDim pchild(1 To n_child)
    ReDim centers(1 To n_child, 1 To pn_dimension)
    ReDim sizes(1 To n_child)

    Set qs = New Collection
    For i = 1 To n_child
        Set qi = New Collection
        qs.Add qi
    Next i

    '=== Define new splitting planes
    tmp_x = pcenter(1) - pwidth / 4: tmp_y = pcenter(1) + pwidth / 4
    For i = 1 To n_child Step 2
        centers(i, 1) = tmp_x
        centers(i + 1, 1) = tmp_y
    Next i
    tmp_x = pcenter(2) + pwidth / 4: tmp_y = pcenter(2) - pwidth / 4
    For i = 1 To n_child Step 4
        centers(i, 2) = tmp_x
        centers(i + 1, 2) = tmp_x
        centers(i + 2, 2) = tmp_y
        centers(i + 3, 2) = tmp_y
    Next i
    If pn_dimension = 3 Then
        tmp_x = pcenter(3) - pwidth / 4: tmp_y = pcenter(3) + pwidth / 4
        For i = 1 To 4
            centers(i, 3) = tmp_x
            centers(i + 4, 3) = tmp_y
        Next i
    End If

    '=== Assign points to children
    Do While pts.count > 0
        idx = pts(pts.count)
        pts.Remove (pts.count)
        If pn_dimension = 2 Then
            If x(idx, 1) < pcenter(1) Then
                If x(idx, 2) < pcenter(2) Then
                    n = 3
                Else
                    n = 1
                End If
            Else
                If x(idx, 2) < pcenter(2) Then
                    n = 4
                Else
                    n = 2
                End If
            End If
        Else
            If x(idx, 1) < pcenter(1) Then
                If x(idx, 2) < pcenter(2) Then
                    If x(idx, 3) < pcenter(3) Then
                        n = 3
                    Else
                        n = 7
                    End If
                Else
                    If x(idx, 3) < pcenter(3) Then
                        n = 1
                    Else
                        n = 5
                    End If
                End If
            Else
                If x(idx, 2) < pcenter(2) Then
                    If x(idx, 3) < pcenter(3) Then
                        n = 4
                    Else
                        n = 8
                    End If
                Else
                    If x(idx, 3) < pcenter(3) Then
                        n = 2
                    Else
                        n = 6
                    End If
                End If
            End If
        End If
        Set qi = qs(n)
        qi.Add idx
        sizes(n) = sizes(n) + 1
    Loop

    ReDim CoM_i(1 To pn_dimension)
    ReDim centers_i(1 To pn_dimension)
    For n = 1 To n_child
        Set q = qs(n)
        If sizes(n) > 0 Then
            
            For j = 1 To pn_dimension
                centers_i(j) = centers(n, j)
            Next j
            If isWgt = False Then
                CoM = CoM_Calc(x, q, , , total_mass)
            Else
                CoM = CoM_Calc(x, q, PtWgts, True, total_mass)
            End If
            Set new_node = New cqtree
            i = -1: If sizes(n) = 1 Then i = q(1)
            Call new_node.Init(centers_i, pwidth / 2, CoM, sizes(n), pdepth + 1, max_depth, i, total_mass)

            Set pchild(n) = new_node
            'Continue splitting if cell has more than one point and maximum deppth is not reached yet
            If sizes(n) > 1 And (pdepth + 1) < max_depth Then
                If isWgt = False Then
                    Call pchild(n).Split(x, q, max_depth)
                Else
                    Call pchild(n).Split(x, q, max_depth, PtWgts, True)
                End If
            End If
        Else
            'Cell is empty, stop splitting
            Set pchild(n) = Nothing
        End If
    Next n
    Erase CoM, centers, centers_i
End Sub


Private Function CoM_Calc(x() As Double, pts As Collection, _
            Optional PtWgts As Variant, Optional isWgt As Boolean = False, _
            Optional total_mass As Double) As Double()
Dim i As Long, ii As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim CoM() As Double
    n = pts.count
    n_dimension = UBound(x, 2)
    ReDim CoM(1 To n_dimension)
    If isWgt = False Then
        total_mass = n
        For ii = 1 To n
            i = pts(ii)
            For j = 1 To n_dimension
                CoM(j) = CoM(j) + x(i, j)
            Next j
        Next ii
    Else
        total_mass = 0
        For ii = 1 To n
            i = pts(ii)
            total_mass = total_mass + PtWgts(i)
            For j = 1 To n_dimension
                CoM(j) = CoM(j) + x(i, j) * PtWgts(i)
            Next j
        Next ii
    End If
    For j = 1 To n_dimension
        CoM(j) = CoM(j) / total_mass
    Next j
    CoM_Calc = CoM
End Function

Sub Print_Tree(vArr As Variant)
Dim i As Long, j As Long, n As Long, m As Long
    m = 2 + 2 * pn_dimension
    If pisRoot = 1 Then ReDim vArr(1 To m, 0 To 0)
    If pSize > 0 Then
        n = UBound(vArr, 2) + 1
        If n = 1 Then
            ReDim vArr(1 To m, 1 To n)
        Else
            ReDim Preserve vArr(1 To m, 1 To n)
        End If
        vArr(1, n) = pSize
        vArr(2, n) = pwidth
        For j = 1 To pn_dimension
            vArr(2 + j, n) = pcenter(j)
        Next j
        For j = 1 To pn_dimension
            vArr(2 + pn_dimension + j, n) = pCoM(j)
        Next j
    End If
    For i = 1 To UBound(pchild)
        If Not pchild(i) Is Nothing Then Call pchild(i).Print_Tree(vArr)
    Next i
End Sub


Sub Print_Grid(vArr As Variant)
Dim i As Long, j As Long, k As Long, n As Long
Dim width As Double
Dim xArr As Variant
    Call Print_Tree(xArr)
    n = UBound(xArr, 2)
    If pn_dimension = 2 Then
        ReDim vArr(1 To n * 6, 1 To 2)
        k = 1
        For i = 1 To n
            width = xArr(2, i) / 2
            vArr(k, 1) = xArr(3, i) - width
            vArr(k, 2) = xArr(4, i) - width
            vArr(k + 1, 1) = xArr(3, i) + width
            vArr(k + 1, 2) = xArr(4, i) - width
            vArr(k + 2, 1) = xArr(3, i) + width
            vArr(k + 2, 2) = xArr(4, i) + width
            vArr(k + 3, 1) = xArr(3, i) - width
            vArr(k + 3, 2) = xArr(4, i) + width
            vArr(k + 4, 1) = xArr(3, i) - width
            vArr(k + 4, 2) = xArr(4, i) - width
            k = k + 6
        Next i
    ElseIf pn_dimension = 3 Then
        ReDim vArr(1 To n * 20, 1 To 3)
        k = 1
        For i = 1 To n
            width = xArr(2, i) / 2
            vArr(k, 1) = xArr(3, i) - width
            vArr(k, 2) = xArr(4, i) - width
            vArr(k, 3) = xArr(5, i) - width
            vArr(k + 1, 1) = xArr(3, i) + width
            vArr(k + 1, 2) = xArr(4, i) - width
            vArr(k + 1, 3) = xArr(5, i) - width
            vArr(k + 2, 1) = xArr(3, i) + width
            vArr(k + 2, 2) = xArr(4, i) + width
            vArr(k + 2, 3) = xArr(5, i) - width
            vArr(k + 3, 1) = xArr(3, i) - width
            vArr(k + 3, 2) = xArr(4, i) + width
            vArr(k + 3, 3) = xArr(5, i) - width
            vArr(k + 4, 1) = xArr(3, i) - width
            vArr(k + 4, 2) = xArr(4, i) - width
            vArr(k + 4, 3) = xArr(5, i) - width
            vArr(k + 5, 1) = xArr(3, i) - width
            vArr(k + 5, 2) = xArr(4, i) - width
            vArr(k + 5, 3) = xArr(5, i) + width
            vArr(k + 6, 1) = xArr(3, i) + width
            vArr(k + 6, 2) = xArr(4, i) - width
            vArr(k + 6, 3) = xArr(5, i) + width
            vArr(k + 7, 1) = xArr(3, i) + width
            vArr(k + 7, 2) = xArr(4, i) + width
            vArr(k + 7, 3) = xArr(5, i) + width
            vArr(k + 8, 1) = xArr(3, i) - width
            vArr(k + 8, 2) = xArr(4, i) + width
            vArr(k + 8, 3) = xArr(5, i) + width
            vArr(k + 9, 1) = xArr(3, i) - width
            vArr(k + 9, 2) = xArr(4, i) - width
            vArr(k + 9, 3) = xArr(5, i) + width
            
            vArr(k + 11, 1) = xArr(3, i) - width
            vArr(k + 11, 2) = xArr(4, i) + width
            vArr(k + 11, 3) = xArr(5, i) - width
            vArr(k + 12, 1) = xArr(3, i) - width
            vArr(k + 12, 2) = xArr(4, i) + width
            vArr(k + 12, 3) = xArr(5, i) + width
    
            vArr(k + 14, 1) = xArr(3, i) + width
            vArr(k + 14, 2) = xArr(4, i) + width
            vArr(k + 14, 3) = xArr(5, i) - width
            vArr(k + 15, 1) = xArr(3, i) + width
            vArr(k + 15, 2) = xArr(4, i) + width
            vArr(k + 15, 3) = xArr(5, i) + width

            vArr(k + 17, 1) = xArr(3, i) + width
            vArr(k + 17, 2) = xArr(4, i) - width
            vArr(k + 17, 3) = xArr(5, i) - width
            vArr(k + 18, 1) = xArr(3, i) + width
            vArr(k + 18, 2) = xArr(4, i) - width
            vArr(k + 18, 3) = xArr(5, i) + width

            k = k + 20
        Next i
    End If
End Sub



'=== Use Barnes-Hut Algorithm to simulate inverse-p repulsive force
'Input: x(1:N, 1:n_dimension), 2D/3D positions of N data points
'       theta, the larger it is the coraser the approximation goes
'Output: NetForce(1:N, 1:n_dimension), components of net force on each of N data points
Function NetForce(x() As Double, Optional force_constant As Double = 1, Optional theta As Double = 0.75, _
            Optional PtWgts As Variant, Optional isWgt As Boolean = False, Optional c_power As Double = 2) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim force_vec() As Double, force() As Double, x_pos() As Double
Dim quadtree As cqtree

    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    Set quadtree = New cqtree
    If isWgt = False Then
        Call quadtree.Build(x)
    Else
        Call quadtree.Build(x, , PtWgts, True)
    End If
            
    ReDim x_pos(1 To n_dimension)
    ReDim force(1 To n, 1 To n_dimension)
    For i = 1 To n
        ReDim force_vec(1 To n_dimension)
        For j = 1 To n_dimension
            x_pos(j) = x(i, j)
        Next j
        Call quadtree.NetForce_Traverse(i, x_pos, force_vec, theta, c_power)
        If isWgt = False Then
            For j = 1 To n_dimension
                force(i, j) = force_constant * force_vec(j)
            Next j
        Else
            For j = 1 To n_dimension
                force(i, j) = force_constant * force_vec(j) * PtWgts(i)
            Next j
        End If
    Next i
    Call quadtree.Reset
    Set quadtree = Nothing
    NetForce = force
    Erase force, force_vec, x_pos
End Function


'Calculate force on single body labelled at idx
Function NetForce_Single(idx As Long, x() As Double, Optional force_constant As Double = 1, Optional theta As Double = 0.75, _
            Optional PtWgts As Variant, Optional isWgt As Boolean = False, Optional c_power As Double = 2) As Double()
Dim j As Long, n_dimension As Long
Dim force() As Double, x_pos() As Double
    n_dimension = UBound(x, 2)
    ReDim x_pos(1 To n_dimension)
    ReDim force(1 To n_dimension)
    For j = 1 To n_dimension
        x_pos(j) = x(idx, j)
    Next j
    Call NetForce_Traverse(idx, x_pos, force, theta, c_power)
    If isWgt = False Then
        For j = 1 To n_dimension
            force(j) = force_constant * force(j)
        Next j
    Else
        For j = 1 To n_dimension
            force(j) = force_constant * force(j) * PtWgts(idx)
        Next j
    End If
    NetForce_Single = force
    Erase force, x_pos
End Function

Sub NetForce_Traverse(ptIdx As Long, x_pos() As Double, force_vec() As Double, theta As Double, c_power As Double)
Dim k As Long
Dim tmp_x As Double, d As Double
    If ptIdx = pLeafIdx Then Exit Sub 'reach itself
    'Distance from cell's center of mass
    d = 0
    For k = 1 To pn_dimension
        d = d + (x_pos(k) - pCoM(k)) ^ 2
    Next k
    d = Sqr(d)
    'Cell is far enough from data or is a leaf
    If d > 0 And (pisLeaf = 1 Or width < (d * theta)) Then
        tmp_x = pMass / (d ^ (c_power + 1))
        For k = 1 To n_dimension
            force_vec(k) = force_vec(k) + (x_pos(k) - pCoM(k)) * tmp_x
        Next k
    'Overlap with another pt, randomly break them
    ElseIf d = 0 And pisLeaf = 1 Then
        For k = 1 To n_dimension
            force_vec(k) = force_vec(k) + Rnd()
        Next k
    Else
        For k = 1 To UBound(pchild)
            If Not pchild(k) Is Nothing Then
                Call pchild(k).NetForce_Traverse(ptIdx, x_pos, force_vec, theta, c_power)
            End If
        Next k
    End If
End Sub

'=== Use Barnes-Hut Algorithm to simulate repulsive terms  in t-SNE
'Input: x(1 to N, 1 to 2), 2D positions of N data points
'Output: tSNE_Force(1 to N, 1 to 2), x & y components of net force on each of N data points
'Output: z, normalization constant
Function tSNE_Force(x() As Double, z As Double, Optional theta As Double = 0.5) As Double()
Dim i As Long, j As Long, k As Long, n_raw As Long
Dim fx As Double, Fy As Double
Dim force() As Double
Dim quadtree As cqtree
n_raw = UBound(x, 1)
Set quadtree = New cqtree
Call quadtree.Build(x)
ReDim force(1 To n_raw, 1 To 2)
z = 0
For i = 1 To n_raw
    fx = 0
    Fy = 0
    Call quadtree.tSNE_Force_Traverse(i, x(i, 1), x(i, 2), fx, Fy, z, theta)
    force(i, 1) = fx
    force(i, 2) = Fy
Next i
Call quadtree.Reset
Set quadtree = Nothing
tSNE_Force = force
Erase force
End Function

Sub tSNE_Force_Traverse(ptIdx As Long, x As Double, y As Double, fx As Double, Fy As Double, z As Double, theta As Double)
Dim n As Long
Dim tmp_x As Double, d As Double

If ptIdx = pLeafIdx Then Exit Sub

d = (x - pCoM(1)) ^ 2 + (y - pCoM(2)) ^ 2

If pisLeaf = 1 Then
    tmp_x = 1# / (1 + d)
    z = z + pSize * tmp_x
    fx = fx + (x - pCoM(1)) * pSize * (tmp_x ^ 2)
    Fy = Fy + (y - pCoM(2)) * pSize * (tmp_x ^ 2)
    Exit Sub
End If

If (pwidth * pwidth / theta) < d Then
    tmp_x = 1# / (1 + d)
    z = z + pSize * tmp_x
    fx = fx + (x - pCoM(1)) * pSize * (tmp_x ^ 2)
    Fy = Fy + (y - pCoM(2)) * pSize * (tmp_x ^ 2)
Else
    For n = 1 To UBound(pchild)
        If Not pchild(n) Is Nothing Then Call pchild(n).tSNE_Force_Traverse(ptIdx, x, y, fx, Fy, z, theta)
    Next n
End If
End Sub




Sub Traverse(x As Double, y As Double, theta As Double, vArr As Variant)
Dim n As Long
Dim tmp_x As Double, d As Double
d = Sqr((x - pCoM(1)) ^ 2 + (y - pCoM(2)) ^ 2)

If pSize = 1 Then
    If d > 0 Then
        n = UBound(vArr, 2) + 1
        ReDim Preserve vArr(1 To 4, 0 To n)
        vArr(1, n) = pCoM(1)
        vArr(2, n) = pCoM(2)
        vArr(3, n) = d
        vArr(4, n) = pSize
        Exit Sub
    End If
End If

If d > 0 Then
    If (pwidth / d) < theta Then
        n = UBound(vArr, 2) + 1
        ReDim Preserve vArr(1 To 4, 0 To n)
        vArr(1, n) = pCoM(1)
        vArr(2, n) = pCoM(2)
        vArr(3, n) = d
        vArr(4, n) = pSize
    Else
        For n = 1 To UBound(pchild)
            If Not pchild(n) Is Nothing Then Call pchild(n).Traverse(x, y, theta, vArr)
        Next n
    End If
End If
End Sub
