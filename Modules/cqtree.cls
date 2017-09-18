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
'Quad-tree data structure
'=================================
'Requires: cqtree_point
'=================================

Private pcenter() As Double
Private pwidth As Double
Private pSize As Long
Private pchild() As cqtree
Private pCoM() As Double
Private pisRoot As Long
Private pLeafIdx As Long
Private pdepth As Long
Private pisLeaf As Long

Public Property Get center() As Double()
center = pcenter
End Property

Public Property Get width() As Double
width = pwidth
End Property

Public Property Get Size() As Long
Size = pSize
End Property


Sub Reset()
Dim i As Long
For i = 1 To UBound(pchild)
    If Not pchild(i) Is Nothing Then Call pchild(i).Reset
Next i
Erase pcenter
Erase pCoM
Erase pchild
End Sub

Sub Init(x As Double, y As Double, width As Double, CoM_x As Double, CoM_y As Double, Size As Long, depth As Long, max_depth As Long, Optional LeafIdx As Long)
    ReDim pcenter(1 To 2)
    ReDim pCoM(1 To 2)
    ReDim pchild(-1 To -1)
    pcenter(1) = x
    pcenter(2) = y
    pwidth = width
    pCoM(1) = CoM_x
    pCoM(2) = CoM_y
    pSize = Size
    pdepth = depth
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


Sub Build(x() As Double, Optional max_depth As Long = 8)
Dim i As Long, j As Long, k As Long, n_raw As Long
Dim max_x As Double, min_x As Double, max_y As Double, min_y As Double
Dim width As Double, mid As Double
Dim tmp_x As Double, tmp_y As Double, INFINITY As Double
Dim xy As cqtree_point
Dim pts As Collection

INFINITY = Exp(70)

n_raw = UBound(x, 1)
max_x = -INFINITY
max_y = -INFINITY
min_x = INFINITY
min_y = INFINITY
For i = 1 To n_raw
    If x(i, 1) > max_x Then max_x = x(i, 1)
    If x(i, 1) < min_x Then min_x = x(i, 1)
    If x(i, 2) > max_y Then max_y = x(i, 2)
    If x(i, 2) < min_y Then min_y = x(i, 2)
Next i

width = max_x - min_x
mid = (max_x + min_x) / 2
If (max_y - min_y) > width Then
    width = max_y - min_y
    mid = (max_y + min_y) / 2
End If

tmp_x = 0
tmp_y = 0
Set pts = New Collection
For i = 1 To n_raw
    Set xy = New cqtree_point
    Call xy.Init(x(i, 1), x(i, 2), i)
    pts.Add xy
    tmp_x = tmp_x + x(i, 1)
    tmp_y = tmp_y + x(i, 2)
Next i

'max_depth = Int(Log(n_raw) / Log(4)) * 2
pisRoot = 1
Call Me.Init(mid, mid, width, tmp_x / n_raw, tmp_y / n_raw, n_raw, 0, max_depth)
Call Me.Split(pts, max_depth)

End Sub


Sub Split(pts As Collection, max_depth As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_raw As Long
Dim xy As cqtree_point
Dim new_node As cqtree
Dim sizes() As Long
Dim centers() As Double
Dim q As Collection, q1 As Collection, q2 As Collection, q3 As Collection, q4 As Collection
Dim CoM() As Double

ReDim pchild(1 To 4)
ReDim centers(1 To 4, 1 To 2)
ReDim sizes(1 To 4)
ReDim CoM(1 To 4, 1 To 2)

Set q1 = New Collection
Set q2 = New Collection
Set q3 = New Collection
Set q4 = New Collection

centers(1, 1) = pcenter(1) - pwidth / 4
centers(2, 1) = pcenter(1) + pwidth / 4
centers(3, 1) = pcenter(1) - pwidth / 4
centers(4, 1) = pcenter(1) + pwidth / 4

centers(1, 2) = pcenter(2) + pwidth / 4
centers(2, 2) = pcenter(2) + pwidth / 4
centers(3, 2) = pcenter(2) - pwidth / 4
centers(4, 2) = pcenter(2) - pwidth / 4

Do While pts.count > 0
    Set xy = pts(pts.count)
    pts.Remove (pts.count)
    If xy.xy(1) < pcenter(1) Then
        If xy.xy(2) < pcenter(2) Then
            n = 3
            q3.Add xy
        Else
            n = 1
            q1.Add xy
        End If
    Else
        If xy.xy(2) < pcenter(2) Then
            n = 4
            q4.Add xy
        Else
            n = 2
            q2.Add xy
        End If
    End If
    sizes(n) = sizes(n) + 1
    CoM(n, 1) = CoM(n, 1) + xy.xy(1)
    CoM(n, 2) = CoM(n, 2) + xy.xy(2)
Loop

For n = 1 To 4
    If n = 1 Then
        Set q = q1
    ElseIf n = 2 Then
         Set q = q2
    ElseIf n = 3 Then
         Set q = q3
    ElseIf n = 4 Then
         Set q = q4
    End If
    If sizes(n) > 0 Then
        CoM(n, 1) = CoM(n, 1) / sizes(n)
        CoM(n, 2) = CoM(n, 2) / sizes(n)
        Set new_node = New cqtree
        If sizes(n) = 1 Then
            Call new_node.Init(centers(n, 1), centers(n, 2), pwidth / 2, CoM(n, 1), CoM(n, 2), sizes(n), pdepth + 1, max_depth, q(1).index)
        Else
            Call new_node.Init(centers(n, 1), centers(n, 2), pwidth / 2, CoM(n, 1), CoM(n, 2), sizes(n), pdepth + 1, max_depth)
        End If
        Set pchild(n) = new_node
        If sizes(n) > 1 And (pdepth + 1) < max_depth Then Call pchild(n).Split(q, max_depth)
    Else
        Set pchild(n) = Nothing
    End If
Next n
Erase CoM, centers
End Sub


Sub Print_Tree(vArr As Variant)
Dim i As Long, n As Long
If pisRoot = 1 Then ReDim vArr(1 To 6, 0 To 0)
If pSize > 0 Then
    n = UBound(vArr, 2) + 1
    If n = 1 Then
        ReDim vArr(1 To 6, 1 To n)
    Else
        ReDim Preserve vArr(1 To 6, 1 To n)
    End If
    vArr(1, n) = pcenter(1)
    vArr(2, n) = pcenter(2)
    vArr(3, n) = pwidth
    vArr(4, n) = pSize
    vArr(5, n) = pCoM(1)
    vArr(6, n) = pCoM(2)
End If
For i = 1 To UBound(pchild)
    If Not pchild(i) Is Nothing Then Call pchild(i).Print_Tree(vArr)
Next i
End Sub


Sub Print_Grid(vArr As Variant)
Dim i As Long, j As Long, k As Long, n As Long
Dim xArr As Variant
Call Me.Print_Tree(xArr)
n = UBound(xArr, 2)
ReDim vArr(1 To n * 6, 1 To 2)
k = 1
For i = 1 To n
    vArr(k, 1) = xArr(1, i) - xArr(3, i) / 2
    vArr(k, 2) = xArr(2, i) - xArr(3, i) / 2
    vArr(k + 1, 1) = xArr(1, i) + xArr(3, i) / 2
    vArr(k + 1, 2) = xArr(2, i) - xArr(3, i) / 2
    vArr(k + 2, 1) = xArr(1, i) + xArr(3, i) / 2
    vArr(k + 2, 2) = xArr(2, i) + xArr(3, i) / 2
    vArr(k + 3, 1) = xArr(1, i) - xArr(3, i) / 2
    vArr(k + 3, 2) = xArr(2, i) + xArr(3, i) / 2
    vArr(k + 4, 1) = xArr(1, i) - xArr(3, i) / 2
    vArr(k + 4, 2) = xArr(2, i) - xArr(3, i) / 2
    k = k + 6
Next i
End Sub



'=== Use Barnes-Hut Algorithm to simulate inverse square repulsive force
'Input: x(1 to N, 1 to 2), 2D positions of N data points
'Output: NetForce(1 to N, 1 to 2), x & y components of net force on each of N data points
Function NetForce(x() As Double, Optional force_constant As Double = 1, Optional theta As Double = 0.75) As Double()
Dim i As Long, j As Long, k As Long, n_raw As Long
Dim fx As Double, Fy As Double
Dim force() As Double
Dim quadtree As cqtree
n_raw = UBound(x, 1)
Set quadtree = New cqtree
Call quadtree.Build(x)
ReDim force(1 To n_raw, 1 To 2)
For i = 1 To n_raw
    fx = 0
    Fy = 0
    Call quadtree.NetForce_Traverse(i, x(i, 1), x(i, 2), fx, Fy, theta)
    force(i, 1) = force_constant * fx
    force(i, 2) = force_constant * Fy
Next i
Call quadtree.Reset
Set quadtree = Nothing
NetForce = force
Erase force
End Function


Sub NetForce_Traverse(ptIdx As Long, x As Double, y As Double, fx As Double, Fy As Double, theta As Double)
Dim n As Long
Dim tmp_x As Double, d As Double

If ptIdx = pLeafIdx Then Exit Sub 'reach itself

d = Sqr((x - pCoM(1)) ^ 2 + (y - pCoM(2)) ^ 2)

If pisLeaf = 1 And d > 0 Then 'reach a leaf
    tmp_x = pSize / (d ^ 3)
    fx = fx + (x - pCoM(1)) * tmp_x
    Fy = Fy + (y - pCoM(2)) * tmp_x
    Exit Sub
End If

If (pwidth / theta) < d And d > 0 Then 'reach a cell that's far enough
    tmp_x = pSize / (d ^ 3)
    fx = fx + (x - pCoM(1)) * tmp_x
    Fy = Fy + (y - pCoM(2)) * tmp_x
Else
    For n = 1 To UBound(pchild)
        If Not pchild(n) Is Nothing Then Call pchild(n).NetForce_Traverse(ptIdx, x, y, fx, Fy, theta)
    Next n
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

