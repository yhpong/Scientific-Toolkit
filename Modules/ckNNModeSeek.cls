VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ckNNModeSeek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pk As Long                  'number of k neighbors
Private pn_dimension As Long        'dimension of input features
Private pn_cluster As Long          'number of clusters discovered
Private px_cluster() As Long        'cluster index of each data point
Private pcluster_size() As Long     'number of points assigned to each mode
Private pmode_list() As Long        'data points selected as modes
Private pmode_val() As Double       'value of each mode
Private pCost As Double
Private Const pINFINITY As Double = 2.5E+30


'========================================
'Access Properties
'========================================
Public Property Get n_neighbor() As Long
    n_neighbor = pk
End Property

Public Property Get n_dimension() As Long
    n_dimension = pn_dimension
End Property

Public Property Get n_cluster() As Long
    n_cluster = pn_cluster
End Property

Public Property Get Cost() As Double
    Cost = pCost
End Property

Public Property Get x_cluster() As Long()
    x_cluster = px_cluster
End Property

Public Property Get mode_list() As Long()
    mode_list = pmode_list
End Property

Public Property Get mode_val() As Double()
    mode_val = pmode_val
End Property

Public Property Get cluster_size() As Long()
    cluster_size = pcluster_size
End Property

Sub Reset()
    pk = 0
    Erase pmode_list, pmode_val, pcluster_size, px_cluster
End Sub


'========================================
'Main Algorithm
'========================================
Sub Clustering(x() As Double, Optional k As Long = 5, Optional dist_type As String = "EUCLIDEAN")
Dim i As Long, j As Long, m As Long, n As Long, ii As Long, iter_chk As Long, imax As Long
Dim tmp_max As Double, tmp_x As Double
Dim kNeighbors As Variant, kNeighbor() As Long, x_mode() As Long
Dim kDist() As Double, rho() As Double
Dim kT1 As ckdTree

    n = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    
    'Find kNN-Density of each point
    ReDim rho(1 To n)
    ReDim kNeighbors(1 To n)
    Set kT1 = New ckdTree
    With kT1
        Call .Build_Tree(x)
        For i = 1 To n
            Call .kNN_Search(i, k, x, kNeighbor, kDist, dist_type, True)
            rho(i) = UBound(kDist) / kDist(UBound(kDist))
            kNeighbors(i) = kNeighbor
        Next i
    End With
    Erase kDist
    Set kT1 = Nothing
    
    'Find mode of each point
    ReDim x_mode(1 To n)
    For ii = 1 To n
        iter_chk = 0
        i = ii
        Do
            kNeighbor = kNeighbors(i)
            tmp_max = rho(i): imax = i
            For m = 1 To UBound(kNeighbor)
                j = kNeighbor(m)
                If rho(j) > tmp_max Then
                    tmp_max = rho(j): imax = j
                End If
            Next m
            If imax = i Then Exit Do
            i = imax
            iter_chk = iter_chk + 1
        Loop While iter_chk < 1000
        If iter_chk >= 1000 Then
            Debug.Print "ckNNModeSeek: Clustering: Error in mode search."
        End If
        x_mode(ii) = imax
    Next ii
    Erase kNeighbor, kNeighbors
    
    'Group modes into clusters
    Call modMath.Unique_Items(x_mode, px_cluster, pmode_list, pn_cluster, pcluster_size, True)
    ReDim pmode_val(1 To pn_cluster, 1 To pn_dimension)
    For i = 1 To pn_cluster
        j = pmode_list(i)
        For m = 1 To pn_dimension
            pmode_val(i, m) = x(j, m)
        Next m
    Next i
    
    'evaluate cost function of current partition
    pCost = 0
    If dist_type = "EUCLIDEAN" Then
        For i = 1 To n
            j = px_cluster(i)
            tmp_x = 0
            For m = 1 To pn_dimension
                tmp_x = tmp_x + (x(i, m) - pmode_val(j, m)) ^ 2
            Next m
            pCost = pCost + tmp_x
        Next i
    ElseIf dist_type = "MANHATTAN" Then
        For i = 1 To n
            j = px_cluster(i)
            tmp_x = 0
            For m = 1 To pn_dimension
                tmp_x = tmp_x + Abs(x(i, m) - pmode_val(j, m))
            Next m
            pCost = pCost + tmp_x
        Next i
    Else
        Debug.Print "Invalid dist_type, Cost not evaluated."
    End If
    pCost = pCost / n
    
End Sub

