VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ckMeanCluster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'========================================
'Requires: modMath, ckdTree
'========================================
'Perform k-Means clustering when distance type is "EUCLIDEAN". k-Medians will be
'used if distance type is set to "MANHATTAN". When "CORREL" is used, mean is used
'to calculate cluster centers and correlation is used to evalulate distance between
'datapoints to clusters.
'When number of k is large, one may set usekdTree to TRUE to speed up NN search.
'========================================

'=== Major input and output data format
'Input:  x(1:N,1:D), N observations of D-dimensional data
'        k, number of desired clusters
'        iter_max, maximum number of iterations to try
'        strType, distance metric "EUCLIDEAN", "MANHATTAN", "CORREL"
'Output: x_cluster(1:N), cluster index of each data point, integer vector
'        cluster_mean(1:k, 1:D), storing centers of each cluster, real array
'        cluster_size(1:k),number of points assigned to each cluster, integer vector
'========================================

'=== Training methods
'.kMean_Brute(x,k,iter_max,strType)
'.kMean_Clustering(x,k,iter_max,strType,usekdtree)
'.kMean_Filtering(x,k,iter_max,strType,usekdtree)
'.kMean_Elkan(x,k,iter_max,strType)
'.kMean_Hamerly(x,k,iter_max,strType)
'.kMean_Annular(x,k,iter_max,strType)
'========================================

Private pk As Long, pn_dimension As Long    'number of clusters and dimensions
Private px_cluster() As Long        'cluster index of each data point
Private pcluster_mean() As Double   'center of each cluster
Private pcluster_size() As Long     'number of points assigned to each cluster
Private pDistType As String         'Metric used in current model
Private Const pINFINITY As Double = 2.5E+30
Private px2() As Double, pc2() As Double, pxavg() As Double, pxsd() As Double, pcavg() As Double, pcsd() As Double


'========================================
'Access Properties
'========================================
Public Property Get k_cluster() As Long
    k_cluster = pk
End Property

Public Property Get n_dimension() As Long
    n_dimension = pn_dimension
End Property

Public Property Get cluster_mean() As Double()
    cluster_mean = pcluster_mean
End Property

Public Property Get cluster_size() As Long()
    cluster_size = pcluster_size
End Property

Public Property Get x_cluster() As Long()
    x_cluster = px_cluster
End Property

Public Property Get DistType() As String
    DistType = pDistType
End Property

Sub Reset()
    pk = 0
    Erase pcluster_mean, pcluster_size, px_cluster
    Erase px2, pc2, pxavg, pxsd, pcavg, pcsd
End Sub

'Output trained model to Excel
Sub Print_Model(vRng As Range)
    With vRng
        .Value = pk                         'Number of centroids
        .Offset(1, 0).Value = pn_dimension  'dimension of data
        .Offset(2, 0).Value = pDistType     'Metric used
        .Offset(3, 0).Resize(pk, pn_dimension).Value = pcluster_mean 'Centroid vectors
    End With
End Sub

'Read trained model from Excel
Sub Read_Model(vRng As Range)
Dim j As Long, k As Long
Dim vArr As Variant
    With vRng
        pk = .Value                         'Number of centroids
        pn_dimension = .Offset(1, 0).Value  'dimension of data
        pDistType = .Offset(2, 0).Value     'Metric used
        vArr = .Offset(3, 0).Resize(pk, pn_dimension).Value
        ReDim pcluster_mean(1 To pk, 1 To pn_dimension)
        For j = 1 To pn_dimension
            For k = 1 To pk
                pcluster_mean(k, j) = vArr(k, j)
            Next k
        Next j
        Erase vArr
    End With
End Sub

'Assign un-trained data to centroid, use upper bound of Elkan's algorithm to speed up search
'Input:  x(1:N,1:D), N observations of D-dimensional data
'Output: x_idx(1:N), cluster index of each data point, integer vector
Sub Assign_Data(x_idx() As Long, x)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim c_dist() As Double, s_dist() As Double, xi() As Double
Dim tmp_x As Double, tmp_y As Double
    n = UBound(x, 1)
    ReDim x_idx(1 To n)
    ReDim xi(1 To pn_dimension)
    If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
        Call Calc_Aux(pcluster_mean, pc2, pcavg, pcsd)
    End If
    Call Compute_Centroid_Centroid_Distance(c_dist, s_dist)
    For i = 1 To n
        For j = 1 To pn_dimension
            xi(j) = x(i, j)
        Next j
        m = 1
        tmp_y = Dist2Center_2(xi, 1)
        For k = 2 To pk
            If tmp_y > (c_dist(k, m) / 2) Then
                tmp_x = Dist2Center_2(xi, k)
                If tmp_x < tmp_y Then
                    tmp_y = tmp_x
                    m = k
                End If
            End If
        Next k
        x_idx(i) = m
    Next i
    Erase c_dist, s_dist, xi
End Sub

'Initialization with k-means++ scheme
Private Sub Init_kplus(x() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double
Dim x_d2() As Double, xi() As Double
Dim x_seeds() As Long
    n = UBound(x, 1)
    Randomize
    ReDim x_seeds(1 To pk)  'mark datapoints chosen as seeds
    ReDim x_d2(1 To n)      'distances from datapoints to closest seed
    ReDim xi(1 To n)        'cumulative probability of datapoints to be chosen
    For i = 1 To n
        x_d2(i) = pINFINITY
    Next i
    x_seeds(1) = Int(Rnd() * n) + 1 'First seed randomly picked
    For k = 2 To pk 'Select next seed base on distances from existing seeds
        tmp_x = 0
        For i = 1 To n
            Call Dist2Seeds(i, x_seeds, x, k - 1, x_d2(i))
            tmp_x = tmp_x + x_d2(i)
        Next i
        xi(1) = x_d2(1) / tmp_x
        For i = 2 To n
            xi(i) = xi(i - 1) + x_d2(i) / tmp_x
        Next i
        x_seeds(k) = Random_Integer_Prob(xi)
    Next k
'*****Fix initialization for testing purpose*****
'    For k = 1 To pk
'        x_seeds(k) = (k - 1) * Int(n / pk) + 1
'    Next k
'*******************************************************
    ReDim pcluster_mean(1 To pk, 1 To pn_dimension)
    For m = 1 To pn_dimension
        For k = 1 To pk
            pcluster_mean(k, m) = x(x_seeds(k), m)
        Next k
    Next m
    Erase xi, x_d2, x_seeds
End Sub


'===========================================================
'Main Algorithm, cache distances in a N x k array, only
'calculate its k-th column when the k-th centroid is moved.
'===========================================================
'Input:  x(1:N,1:D), N observations of D-dimensional data
'        k, number of desired clusters
'Output: x_cluster(1:N), cluster index of each data point, integer vector
'        cluster_mean(1:k, 1:D), storing centers of each cluster, real array
'        cluster_size(1:k),number of points assigned to each cluster, integer vector
Sub kMean_Clustering(x() As Double, k As Long, Optional iter_max As Long = 100, _
        Optional strType As String = "EUCLIDEAN", _
        Optional usekdtree As Boolean = False)
Dim i As Long, j As Long, m As Long, n As Long, p As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, tmp_min As Double
Dim n_raw As Long, n_chg As Long, cluster_reset As Long
Dim xi() As Double, x_dist() As Double
Dim isChg() As Long, isChg_Prev() As Long
Dim k1 As ckdTree
Dim tree_type As String

    pk = k
    n_raw = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    pDistType = VBA.UCase$(strType)
    If pDistType = "EUCLIDEAN" Then
        tree_type = "EUCLIDEAN"
    ElseIf pDistType = "MANHATTAN" Then
        tree_type = "MANHATTAN"
    ElseIf pDistType = "CORREL" Then
        If usekdtree = True Then
            Debug.Print "k-d Tree speedup not compatible with CORREL metric."
            End
        End If
    Else
        Debug.Print "ckMeanCluster:Invalid distance type."
        End
    End If

    '=== Initialization with k-means++ scheme
    Call Init_kplus(x)
    
    '=== Pre-calculate data to centroid distance
    ReDim isChg_Prev(0 To pk)
    ReDim x_dist(1 To n_raw, 1 To pk)
    If usekdtree = False Then
        ReDim xi(1 To pn_dimension)
        For i = 1 To n_raw
            For m = 1 To pn_dimension
                xi(m) = x(i, m)
            Next m
            For j = 1 To pk
                x_dist(i, j) = Dist2Center(xi, j)
            Next j
        Next i
    Else
        'if k-d tree is used, negative value flags
        'distance that needs to be recalculated
        For j = 1 To pk
            For i = 1 To n_raw
                x_dist(i, j) = -1
            Next i
        Next j
    End If
    
    '=== For euclidean pre-calculate norm of centroid
    If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
        Call Calc_Aux(pcluster_mean, pc2, pcavg, pcsd)
    End If

    '=== Iterate until convergence
    ReDim px_cluster(1 To n_raw)
    ReDim xi(1 To pn_dimension)
    For iterate = 1 To iter_max
        If iterate Mod 20 = 0 Then
            DoEvents: Application.StatusBar = "k-Means Clustering: " & iterate & "/" & iter_max
        End If
        
        'Update auxiliary varaibles for moved centroids
        If pDistType = "EUCLIDEAN" Then
            For i = 1 To pk
                If isChg_Prev(i) > 0 Then
                    pc2(i) = 0
                    For j = 1 To pn_dimension
                        pc2(i) = pc2(i) + pcluster_mean(i, j) ^ 2
                    Next j
                End If
            Next i
        ElseIf pDistType = "CORREL" Then
            For i = 1 To pk
                If isChg_Prev(i) > 0 Then
                    pcavg(i) = 0: pcsd(i) = 0
                    For j = 1 To pn_dimension
                        pcavg(i) = pcavg(i) + pcluster_mean(i, j)
                        pcsd(i) = pcsd(i) + pcluster_mean(i, j) ^ 2
                    Next j
                    pcavg(i) = pcavg(i) / pn_dimension
                    pcsd(i) = Sqr(pcsd(i) / pn_dimension - pcavg ^ 2)
                End If
            Next i
        End If

        If usekdtree = False Then
            'Update distance
            Call Batch_Dist_Update(x_dist, x, isChg_Prev)
        Else
            'Prepare k-d Tree to speed up search
            Set k1 = New ckdTree
            Call k1.Build_Tree(pcluster_mean)
        End If
        
        'Assign each point to the closest center
        n_chg = 0
        ReDim isChg(0 To pk)
        ReDim pcluster_size(1 To pk)
        For i = 1 To n_raw
            If usekdtree = True Then
                For m = 1 To pn_dimension
                    xi(m) = x(i, m)
                Next m
                n = k1.NN_Search_Speed(xi, pcluster_mean, i, x_dist, tree_type)
            Else
                n = 1: tmp_min = x_dist(i, 1)
                For j = 2 To pk
                    tmp_x = x_dist(i, j)
                    If tmp_x < tmp_min Then
                        tmp_min = tmp_x
                        n = j
                    End If
                Next j
            End If
            'Record if any centroid needs to be moved
            If n <> px_cluster(i) Then
                n_chg = n_chg + 1
                isChg(n) = isChg(n) + 1
                isChg(px_cluster(i)) = isChg(px_cluster(i)) + 1
            End If
            px_cluster(i) = n
            pcluster_size(n) = pcluster_size(n) + 1
        Next i
        
        'If there are centroids without any members, set them to be
        'datapoints that are furthest from existing clusters.
        cluster_reset = Fill_Empty_Cluster(x, n_chg, isChg)

        'Clear k-d Tree if used
        If usekdtree = True Then
            Call k1.Reset
            Set k1 = Nothing
            'if any centroid is changed then reset their distances
            For j = 1 To pk
                If isChg(j) > 0 Then
                    For i = 1 To n_raw
                        x_dist(i, j) = -1
                    Next i
                End If
            Next j
        End If
        
        isChg_Prev = isChg 'Record centroids that were moved
        
        If cluster_reset = 0 Then
            If n_chg = 0 Then Exit For
        End If
        
        'Re-estimate centroid vector base on new cluster membership
        Call ReCalculate_Centroids(x)
        
    Next iterate
    '=========================================
    If iterate >= iter_max Then
        Debug.Print "ckMeanCluster: Fail to converge in " & iter_max & " steps."
    End If
    Erase xi, x_dist, isChg, isChg_Prev
    If pDistType = "EUCLIDEAN" Then
        Erase px2, pc2
    ElseIf pDistType = "CORREL" Then
        Erase pxavg, pxsd, pcavg, pcsd
    End If
    Application.StatusBar = False
End Sub


'=========================================================================
'Brute Force Approach, no caching of distances, no nothing
'=========================================================================
Sub kMean_Brute(x() As Double, k As Long, Optional iter_max As Long = 100, _
        Optional strType As String = "EUCLIDEAN")
Dim i As Long, j As Long, m As Long, n As Long, p As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, tmp_min As Double
Dim n_raw As Long, n_chg As Long, cluster_reset As Long
Dim xi() As Double

    pk = k
    pDistType = VBA.UCase$(strType)
    If pDistType <> "EUCLIDEAN" And pDistType <> "MANHATTAN" And pDistType <> "CORREL" Then
        Debug.Print "ckMeanCluster: Brute: Invalid distance type."
        End
    End If
    n_raw = UBound(x, 1)
    pn_dimension = UBound(x, 2)

    '=== Initialization with k-means++ scheme
    Call Init_kplus(x)

    '=== Iterate until convergence
    ReDim px_cluster(1 To n_raw)
    ReDim xi(1 To pn_dimension)
    For iterate = 1 To iter_max
        If iterate Mod 20 = 0 Then
            DoEvents: Application.StatusBar = "k-Means Clustering: " & iterate & "/" & iter_max
        End If

        'Assign each point to the closest center
        n_chg = 0
        ReDim pcluster_size(1 To pk)
        For i = 1 To n_raw
            For m = 1 To pn_dimension
                xi(m) = x(i, m)
            Next m
            tmp_min = pINFINITY
            For j = 1 To pk
                tmp_x = Dist2Center(xi, j)
                If tmp_x < tmp_min Then
                    tmp_min = tmp_x
                    n = j
                End If
            Next j
            If px_cluster(i) <> n Then n_chg = n_chg + 1
            px_cluster(i) = n
            pcluster_size(n) = pcluster_size(n) + 1
        Next i

        'If there are centroids without any members, set them to be
        'datapoints that are furthest from existing clusters.
        cluster_reset = Fill_Empty_Cluster(x, n_chg)

        If cluster_reset = 0 Then
            If n_chg = 0 Then Exit For
        End If
        
        'Re-estimate centroid vector base on new cluster membership
        Call ReCalculate_Centroids(x)

    Next iterate
    '=========================================
    If iterate >= iter_max Then
        Debug.Print "ckMeanCluster: Brute: Fail to converge in " & iter_max & " steps."
    End If
    Erase xi
    Application.StatusBar = False
End Sub


'===========================================================
'Speed up k-Means clustering with filtering algorithm
'===========================================================
'"An Efficientk-Means Clustering Algorithm:Analysis and Implementation", Tapas Kanungo (2002)
Sub kMean_Filtering(x() As Double, k As Long, Optional iter_max As Long = 100, _
        Optional strType As String = "EUCLIDEAN", Optional use_kdTree As Boolean = False)
Dim i As Long, j As Long, m As Long, n As Long, p As Long, iterate As Long
Dim n_raw As Long, n_chg As Long, cluster_reset As Long
Dim prev_idx() As Long
Dim node_size() As Long, node_wgtcenter As Variant, node_min() As Double, node_max() As Double, node_idx As Variant
Dim k1 As ckdTree
Dim tree_type As String
Dim x_dist() As Double, isChg() As Long

    pk = k
    pDistType = VBA.UCase$(strType)
    If pDistType = "EUCLIDEAN" Then
        tree_type = "EUCLIDEAN"
    ElseIf pDistType = "MANHATTAN" Then
        tree_type = "MANHATTAN"
    Else
        Debug.Print "ckMeanCluster: Filtering: Invalid distance type."
        End
    End If
    n_raw = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    
    'Initialization with k-means++ scheme
    Call Init_kplus(x)
    
    'Build k-d Tree for data points
    Set k1 = New ckdTree
    Call k1.kMean_Build_Tree(x, node_size, node_wgtcenter, node_min, node_max, pDistType)
    
    'Array holding pairwise distance between data and centroids
    ', use -1 to indicate values that are not calculated yet
    ReDim x_dist(1 To n_raw, 1 To pk)
    For j = 1 To pk
        For i = 1 To n_raw
            x_dist(i, j) = -1
        Next i
    Next j
    
    'Iterate until convergence
    ReDim px_cluster(1 To n_raw)
    ReDim prev_idx(1 To n_raw)
    For iterate = 1 To iter_max
        If iterate Mod 20 = 0 Then
            DoEvents: Application.StatusBar = "k-Means Clustering: " & iterate & "/" & iter_max
        End If
        
        'Use filtering algorithm to assign data to centroids
        Call k1.kMean_Assign_Center(x, node_size, node_wgtcenter, node_min, node_max, _
                pcluster_mean, pcluster_size, px_cluster, x_dist, pDistType, use_kdTree)
        
        'check if any cluster assignment has changed
        n_chg = 0: ReDim isChg(0 To pk)
        For i = 1 To n_raw
            If px_cluster(i) <> prev_idx(i) Then
                n_chg = n_chg + 1
                isChg(px_cluster(i)) = isChg(px_cluster(i)) + 1
                isChg(prev_idx(i)) = isChg(prev_idx(i)) + 1
            End If
        Next i
        
        'Reset empty centroids
        cluster_reset = Fill_Empty_Cluster(x, n_chg, isChg)

        If cluster_reset = 0 Then
            If n_chg = 0 Then Exit For
        Else
            Call ReCalculate_Centroids(x)
        End If
        
        'If any centroid has moved reset its distance to data
        For j = 1 To pk
            If isChg(j) > 0 Then
                For i = 1 To n_raw
                    x_dist(i, j) = -1
                Next i
            End If
        Next j
        
        prev_idx = px_cluster
    Next iterate

    If iterate >= iter_max Then
        Debug.Print "ckMeanCluster: Filtering: Fail to converge in " & iter_max & " steps."
    End If
    Call k1.Reset: Set k1 = Nothing
    Erase node_size, node_wgtcenter, node_min, node_max, prev_idx, x_dist, isChg
    Application.StatusBar = False
End Sub


'===========================================================
'Speed up k-Means clustering with triangle inequality
'===========================================================
'"Using the Triangle Inequality to Accelerate k-Means", Charlkes Elkan (2003)
'Note that since triangle inequalit is used, square root needs to to be taken on Euclidean distance
Sub kMean_Elkan(x() As Double, k As Long, Optional iter_max As Long = 100, _
        Optional strType As String = "EUCLIDEAN")
Dim i As Long, j As Long, m As Long, n As Long, p As Long, iterate As Long
Dim n_raw As Long, n_chg As Long, cluster_reset As Long, c_cur As Long
Dim prev_idx() As Long
Dim c_dist() As Double, s_dist() As Double, prev_c() As Double, c_chg() As Double
Dim ub() As Double, lb() As Double, xi() As Double
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim r_chk As Boolean
    
    pDistType = VBA.UCase$(strType)
    If pDistType <> "EUCLIDEAN" And pDistType <> "MANHATTAN" And pDistType <> "CORREL" Then
        Debug.Print "ckMeanCluster: Elkan: Invalid distance type."
        End
    End If
    
    pk = k
    n_raw = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    
    'Initialization with k-means++ scheme
    Call Init_kplus(x)
    
    'Assign all points to same cluster, initialize invalid bounds
    ReDim px_cluster(1 To n_raw)
    ReDim ub(1 To n_raw)
    ReDim lb(1 To n_raw, 1 To pk)
    For i = 1 To n_raw
        px_cluster(i) = 1
        ub(i) = pINFINITY
    Next i
    
    If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
        Call Calc_Aux(x, px2, pxavg, pxsd)
    End If
    
    'Iterate until convergence
    prev_idx = px_cluster
    ReDim xi(1 To pn_dimension)
    For iterate = 1 To iter_max
        If iterate Mod 20 = 0 Then
            DoEvents: Application.StatusBar = "k-Means Elkan: " & iterate & "/" & iter_max
        End If
        
        If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
            Call Calc_Aux(pcluster_mean, pc2, pcavg, pcsd)
        End If
        
        'Compute centroid-centroid distances
        Call Compute_Centroid_Centroid_Distance(c_dist, s_dist)

        'Assign data to nearest centroid
        For i = 1 To n_raw
            c_cur = px_cluster(i): tmp_x = ub(i)
            If tmp_x > s_dist(c_cur) Then
                r_chk = True    'Boolean to check if upper bound needs tightening
                For p = 1 To pn_dimension
                    xi(p) = x(i, p)
                Next p
                For n = 1 To pk
                    If n <> c_cur Then
                        tmp_z = (0.5 * c_dist(n, c_cur)): If lb(i, n) > tmp_z Then tmp_z = lb(i, n)
                        If tmp_x > tmp_z Then
                            If r_chk = True Then
                                ub(i) = Dist2Center_2_Fast(xi, i, c_cur)
                                tmp_x = ub(i)
                                r_chk = False
                            End If
                            If tmp_x > tmp_z Then
                                tmp_y = Dist2Center_2_Fast(xi, i, n)
                                lb(i, n) = tmp_y
                                If tmp_y < tmp_x Then
                                    px_cluster(i) = n: c_cur = n
                                    ub(i) = tmp_y: tmp_x = tmp_y
                                End If
                            End If
                        End If
                    End If
                Next n
            End If
        Next i

        'Re-estimate cluster size
        ReDim pcluster_size(1 To pk)
        For i = 1 To n_raw
            j = px_cluster(i)
            pcluster_size(j) = pcluster_size(j) + 1
        Next i
        
        n_chg = 0
        cluster_reset = Fill_Empty_Cluster(x, n_chg)
        
        If cluster_reset = 0 Then
        
            'Re-estimate centroid positions and record distance moved
            prev_c = pcluster_mean
            ReDim c_chg(1 To pk)
            Call ReCalculate_Centroids(x)
            For j = 1 To pk
                For p = 1 To pn_dimension
                    xi(p) = prev_c(j, p)
                Next p
                c_chg(j) = Dist2Center_2(xi, j)
            Next j
    
            'Update lower and upper bound
            For i = 1 To n_raw
                ub(i) = ub(i) + c_chg(px_cluster(i))
            Next i
            For j = 1 To pk
                tmp_x = c_chg(j)
                For i = 1 To n_raw
                    lb(i, j) = lb(i, j) - tmp_x
                    If lb(i, j) < 0 Then lb(i, j) = 0
                Next i
            Next j
    
            'check if any cluster assignment has changed
            n_chg = 0
            For i = 1 To n_raw
                If px_cluster(i) <> prev_idx(i) Then
                    n_chg = n_chg + 1
                    Exit For
                End If
            Next i
            If n_chg = 0 Then Exit For
            
            prev_idx = px_cluster
            
        End If
    Next iterate
    If iterate >= iter_max Then
        Debug.Print "ckMeanCluster: Elkan: Fail to converge in " & iter_max & " steps."
    End If
    Erase c_dist, s_dist, ub, lb, prev_idx, xi, c_chg
    If pDistType = "EUCLIDEAN" Then
        Erase px2, pc2
    ElseIf pDistType = "CORREL" Then
        Erase pxavg, pxsd, pcavg, pcsd
    End If
    Application.StatusBar = False
End Sub


'=================================================================================
'Speed up k-Means clustering with triangle inequality and only one lower bound
'=================================================================================
'"Making k-means even faster", Greg Hamerly (2010)
Sub kMean_Hamerly(x() As Double, k As Long, Optional iter_max As Long = 100, _
        Optional strType As String = "EUCLIDEAN")
Dim i As Long, j As Long, m As Long, n As Long, p As Long, iterate As Long
Dim n_raw As Long, n_chg As Long, cluster_reset As Long, c_cur As Long
Dim prev_idx() As Long
Dim c_dist() As Double, s_dist() As Double, prev_c() As Double
Dim c_chg() As Double, c_chg_max As Double, c_chg_max2 As Double
Dim ub() As Double, lb() As Double, xi() As Double
Dim tmp_x As Double, tmp_y As Double, z As Double
Dim h1 As cHeap

    pDistType = VBA.UCase(strType)
    If pDistType <> "EUCLIDEAN" And pDistType <> "MANHATTAN" And pDistType <> "CORREL" Then
        Debug.Print "ckMeanCluster: Hamerly: Invalid distance type."
        End
    End If

    pk = k
    n_raw = UBound(x, 1)
    pn_dimension = UBound(x, 2)

    'Initialization with k-means++ scheme
    Call Init_kplus(x)

    'Assign all points to same cluster, initialize invalid bounds
    ReDim px_cluster(1 To n_raw)    'index to nearest centroid
    ReDim ub(1 To n_raw)            'upper bound nearest centroid
    ReDim lb(1 To n_raw)            'lower bound 2nd nearest centroid
    For i = 1 To n_raw
        px_cluster(i) = 1
        ub(i) = pINFINITY
    Next i

    If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
        Call Calc_Aux(x, px2, pxavg, pxsd)
    End If

    'Iterate until convergence
    Set h1 = New cHeap     'Heap is used to find closest and 2nd closest centroids
    prev_idx = px_cluster
    ReDim xi(1 To pn_dimension)
    For iterate = 1 To iter_max
        If iterate Mod 20 = 0 Then
            DoEvents: Application.StatusBar = "k-Means Hamerly: " & iterate & "/" & iter_max
        End If

        If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
            Call Calc_Aux(pcluster_mean, pc2, pcavg, pcsd)
        End If

        'Compute centroid-centroid distances
        Call Compute_Centroid_Centroid_Distance(c_dist, s_dist)

        'Assign data to nearest centroid
        For i = 1 To n_raw
            c_cur = px_cluster(i)
            z = s_dist(c_cur): If lb(i) > z Then z = lb(i)
            If ub(i) > z Then
                For p = 1 To pn_dimension
                    xi(p) = x(i, p)
                Next p
                ub(i) = Dist2Center_2_Fast(xi, i, c_cur)
                If ub(i) > z Then
                    With h1
                        Call .Init("MIN")
                        Call .Add(ub(i), c_cur)
                        For j = 1 To pk
                            If j <> c_cur Then Call .Add(Dist2Center_2_Fast(xi, i, j), j)
                        Next j
                        Call .Pop_Min(tmp_x, j)
                        If j <> c_cur Then
                            ub(i) = tmp_x
                            px_cluster(i) = j
                        End If
                        Call .Pop_Min(tmp_x, j)
                        lb(i) = tmp_x
                    End With
                End If
            End If
        Next i

        'Re-estimate cluster size
        ReDim pcluster_size(1 To pk)
        For i = 1 To n_raw
            j = px_cluster(i)
            pcluster_size(j) = pcluster_size(j) + 1
        Next i

        n_chg = 0
        cluster_reset = Fill_Empty_Cluster(x, n_chg)

        If cluster_reset = 0 Then

            'Re-estimate centroid positions and record distance moved
            prev_c = pcluster_mean
            Call ReCalculate_Centroids(x)
            ReDim c_chg(1 To pk)
            Call h1.Init("MAX")
            For j = 1 To pk
                For p = 1 To pn_dimension
                    xi(p) = prev_c(j, p)
                Next p
                c_chg(j) = Dist2Center_2(xi, j)
                Call h1.Add(c_chg(j), j)
            Next j
            Call h1.Pop_Max(c_chg_max, n)
            Call h1.Pop_Max(c_chg_max2, m)

            'Update lower and upper bound
            For i = 1 To n_raw
                ub(i) = ub(i) + c_chg(px_cluster(i))
                If px_cluster(i) = n Then
                    lb(i) = lb(i) - c_chg_max2
                Else
                    lb(i) = lb(i) - c_chg_max
                End If
            Next i

            'check if any cluster assignment has changed
            n_chg = 0
            For i = 1 To n_raw
                If px_cluster(i) <> prev_idx(i) Then
                    n_chg = n_chg + 1
                    Exit For
                End If
            Next i
            If n_chg = 0 Then Exit For

            prev_idx = px_cluster

        End If
    Next iterate
    If iterate >= iter_max Then
        Debug.Print "ckMeanCluster: Hamerly: Fail to converge in " & iter_max & " steps."
    End If
    Call h1.Reset
    Set h1 = Nothing
    Erase c_dist, s_dist, ub, lb, prev_idx
    If pDistType = "EUCLIDEAN" Then
        Erase px2, pc2
    ElseIf pDistType = "CORREL" Then
        Erase pxavg, pxsd, pcavg, pcsd
    End If
    Application.StatusBar = False
End Sub




'=================================================================================
'Speed up k-Means clustering with triangle inequality and annular search constraint
'=================================================================================
'"Accelerating Lloyd's Algorithm for k-Means Clsutering", Greg Hamerly and Jonathan Drake (2014)
'Note that since triangle inequality is used, square root needs to to be taken on Euclidean distance
Sub kMean_Annular(x() As Double, k As Long, Optional iter_max As Long = 100, _
        Optional strType As String = "EUCLIDEAN")
Dim i As Long, j As Long, m As Long, n As Long, p As Long, iterate As Long, jj As Long, j_start As Long
Dim n_raw As Long, n_chg As Long, cluster_reset As Long, c_cur As Long, c_cur2 As Long
Dim prev_idx() As Long
Dim c_dist() As Double, s_dist() As Double, prev_c() As Double
Dim c_chg() As Double, c_chg_max As Double, c_chg_max2 As Double
Dim ub() As Double, lb() As Double, xi() As Double
Dim tmp_x As Double, tmp_y As Double, z As Double
Dim x_norm() As Double, c_norm() As Double, sort_idx() As Long, c_sort() As Long, x_idx2() As Long
Dim h1 As cHeap
Dim c_max As Double, c_min As Double

    pDistType = VBA.UCase(strType)
    If pDistType <> "EUCLIDEAN" And pDistType <> "MANHATTAN" And pDistType <> "CORREL" Then
        Debug.Print "ckMeanCluster: Elkan: Invalid distance type."
        End
    End If

    pk = k
    n_raw = UBound(x, 1)
    pn_dimension = UBound(x, 2)

    'Initialization with k-means++ scheme
    Call Init_kplus(x)

    'Find norm of x()
    Call Calc_Vec_Norm(x, x_norm, strType)

    'Assign all points to same cluster, initialize invalid bounds
    ReDim px_cluster(1 To n_raw)    'index to nearest centroid
    ReDim x_idx2(1 To n_raw)        'index to 2nd nearest centroid
    ReDim ub(1 To n_raw)            'upper bound nearest centroid
    ReDim lb(1 To n_raw)            'lower bound on 2nd nearest centroid
    For i = 1 To n_raw
        px_cluster(i) = 1
        x_idx2(i) = 2
        ub(i) = pINFINITY
    Next i
    
    If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
        Call Calc_Aux(x, px2, pxavg, pxsd)
    End If
        
    'Iterate until convergence
    Set h1 = New cHeap     'Heap is used to find closest and 2nd closest centroids
    prev_idx = px_cluster
    ReDim xi(1 To pn_dimension)
    For iterate = 1 To iter_max
        If iterate Mod 20 = 0 Then
            DoEvents: Application.StatusBar = "k-Means Annular: " & iterate & "/" & iter_max
        End If
        
        If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
            Call Calc_Aux(pcluster_mean, pc2, pcavg, pcsd)
        End If
        
        'Compute centroid-centroid distances
        Call Compute_Centroid_Centroid_Distance(c_dist, s_dist)

        'Find norm of centroids and sort in ascending order
        Call Calc_Vec_Norm(pcluster_mean, c_norm, strType)
        Call modMath.Sort_Quick_A(c_norm, 1, pk, c_sort)

        'Assign data to nearest centroid
        For i = 1 To n_raw
            c_cur = px_cluster(i)
            z = s_dist(c_cur): If lb(i) > z Then z = lb(i)
            If ub(i) > z Then
                For p = 1 To pn_dimension
                    xi(p) = x(i, p)
                Next p
                ub(i) = Dist2Center_2_Fast(xi, i, c_cur)
                If ub(i) > z Then
                    c_cur2 = x_idx2(i)
                    lb(i) = Dist2Center_2_Fast(xi, i, c_cur2)
                    z = ub(i): If lb(i) > z Then z = lb(i)
                    c_max = x_norm(i) + z   'Outer ring of annulus
                    c_min = x_norm(i) - z   'Inner ring of annulus
                    j_start = Binary_Search(c_norm, c_min)
                    If j_start = 0 Then j_start = 1
                    With h1
                        Call .Init("MIN")
                        Call .Add(ub(i), c_cur)
                        Call .Add(lb(i), c_cur2)
                        For j = j_start To pk
                            If c_norm(j) > c_max Then Exit For
                            jj = c_sort(j)
                            If jj <> c_cur And jj <> c_cur2 Then
                                Call .Add(Dist2Center_2_Fast(xi, i, jj), jj)
                            End If
                        Next j
                        Call .Pop_Min(tmp_x, j)
                        If tmp_x < ub(i) Then
                            ub(i) = tmp_x
                            px_cluster(i) = j
                        End If
                        Call .Pop_Min(tmp_x, j)
                        If tmp_x < lb(i) Then
                            lb(i) = tmp_x
                            x_idx2(i) = j
                        End If
                    End With
                End If
            End If
        Next i

        'Re-estimate cluster size
        ReDim pcluster_size(1 To pk)
        For i = 1 To n_raw
            j = px_cluster(i)
            pcluster_size(j) = pcluster_size(j) + 1
        Next i

        n_chg = 0
        cluster_reset = Fill_Empty_Cluster(x, n_chg)

        If cluster_reset = 0 Then

            'Re-estimate centroid positions and record distance moved
            prev_c = pcluster_mean
            Call ReCalculate_Centroids(x)
            ReDim c_chg(1 To pk)
            Call h1.Init("MAX")
            For j = 1 To pk
                For p = 1 To pn_dimension
                    xi(p) = prev_c(j, p)
                Next p
                c_chg(j) = Dist2Center_2(xi, j)
                Call h1.Add(c_chg(j), j)
            Next j
            Call h1.Pop_Max(c_chg_max, n)
            Call h1.Pop_Max(c_chg_max2, m)

            'Update lower and upper bound
            For i = 1 To n_raw
                ub(i) = ub(i) + c_chg(px_cluster(i))
                If px_cluster(i) = n Then
                    lb(i) = lb(i) - c_chg_max2
                Else
                    lb(i) = lb(i) - c_chg_max
                End If
            Next i

            'check if any cluster assignment has changed
            n_chg = 0
            For i = 1 To n_raw
                If px_cluster(i) <> prev_idx(i) Then
                    n_chg = n_chg + 1
                    Exit For
                End If
            Next i
            If n_chg = 0 Then Exit For

            prev_idx = px_cluster

        End If
    Next iterate
    If iterate >= iter_max Then
        Debug.Print "ckMeanCluster: Annular: Fail to converge in " & iter_max & " steps."
    End If
    Call h1.Reset
    Set h1 = Nothing
    Erase c_dist, s_dist, ub, lb, prev_idx, x_idx2, x_norm, c_norm, c_sort
    If pDistType = "EUCLIDEAN" Then
        Erase px2, pc2
    ElseIf pDistType = "CORREL" Then
        Erase pxavg, pxsd, pcavg, pcsd
    End If
    Application.StatusBar = False
End Sub


Private Sub Compute_Centroid_Centroid_Distance(c_dist() As Double, s_dist() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double
    ReDim c_dist(1 To pk, 1 To pk)
    ReDim s_dist(1 To pk)
    
    'Compute Pair-wise distances between centroids
    For n = 1 To pk - 1
        For m = n + 1 To pk
            c_dist(m, n) = DistC2C_Fast(m, n) 'sqroot has to be taken
            c_dist(n, m) = c_dist(m, n)
        Next m
    Next n

    'Compute half-distance of each centroid to its nearest centroid
    For n = 1 To pk
        tmp_x = pINFINITY
        For m = 1 To pk
            If m <> n Then
                If c_dist(m, n) < tmp_x Then
                    tmp_x = c_dist(m, n)
                End If
            End If
        Next m
        s_dist(n) = tmp_x / 2
    Next n
End Sub


'Update distances between data and moved centroids, innermost loop
'arranged to run over first index for better efficiency
Private Sub Batch_Dist_Update(x_dist() As Double, x() As Double, isChg() As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_y As Double
    n = UBound(x, 1)
    For k = 1 To pk
        If isChg(k) > 0 Then
            For i = 1 To n
                x_dist(i, k) = 0
            Next i
            If pDistType = "EUCLIDEAN" Then
                'use ||x||^2+||c||^2-2x.c, but skip ||x||^2 since it won't affect rankings
                For j = 1 To pn_dimension
                    tmp_x = 2 * pcluster_mean(k, j)
                    For i = 1 To n
                        x_dist(i, k) = x_dist(i, k) + x(i, j) * tmp_x
                    Next i
                Next j
                tmp_x = pc2(k)
                For i = 1 To n
                    x_dist(i, k) = tmp_x - x_dist(i, k)
                Next i
            ElseIf pDistType = "MANHATTAN" Then
                For j = 1 To pn_dimension
                    tmp_x = pcluster_mean(k, j)
                    For i = 1 To n
                        x_dist(i, k) = x_dist(i, k) + Abs(x(i, j) - tmp_x)
                    Next i
                Next j
            ElseIf pDistType = "CORREL" Then
                'use correl(x,c)= (<xc>-<x><c>) /(sx sc)
                For j = 1 To pn_dimension
                    tmp_x = pcluster_mean(k, j)
                    For i = 1 To n
                        x_dist(i, k) = x_dist(i, k) + x(i, j) * tmp_x
                    Next i
                Next j
                tmp_x = pcavg(k): tmp_y = pcsd(k)
                For i = 1 To n
                    x_dist(i, k) = 2 * (1 - (x_dist(i, k) / pn_dimension - pxavg(i) * tmp_x) / (pxsd(i) * tmp_y))
                Next i
            Else
                Debug.Print "ckMeanCluster: Batch_Dist_Update: Invalid metric."
                End
            End If
        End If
    Next k
End Sub

'Find distance between point i to cluster center k
Private Function Dist2Center(xi() As Double, k As Long) As Double
Dim j As Long
Dim tmp_x As Double, avg1 As Double, avg2 As Double, sd1 As Double, sd2 As Double
    tmp_x = 0
    If pDistType = "EUCLIDEAN" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + (xi(j) - pcluster_mean(k, j)) ^ 2
        Next j
    ElseIf pDistType = "MANHATTAN" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + Abs(xi(j) - pcluster_mean(k, j))
        Next j
    ElseIf pDistType = "CORREL" Then
        avg1 = 0: avg2 = 0
        sd1 = 0: sd2 = 0
        For j = 1 To pn_dimension
            avg1 = avg1 + xi(j)
            avg2 = avg2 + pcluster_mean(k, j)
            sd1 = sd1 + xi(j) ^ 2
            sd2 = sd2 + pcluster_mean(k, j) ^ 2
        Next j
        avg1 = avg1 / pn_dimension
        avg2 = avg2 / pn_dimension
        sd1 = Sqr(sd1 / pn_dimension - avg1 ^ 2)
        sd2 = Sqr(sd2 / pn_dimension - avg2 ^ 2)
        For j = 1 To pn_dimension
            tmp_x = tmp_x + (xi(j) - avg1) * (pcluster_mean(k, j) - avg2)
        Next j
        tmp_x = 2 * (1 - tmp_x / (sd1 * sd2 * pn_dimension))
    Else
        Debug.Print "k-means clustering:Dist2Center: Invalid distance metric."
        Exit Function
    End If
    Dist2Center = tmp_x
End Function

'Find distance between point i to cluster center k, taking square root on Euclidean distance
Private Function Dist2Center_2(xi() As Double, k As Long) As Double
Dim j As Long
Dim tmp_x As Double, avg1 As Double, avg2 As Double, sd1 As Double, sd2 As Double
    tmp_x = 0
    If pDistType = "EUCLIDEAN" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + (xi(j) - pcluster_mean(k, j)) ^ 2
        Next j
        tmp_x = Sqr(tmp_x)
    ElseIf pDistType = "MANHATTAN" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + Abs(xi(j) - pcluster_mean(k, j))
        Next j
    ElseIf pDistType = "CORREL" Then
        avg1 = 0: avg2 = 0
        sd1 = 0: sd2 = 0
        For j = 1 To pn_dimension
            avg1 = avg1 + xi(j)
            avg2 = avg2 + pcluster_mean(k, j)
            sd1 = sd1 + xi(j) ^ 2
            sd2 = sd2 + pcluster_mean(k, j) ^ 2
        Next j
        avg1 = avg1 / pn_dimension
        avg2 = avg2 / pn_dimension
        sd1 = Sqr(sd1 / pn_dimension - avg1 ^ 2)
        sd2 = Sqr(sd2 / pn_dimension - avg2 ^ 2)
        For j = 1 To pn_dimension
            tmp_x = tmp_x + (xi(j) - avg1) * (pcluster_mean(k, j) - avg2)
        Next j
        tmp_x = Sqr(2 * (1 - tmp_x / (sd1 * sd2 * pn_dimension)))
    Else
        Debug.Print "k-means clustering:Dist2Center_2: Invalid distance metric."
        Exit Function
    End If
    Dist2Center_2 = tmp_x
End Function

'Find distance between point i to cluster center k, taking square root on Euclidean distance
'and auxiliary variables to speed up calculations
Private Function Dist2Center_2_Fast(xi() As Double, i As Long, k As Long) As Double
Dim j As Long
Dim tmp_x As Double
    tmp_x = 0
    If pDistType = "EUCLIDEAN" Then
        tmp_x = px2(i) + pc2(k)
        For j = 1 To pn_dimension
            tmp_x = tmp_x - 2 * xi(j) * pcluster_mean(k, j)
        Next j
        If tmp_x < 0 Then tmp_x = 0 'Caution: possible negative number from floating error
        tmp_x = Sqr(tmp_x)
    ElseIf pDistType = "MANHATTAN" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + Abs(xi(j) - pcluster_mean(k, j))
        Next j
    ElseIf pDistType = "CORREL" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + xi(j) * pcluster_mean(k, j)
        Next j
        tmp_x = (tmp_x / pn_dimension - pxavg(i) * pcavg(k)) / (pxsd(i) * pxsd(k))
        tmp_x = Sqr(2 * (1 - tmp_x))
    Else
        Debug.Print "k-means clustering:Dist2Center_2_Fast: Invalid distance metric."
        Exit Function
    End If
    Dist2Center_2_Fast = tmp_x
End Function

'Find distance between centroids i & j
Private Function DistC2C_Fast(i As Long, j As Long) As Double
Dim p As Long
Dim tmp_x As Double
    If pDistType = "EUCLIDEAN" Then
        tmp_x = pc2(i) + pc2(j)
        For p = 1 To pn_dimension
            tmp_x = tmp_x - 2 * pcluster_mean(i, p) * pcluster_mean(j, p)
        Next p
        If tmp_x < 0 Then tmp_x = 0 'Caution: possible negative number from floating error
        tmp_x = Sqr(tmp_x)
    ElseIf pDistType = "MANHATTAN" Then
        tmp_x = 0
        For p = 1 To pn_dimension
            tmp_x = tmp_x + Abs(pcluster_mean(i, p) - pcluster_mean(j, p))
        Next p
    ElseIf pDistType = "CORREL" Then
        tmp_x = 0
        For p = 1 To pn_dimension
            tmp_x = tmp_x + pcluster_mean(i, p) * pcluster_mean(j, p)
        Next p
        tmp_x = (tmp_x / pn_dimension - pcavg(i) * pcavg(j)) / (pcsd(i) * pcsd(j))
        tmp_x = Sqr(2 * (1 - tmp_x))
    Else
        Debug.Print "k-means clustering:DistC2C_Fast: Invalid distance metric."
        Exit Function
    End If
    DistC2C_Fast = tmp_x
End Function


'Returns number of empty clusters. Also assign
'empty clusters to points farthest from existing clusters.
Private Function Fill_Empty_Cluster(x() As Double, _
            n_chg As Long, Optional isChg As Variant) As Long
Dim i As Long, j As Long, k As Long, m As Long, n As Long, k2 As Long
Dim tmp_x As Double, tmp_y As Double
Dim xi() As Double
    Fill_Empty_Cluster = 0
    n = UBound(x, 1)
    ReDim xi(1 To pn_dimension)
    For k = 1 To pk
        If pcluster_size(k) = 0 Then
            Fill_Empty_Cluster = Fill_Empty_Cluster + 1
            tmp_y = -pINFINITY
            For i = 1 To n
                For m = 1 To pn_dimension
                    xi(m) = x(i, m)
                Next m
                tmp_x = Dist2Center(xi, px_cluster(i))
                If tmp_x > tmp_y Then
                    tmp_y = tmp_x
                    j = i
                End If
            Next i
            
            k2 = px_cluster(j)
            
            If IsMissing(isChg) = False Then
                isChg(k) = isChg(k) + 1
                isChg(k2) = isChg(k2) + 1
            End If
            
            n_chg = n_chg + 1
            px_cluster(j) = k
            pcluster_size(k) = pcluster_size(k) + 1
            pcluster_size(k2) = pcluster_size(k2) - 1
            For m = 1 To pn_dimension
                pcluster_mean(k, m) = x(j, m)
            Next m

        End If
    Next k
    If Fill_Empty_Cluster > 0 Then
        Debug.Print "kMeanCluster: " & Fill_Empty_Cluster & " centroids with no membership. Consider reducing k."
    End If
End Function

'Find distance between point i and closest seed value
Private Sub Dist2Seeds(i As Long, seed() As Long, x() As Double, n_seed As Long, cur_best As Double)
Dim j As Long, d As Long
Dim tmp_x As Double
Dim avg1 As Double, avg2 As Double, sd1 As Double, sd2 As Double
    j = seed(n_seed)    'newly picked seed from last iteration
    If i = j Then       'i is seed itself
        cur_best = 0
        Exit Sub
    End If
    'Distance from i to last picked seed
    tmp_x = 0
    If pDistType = "EUCLIDEAN" Then
        For d = 1 To pn_dimension
            tmp_x = tmp_x + (x(i, d) - x(j, d)) ^ 2
        Next d
    ElseIf pDistType = "MANHATTAN" Then
        For d = 1 To pn_dimension
            tmp_x = tmp_x + Abs(x(i, d) - x(j, d))
        Next d
    ElseIf pDistType = "CORREL" Then
        avg1 = 0: sd1 = 0
        avg2 = 0: sd2 = 0
        For d = 1 To pn_dimension
            avg1 = avg1 + x(i, d): sd1 = sd1 + x(i, d) ^ 2
            avg2 = avg2 + x(j, d): sd2 = sd2 + x(j, d) ^ 2
        Next d
        avg1 = avg1 / pn_dimension
        avg2 = avg2 / pn_dimension
        sd1 = Sqr(sd1 / pn_dimension - avg1 ^ 2)
        sd2 = Sqr(sd2 / pn_dimension - avg2 ^ 2)
        For d = 1 To pn_dimension
            tmp_x = tmp_x + (x(i, d) - avg1) * (x(j, d) - avg2)
        Next d
        tmp_x = 2 * (1 - tmp_x / (sd1 * sd2 * pn_dimension))
     Else
        Debug.Print "k-means clustering:Dist2Seeds: Invalid distance metric"
        Exit Sub
    End If
    If tmp_x < cur_best Then cur_best = tmp_x 'new seed is closer to i, then update cur_best
End Sub

'Re-estimate centroid vector base on new cluster membership
Private Sub ReCalculate_Centroids(x() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim iArr() As Long, xj() As Double
    n = UBound(x, 1)
    ReDim pcluster_mean(1 To pk, 1 To pn_dimension)
    If pDistType = "EUCLIDEAN" Or pDistType = "CORREL" Then
        For i = 1 To n
            k = px_cluster(i)
            For j = 1 To pn_dimension
                pcluster_mean(k, j) = pcluster_mean(k, j) + x(i, j)
            Next j
        Next i
        For j = 1 To pn_dimension
            For k = 1 To pk
                pcluster_mean(k, j) = pcluster_mean(k, j) / pcluster_size(k)
            Next k
        Next j
    ElseIf pDistType = "MANHATTAN" Then
        For k = 1 To pk
            m = 0
            ReDim iArr(1 To pcluster_size(k))
            For i = 1 To n
                If px_cluster(i) = k Then
                    m = m + 1
                    iArr(m) = i
                End If
            Next i
            For j = 1 To pn_dimension
                Call modMath.Filter_Array(x, xj, iArr, j)
                pcluster_mean(k, j) = modMath.fmedian(xj)
            Next j
        Next k
        Erase iArr, xj
    Else
        Debug.Print "ckMeanCluster: ReCalculate_Centroids: Invalid distance type."
        End
    End If
End Sub

'Pick a random integer between 1 to N with given probablity distribution
Private Function Random_Integer_Prob(prob_C() As Double) As Long
Dim i As Long, n As Long
Dim tmp_x As Double
    n = UBound(prob_C)
    tmp_x = Rnd()
    If tmp_x <= prob_C(1) Then
        Random_Integer_Prob = 1
        Exit Function
    End If
    For i = 2 To n
'        Random_Integer_Prob = Binary_Search(prob_C, tmp_x) + 1
'        If Random_Integer_Prob > n Then Random_Integer_Prob = n
        If tmp_x > prob_C(i - 1) And tmp_x <= prob_C(i) Then
            Random_Integer_Prob = i
            Exit For
        End If
    Next i
End Function

'Search for address i in sorted array x() s.t. x(i) <= tgt < x(i+1)
Private Function Binary_Search(x() As Double, tgt As Double) As Long
Dim iL As Long, iR As Long, n As Long, m As Long
    n = UBound(x): iL = 0: iR = n + 1
    Do While (iR - iL) > 1
        m = (iL + iR) \ 2
        If tgt >= x(m) Then
            iL = m
        Else
            iR = m
        End If
    Loop
    If tgt = x(1) Then
        Binary_Search = 1
    ElseIf tgt = x(n) Then
        Binary_Search = n - 1
    Else
        Binary_Search = iL
    End If
End Function

'Calculate auxiliary  variables
Private Sub Calc_Aux(x() As Double, x2() As Double, xavg, xsd)
Dim i As Long, j As Long, k As Long, n As Long
    n = UBound(x, 1)
    If pDistType = "EUCLIDEAN" Then
        ReDim x2(1 To n)
        For j = 1 To pn_dimension
            For i = 1 To n
                x2(i) = x2(i) + x(i, j) ^ 2
            Next i
        Next j
    ElseIf pDistType = "CORREL" Then
        ReDim xavg(1 To n)
        ReDim xsd(1 To n)
        For j = 1 To pn_dimension
            For i = 1 To n
                xavg(i) = xavg(i) + x(i, j)
            Next i
            For i = 1 To n
                xsd(i) = xsd(i) + x(i, j) ^ 2
            Next i
        Next j
        For i = 1 To n
            xavg(i) = xavg(i) / pn_dimension
        Next i
        For i = 1 To n
            xsd(i) = Sqr(xsd(i) / pn_dimension - xavg(i) ^ 2)
        Next i
    End If
End Sub

'Norm of vector to be used in annular algorithm
Private Sub Calc_Vec_Norm(x() As Double, x_norm() As Double, strType As String)
Dim i As Long, j As Long, n As Long, n_dimension As Long, tmp_x As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim x_norm(1 To n)
    If strType = "EUCLIDEAN" Then
        For j = 1 To n_dimension
            For i = 1 To n
                x_norm(i) = x_norm(i) + x(i, j) ^ 2
            Next i
        Next j
        For i = 1 To n
            x_norm(i) = Sqr(x_norm(i))
        Next i
    ElseIf strType = "MANHATTAN" Then
        For j = 1 To n_dimension
            For i = 1 To n
                x_norm(i) = x_norm(i) + Abs(x(i, j))
            Next i
        Next j
    Else
        Debug.Print "Calc_Vec_Norm: Invalid norm type."
    End If
End Sub
