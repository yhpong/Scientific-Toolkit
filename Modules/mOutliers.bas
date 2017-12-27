Attribute VB_Name = "mOutliers"
Option Explicit


'============================================================================
'"Linear-Time Outlier Detection via Sensitivity", Mario Lucic, 2016
'============================================================================
'Input x(1 to n_raw, 1 to n_dimension)
Function Influence_Iterate(x() As Double, Optional iterate_max As Long = 30, _
    Optional k_min As Long = 2, Optional k_max As Long = 10, Optional k_step As Long = 2)
Dim i As Long, k As Long, n_raw As Long, iterate As Long
Dim s() As Double, s_temp() As Double
    n_raw = UBound(x, 1)
    ReDim s(1 To n_raw)
    For iterate = 1 To iterate_max
        DoEvents
        Application.StatusBar = "Outlier by Influence: " & iterate & "/" & iterate_max
        s_temp = Influence(x, k_min, k_max, k_step)
        For i = 1 To n_raw
            s(i) = s(i) + s_temp(i) / iterate_max
        Next i
    Next iterate
    Influence_Iterate = s
    Erase s, s_temp
    Application.StatusBar = False
End Function

'Input x(1 to n_raw, 1 to n_dimension)
Private Function Influence(x() As Double, Optional k_min As Long = 2, Optional k_max As Long = 10, Optional k_step As Long = 2) As Double()
Dim i As Long, j As Long, m As Long, n As Long, k As Long, n_k As Long
Dim tmp_x As Double, tmp_min As Double
Dim n_raw As Long, bi As Long
Dim alpha As Double, c_phi As Double, INFINITY As Double
Dim B_Set() As Long, B_Master() As Long
Dim x2B() As Long
Dim P_Dist() As Double
Dim P_Count() As Long
Dim dist() As Double, min_dist() As Double
Dim s() As Double, s_avg() As Double

INFINITY = Exp(70)
n_raw = UBound(x, 1)
ReDim s(1 To n_raw, 1 To ((k_max - k_min) \ k_step + 1))
ReDim B_Master(0 To 0)
ReDim dist(1 To n_raw)
ReDim x2B(1 To n_raw, 1 To k_max)
ReDim min_dist(1 To n_raw, 1 To k_max)
ReDim P_Count(1 To k_max, 1 To k_max)

Randomize
k = Int(Rnd() * n_raw) + 1
Call Append_1D(B_Master, k)
ReDim P_Dist(1 To n_raw)
For i = 1 To n_raw
    dist(i) = INFINITY
Next i
For k = 1 To k_max
    tmp_x = 0
    For i = 1 To n_raw
        If k > 1 Then
            j = x2B(i, k - 1)
        ElseIf k = 1 Then
            j = 1
        End If
        Call Dist_n_B(i, B_Master, x, k, j, dist(i))
        x2B(i, k) = j
        min_dist(i, k) = dist(i)
        P_Count(j, k) = P_Count(j, k) + 1
        tmp_x = tmp_x + dist(i)
    Next i
    If k < k_max Then
        For i = 1 To n_raw
            P_Dist(i) = dist(i) / tmp_x
        Next i
        i = modMath.Random_Integer_Prob(P_Dist)
        Call Append_1D(B_Master, i)
    End If
Next k
Erase dist, P_Dist

n_k = 0
For k = k_min To k_max Step k_step
    
    n_k = n_k + 1
    
    B_Set = B_Master
    ReDim Preserve B_Set(0 To k)
    
    alpha = 16 * (Log(k) / Log(2) + 2)
    c_phi = 0
    For i = 1 To n_raw
        c_phi = c_phi + min_dist(i, k)
    Next i
    c_phi = c_phi / n_raw

    ReDim P_Dist(1 To UBound(B_Set))
    For i = 1 To n_raw
        j = x2B(i, k)
        P_Dist(j) = P_Dist(j) + min_dist(i, k)
    Next i
    
    For i = 1 To n_raw
        j = x2B(i, k)
        s(i, n_k) = alpha * min_dist(i, k) / c_phi + _
                2 * alpha * P_Dist(j) / (P_Count(j, k) * c_phi) + _
                n_raw * 4# / P_Count(j, k)
        s(i, n_k) = s(i, n_k) / n_raw
    Next i

Next k

ReDim Preserve s(1 To n_raw, 1 To n_k)
ReDim s_avg(1 To n_raw)
For i = 1 To n_raw
    tmp_x = 0
    For k = 1 To n_k
        tmp_x = tmp_x + s(i, k)
    Next k
    s_avg(i) = tmp_x / n_k
Next i
Erase s, B_Master, P_Dist, P_Count, min_dist
Influence = s_avg
Erase s_avg
End Function

'Find Euclidean distance between point n and set B() using minimum distance
'x2B = the member of B that is closest to n
Private Sub Dist_n_B(n As Long, B() As Long, x() As Double, n_seed As Long, x2B As Long, cur_dist As Double)
Dim tmp_x As Double
    tmp_x = Dist_n_m(n, B(n_seed), x)
    If tmp_x < cur_dist Then
        cur_dist = tmp_x
        x2B = n_seed
    End If
End Sub

'Find Euclidean distance between points n & m
Private Function Dist_n_m(n As Long, m As Long, x() As Double) As Double
Dim i As Long, d As Long
    d = UBound(x, 2)
    Dist_n_m = 0
    For i = 1 To d
        Dist_n_m = Dist_n_m + (x(n, i) - x(m, i)) ^ 2
    Next i
End Function


'=== Compute Mahalanobis Distance
'Input:  x(1:N,1:D), N by D data array
'Output: MD(1:N)
Function MahalanobisDist(x() As Double) As Double()
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim dimension As Long, n_raw As Long
Dim x_avg As Double, tmp_x As Double
Dim covar_m() As Double, x_centered() As Double, MD() As Double
    DoEvents
    Application.StatusBar = "Calculating Mahalanobis Distance..."
    n_raw = UBound(x, 1)
    dimension = UBound(x, 2)
    ReDim covar_m(1 To dimension, 1 To dimension)
    ReDim x_centered(1 To n_raw, 1 To dimension)

    'Calculate covariance matrix and Center each dimension to mean zero
    For m = 1 To dimension
        x_avg = 0
        For i = 1 To n_raw
            x_avg = x_avg + x(i, m)
        Next i
        x_avg = x_avg / n_raw
        For i = 1 To n_raw
            x_centered(i, m) = x(i, m) - x_avg
            covar_m(m, m) = covar_m(m, m) + x_centered(i, m) ^ 2
        Next i
        covar_m(m, m) = covar_m(m, m) / (n_raw - 1)
    Next m
    For m = 1 To dimension - 1
        For n = m + 1 To dimension
            tmp_x = 0
            For i = 1 To n_raw
                tmp_x = tmp_x + x_centered(i, m) * x_centered(i, n)
            Next i
            covar_m(n, m) = tmp_x / (n_raw - 1)
            covar_m(m, n) = covar_m(n, m)
        Next n
    Next m
    
    'Inverse of covariance matrix
    covar_m = modMath.Matrix_Inverse(covar_m)
    
    'Calculate Mahalanobis Distance
    ReDim MD(1 To n_raw)
    For n = 1 To dimension - 1
        For m = n + 1 To dimension
            tmp_x = 2 * covar_m(m, n)
            For i = 1 To n_raw
                MD(i) = MD(i) + x_centered(i, n) * x_centered(i, m) * tmp_x
            Next i
        Next m
    Next n
    For n = 1 To dimension
        tmp_x = covar_m(n, n)
        For i = 1 To n_raw
            MD(i) = MD(i) + (x_centered(i, n) ^ 2) * tmp_x
        Next i
    Next n
    For i = 1 To n_raw
        MD(i) = Sqr(MD(i))
    Next i
    MahalanobisDist = MD
    Erase covar_m, MD, x_centered
    Application.StatusBar = False
End Function


'=== Find Distance to K-th Nearest Neighbor
'Input:  x(1:N,1:D), N by D data array
'        k, number of neighbors
'        usekdTree, use k-d Tree to speed up nearest neighbor seach when set to TRUE
'        dis_type, "EUCLIDEAN" or 'MANHATTAN"
'Output: KthNeighborDist(1:N), real vector of k-th neighbor distance for each data point
Function KthNeighborDist(x() As Double, Optional k As Long = 10, Optional usekdtree As Boolean = False, _
            Optional dist_type As String = "EUCLIDEAN") As Double()
Dim i As Long, j As Long, n As Long, n_raw As Long
Dim neighbor_dist() As Double, neighbor() As Long
Dim dist() As Double, Dk() As Double
Dim kT1 As ckdTree
Dim strType As String
    DoEvents
    Application.StatusBar = "Calculating k-th Nearest Neighbor..."
    strType = VBA.UCase(dist_type)
    
    If usekdtree = True Then
        Set kT1 = New ckdTree
        Call kT1.kNN_All(neighbor, dist, x, k, 1, strType)
        KthNeighborDist = dist
        Erase neighbor, dist
        Set kT1 = Nothing
        Exit Function
    End If

    n_raw = UBound(x, 1)
    If strType = "EUCLIDEAN" Then
        dist = modMath.Calc_Euclidean_Dist(x, False)
    ElseIf strType = "MANHATTAN" Then
        dist = modMath.Calc_Manhattan_Dist(x)
    Else
        Debug.Print "mOutliers:KthNeighborDist:Invalid metric " & dist_type
        Exit Function
    End If
    ReDim Dk(1 To n_raw)
    ReDim neighbor_dist(1 To n_raw - 1)
    ReDim neighbor(1 To n_raw - 1)
    For i = 1 To n_raw
        n = 0
        For j = 1 To n_raw
            If j <> i Then
                n = n + 1
                neighbor_dist(n) = dist(i, j)
                neighbor(n) = j
            End If
        Next j
        Call modMath.Sort_Quick_A(neighbor_dist, 1, n_raw - 1, neighbor, 0)
        Dk(i) = neighbor_dist(k)
    Next i
    If strType = "EUCLIDEAN" Then
        For i = 1 To n_raw
            Dk(i) = Sqr(Dk(i))
        Next i
    End If
    KthNeighborDist = Dk
    Erase dist, neighbor_dist, neighbor, Dk
    Application.StatusBar = False
End Function


'=== Find Local Outlier Factors, using k-d tree to speed up search
'Input:  feature vectors x(1:N,1:D), and number of neighbors k
'Output: LOF(1:N)
Function LOF(x() As Double, Optional k As Long = 5) As Double()
Dim i As Long, j As Long, m As Long, n As Long, n_raw As Long, v As Long
Dim tmp_x As Double, tmp_y As Double
Dim kDist() As Double, Dk() As Double, kDists As Variant
Dim kNeighbor() As Long, kNeighbors As Variant
Dim LRD() As Double, LOF_Output() As Double
Dim kT1 As ckdTree
    DoEvents
    Application.StatusBar = "Calculating Local Outlier Factor..."
    
    n_raw = UBound(x, 1)
    
    Set kT1 = New ckdTree
    With kT1
        Call .Build_Tree(x)
        ReDim kNeighbors(1 To n_raw) 'list of k-neighbors
        ReDim kDists(1 To n_raw)     'Distances to the k neighbors
        ReDim Dk(1 To n_raw)         'Distance to the k-th neighbor
        For i = 1 To n_raw
            Call .kNN_Search(i, k, x, kNeighbor, kDist, "EUCLIDEAN", True)
            kNeighbors(i) = kNeighbor
            kDists(i) = kDist
            Dk(i) = kDist(UBound(kDist, 1))
        Next i
    End With
    
    ReDim LRD(1 To n_raw)
    For i = 1 To n_raw
        tmp_y = 0
        kNeighbor = kNeighbors(i)
        kDist = kDists(i)
        For j = 1 To UBound(kNeighbor, 1)
            v = kNeighbor(j)
            tmp_x = kDist(j)
            If Dk(v) > tmp_x Then tmp_x = Dk(v)
            tmp_y = tmp_y + tmp_x
        Next j
        LRD(i) = UBound(kNeighbor) / tmp_y
    Next i
    Erase Dk, kDist, kDists

    ReDim LOF_Output(1 To n_raw)
    For i = 1 To n_raw
        tmp_x = 0
        kNeighbor = kNeighbors(i)
        For j = 1 To UBound(kNeighbor)
            tmp_x = tmp_x + LRD(kNeighbor(j))
        Next j
        LOF_Output(i) = tmp_x / (LRD(i) * UBound(kNeighbor))
    Next i
    
    LOF = LOF_Output
    Erase kNeighbors, LRD, LOF_Output
    Application.StatusBar = False
End Function



'Directly calculate LOF with brute force
Function LOF_Direct(x() As Double, Optional k As Long = 5) As Double()
Dim i As Long, j As Long, m As Long, n As Long, n_raw As Long
Dim tmp_x As Double
Dim neighbor_dist() As Double, neighbor() As Long
Dim dist() As Double, Dk() As Double
Dim kNeighbors() As Long
Dim ReachDist() As Double, LRD() As Double, LOF_Output() As Double

    DoEvents
    Application.StatusBar = "Calculating Local Outlier Factor..."

    n_raw = UBound(x, 1)
    dist = modMath.Calc_Euclidean_Dist(x, True)

    ReDim Dk(1 To n_raw)    'Distance to k-th neighbor
    ReDim kNeighbors(1 To n_raw, 1 To n_raw) 'neighbors that are as near as the k-th neighbor

    ReDim neighbor_dist(1 To n_raw - 1)
    ReDim neighbor(1 To n_raw - 1)
    For i = 1 To n_raw
        n = 0
        For j = 1 To n_raw
            If j <> i Then
                n = n + 1
                neighbor_dist(n) = dist(i, j)
                neighbor(n) = j
            End If
        Next j
        Call modMath.Sort_Quick_A(neighbor_dist, 1, n_raw - 1, neighbor, 0)
        Dk(i) = neighbor_dist(k)
        j = 1
        Do While neighbor_dist(j) <= Dk(i)
            kNeighbors(i, j) = neighbor(j)
            j = j + 1
        Loop
    Next i

    Erase neighbor_dist, neighbor

    ReDim ReachDist(1 To n_raw, 1 To n_raw)
    For i = 1 To n_raw
        For j = 1 To n_raw
            If i <> j Then
                If Dk(j) > dist(i, j) Then
                    ReachDist(i, j) = Dk(j)
                Else
                    ReachDist(i, j) = dist(i, j)
                End If
            End If
        Next j
    Next i
    Erase Dk

    ReDim LRD(1 To n_raw)
    For i = 1 To n_raw
        j = 1
        Do While kNeighbors(i, j) > 0
            LRD(i) = LRD(i) + ReachDist(i, kNeighbors(i, j))
            j = j + 1
        Loop
        LRD(i) = (j - 1) / LRD(i)
    Next i
    Erase ReachDist

    ReDim LOF_Output(1 To n_raw)
    For i = 1 To n_raw
        j = 1
        tmp_x = 0
        Do While kNeighbors(i, j) > 0
            tmp_x = tmp_x + LRD(kNeighbors(i, j))
            j = j + 1
        Loop
        LOF_Output(i) = tmp_x / (LRD(i) * (j - 1))
    Next i

    LOF_Direct = LOF_Output
    Erase kNeighbors, LRD
    Application.StatusBar = False
End Function


'=== Add i to the last element of x()
Private Sub Append_1D(x() As Long, i As Long)
Dim m As Long, n As Long
    m = LBound(x)
    n = UBound(x) + 1
    ReDim Preserve x(m To n)
    x(n) = i
End Sub


'=== Consistent data selection
'"Outlier Detection by Consistent Data Selection Method", Utkarsh Porwal et al (Dec 2017)
Function ConsistentData(x() As Double, Optional iter_max As Long = 3, _
        Optional dist_type As String = "EUCLIDEAN", Optional kList As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, p As Long, n1 As Long, n2 As Long
Dim ii As Long, jj As Long, kk As Long, iterate As Long, n_dimension As Long, n_k As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim kC1 As ckMeanCluster
Dim x1s As Variant, centroids As Variant, n1s As Variant, cmags As Variant
Dim c1() As Double, c2() As Double, cmag() As Double
Dim OutScore() As Double
Dim strType As String
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    strType = VBA.UCase(dist_type)
    If IsMissing(kList) = False Then
        n_k = UBound(kList)
    Else
        n_k = 10
        ReDim kList(1 To n_k)
        For kk = 1 To n_k
            kList(kk) = 5 * (2 ^ (kk - 1))
            If kList(kk) > Int(n * 0.75) Then
                n_k = kk - 1
                ReDim Preserve kList(1 To n_k)
                Exit For
            End If
        Next kk
    End If
    
    i = 0
    ReDim x1s(1 To n_k * iter_max)
    ReDim n1s(1 To n_k * iter_max)
    ReDim cmags(1 To n_k * iter_max)
    ReDim centroids(1 To n_k * iter_max)
    For kk = 1 To n_k
        DoEvents: Application.StatusBar = "ConsistentData:Step1: " & kk & "/" & n_k
        k = kList(kk)
        For iterate = 1 To iter_max
            i = i + 1
            Set kC1 = New ckMeanCluster
            With kC1
                If k < 50 Then
                    Call .kMean_Clustering(x, k, , strType, usekdtree:=False)
                Else
                    Call .kMean_Clustering(x, k, , strType, usekdtree:=True)
                End If
                x1s(i) = .x_cluster
                n1s(i) = .cluster_size
                c1 = .cluster_mean
                ReDim cmag(1 To k)
                For m = 1 To n_dimension
                    For j = 1 To k
                        cmag(j) = cmag(j) + c1(j, m) ^ 2
                    Next j
                Next m
                cmags(i) = cmag
                centroids(i) = c1
                Call .Reset
            End With
            Set kC1 = Nothing
        Next iterate
    Next kk
    
    ReDim OutScore(1 To n)
    For m = 1 To n
        If m Mod 500 = 0 Then
            DoEvents: Application.StatusBar = "ConsistentData:Step2 " & m & "/" & n
        End If
        p = 0
        For i = 1 To (n_k * iter_max - 1)
            ii = x1s(i)(m): n1 = n1s(i)(ii): tmp_x = cmags(i)(ii)
            c1 = centroids(i)
            For j = i + 1 To (n_k * iter_max)
                jj = x1s(j)(m): n2 = n1s(j)(jj): tmp_y = cmags(j)(jj)
                c2 = centroids(j)
                tmp_z = 0
                For k = 1 To n_dimension
                    tmp_z = tmp_z + c1(ii, k) * c2(jj, k)
                Next k
                OutScore(m) = OutScore(m) + (n1 + n2) * tmp_z / Sqr(tmp_x * tmp_y)
                p = p + (n1 + n2)
            Next j
        Next i
        OutScore(m) = 1 - OutScore(m) / p
    Next m
    ConsistentData = OutScore
    Erase x1s, n1s, cmags, c1, c2, OutScore
    Application.StatusBar = False
End Function
