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
'Requires: modMath , ckdTree
'Perform k-Means clustering when distance type is "EUCLIDEAN". k-Medians will be
'used if distance type is set to "MANHATTAN". When "CORREL" is used, mean is used
'to calculate cluster centers and correlation is used to evalulate distance between
'datapoints to clusters. kd-Tree in this case is assumed to be euclidean for simplicity.
'When number of k is large, one may set usekdTree to TRUE to speed up NN search.
'========================================

Private pk As Long, pn_dimension As Long    'number of clusters and dimensions
Private px_cluster() As Long        'cluster index of each data point
Private pcluster_mean() As Double   'center of each cluster
Private pcluster_size() As Long     'number of points assigned to each cluster
Private Const pINFINITY As Double = 2.5E+30

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

Sub Reset()
    pk = 0
    Erase pcluster_mean, pcluster_size, px_cluster
End Sub


'========================================
'Main Algorithm
'========================================
'Input:  x(1:N,1:D), N observations of D-dimensional data
'        k, number of desired clusters
'Output: x_cluster(1:N), cluster index of each data point, integer vector
'        cluster_mean(1:k, 1:D), storing centers of each cluster, real array
'        cluster_size(1:k),number of points assigned to each cluster, integer vector
Sub kMean_Clustering(x() As Double, k As Long, Optional iterate_max As Long = 100, _
        Optional strType As String = "EUCLIDEAN", _
        Optional usekdTree As Boolean = False)
Dim i As Long, j As Long, m As Long, n As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, tmp_min As Double
Dim n_raw As Long, n_chg As Long
Dim x_d2() As Double, xi() As Double, xj() As Double
Dim x_seeds() As Long, iArr() As Long
Dim k1 As ckdTree
Dim tree_type As String

    pk = k
    strType = VBA.UCase(strType)
    If strType = "CORREL" Then
        tree_type = "EUCLIDEAN"
    ElseIf strType = "EUCLIDEAN" Then
        tree_type = "EUCLIDEAN"
    ElseIf strType = "MANHATTAN" Then
        tree_type = "MANHATTAN"
    Else
        Debug.Print "ckMeanCluster:Invalid distance type."
        End
    End If
    n_raw = UBound(x, 1)
    pn_dimension = UBound(x, 2)
    'Initialization with k-means++ scheme
    Randomize
    ReDim x_seeds(1 To pk)
    ReDim x_d2(1 To n_raw)
    ReDim xi(1 To n_raw)
    For i = 1 To n_raw
        x_d2(i) = pINFINITY
    Next i
    x_seeds(1) = Int(Rnd() * n_raw) + 1 'First seed randomly picked
    For n = 2 To pk 'Select next seed base on distances from existing seeds
        tmp_x = 0
        For i = 1 To n_raw
            Call Dist2Seeds(i, x_seeds, x, n - 1, x_d2(i), strType)
            tmp_x = tmp_x + x_d2(i)
        Next i
        xi(1) = x_d2(1) / tmp_x
        For i = 2 To n_raw
            xi(i) = xi(i - 1) + x_d2(i) / tmp_x
        Next i
        x_seeds(n) = Random_Integer_Prob(xi)
    Next n
    ReDim pcluster_mean(1 To pk, 1 To pn_dimension)
    For m = 1 To pn_dimension
        For n = 1 To pk
            pcluster_mean(n, m) = x(x_seeds(n), m)
        Next n
    Next m
    Erase xi, x_d2, x_seeds
    '=========================================
    
    '=== Iterate until convergence
    n_chg = 0
    ReDim px_cluster(1 To n_raw)
    ReDim pcluster_size(1 To pk)
    ReDim xi(1 To pn_dimension)
    For iterate = 1 To iterate_max
        If iterate Mod 5 = 0 Then
            DoEvents: Application.StatusBar = "k-Means Clustering: " & iterate & "/" & iterate_max
        End If
        'Prepare k-d Tree to speed up search
        If usekdTree = True Then
            Set k1 = New ckdTree
            Call k1.Build_Tree(pcluster_mean)
        End If
        'Assign each point to the closest center
        ReDim pcluster_size(1 To pk)
        For i = 1 To n_raw
            For m = 1 To pn_dimension
                xi(m) = x(i, m)
            Next m
            If usekdTree = True Then
                n = k1.NN_Search(xi, pcluster_mean, tree_type)
            Else
                tmp_min = pINFINITY
                For j = 1 To pk
                    tmp_x = Dist2Center(xi, j, strType)
                    If tmp_x < tmp_min Then
                        tmp_min = tmp_x
                        n = j
                    End If
                Next j
            End If
            n_chg = n_chg + Abs(n - px_cluster(i))
            px_cluster(i) = n
            pcluster_size(n) = pcluster_size(n) + 1
        Next i
        
        If n_chg = 0 Then
            Exit For
        ElseIf n_chg > 0 Then
            ReDim pcluster_mean(1 To pk, 1 To pn_dimension)
            If strType = "EUCLIDEAN" Or strType = "CORREL" Then
                For i = 1 To n_raw
                    n = px_cluster(i)
                    For m = 1 To pn_dimension
                        pcluster_mean(n, m) = pcluster_mean(n, m) + x(i, m)
                    Next m
                Next i
                For m = 1 To pn_dimension
                    For n = 1 To pk
                        pcluster_mean(n, m) = pcluster_mean(n, m) / pcluster_size(n)
                    Next n
                Next m
            ElseIf strType = "MANHATTAN" Then
                For n = 1 To k
                    m = 0
                    ReDim iArr(1 To pcluster_size(n))
                    For i = 1 To n_raw
                        If px_cluster(i) = n Then
                            m = m + 1
                            iArr(m) = i
                        End If
                    Next i
                    For j = 1 To pn_dimension
                        Call modMath.Filter_Array(x, xj, iArr, j)
                        pcluster_mean(n, j) = modMath.fmedian(xj)
                    Next j
                Next n
            Else
                Debug.Print "ckMeanCluster:Invalid distance type."
                End
            End If
            n_chg = 0
        End If
    Next iterate
    '=========================================
    Application.StatusBar = False
End Sub


'Find distance between point i and closest seed value
Private Sub Dist2Seeds(i As Long, seed() As Long, x() As Double, n_seed As Long, _
                cur_best As Double, strType)
Dim j As Long, d As Long
Dim tmp_x As Double
Dim avg1 As Double, avg2 As Double, sd1 As Double, sd2 As Double
    j = seed(n_seed)
    If i = j Then
        cur_best = 0
        Exit Sub
    End If
    tmp_x = 0
    If strType = "EUCLIDEAN" Then
        For d = 1 To pn_dimension
            tmp_x = tmp_x + (x(i, d) - x(j, d)) ^ 2
        Next d
    ElseIf strType = "MANHATTAN" Then
        For d = 1 To pn_dimension
            tmp_x = tmp_x + Abs(x(i, d) - x(j, d))
        Next d
    ElseIf strType = "CORREL" Then
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
    If tmp_x < cur_best Then cur_best = tmp_x
End Sub

'Find distance between point i to cluster center k
Private Function Dist2Center(xi() As Double, k As Long, strType) As Double
Dim j As Long
Dim tmp_x As Double, avg1 As Double, avg2 As Double, sd1 As Double, sd2 As Double
    tmp_x = 0
    If strType = "EUCLIDEAN" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + (xi(j) - pcluster_mean(k, j)) ^ 2
        Next j
    ElseIf strType = "MANHATTAN" Then
        For j = 1 To pn_dimension
            tmp_x = tmp_x + Abs(xi(j) - pcluster_mean(k, j))
        Next j
    ElseIf strType = "CORREL" Then
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

'Pick a random integer between 1 to N with given probablity distribution
Private Function Random_Integer_Prob(Prob_C() As Double) As Long
Dim i As Long, n As Long
Dim tmp_x As Double
    n = UBound(Prob_C)
    tmp_x = Rnd()
    If tmp_x <= Prob_C(1) Then
        Random_Integer_Prob = 1
        Exit Function
    End If
    For i = 2 To n
        If tmp_x > Prob_C(i - 1) And tmp_x <= Prob_C(i) Then
            Random_Integer_Prob = i
            Exit For
        End If
    Next i
End Function
