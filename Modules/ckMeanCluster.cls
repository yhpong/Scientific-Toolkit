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


Private px() As Double
Private pk As Long
Private px_cluster() As Long        'cluster index of each data point
Private pcluster_mean() As Double   'center of each cluster
Private pcluster_size() As Long     'number of points  assigned to each cluster


Public Property Get raw_data() As Double()
    raw_data = px
End Property

Public Property Get k_cluster() As Long
    k_cluster = pk
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
    Erase px, pcluster_mean, pcluster_size, px_cluster
End Sub


'Input: x(1 to n_raw, 1 to n_dimension)
'Input: k, number of desired clusters
'Output: x_cluster(1 to n_raw), cluster index of each data point
'Output: cluster_mean(1 to k, 1 to n_dimension), center of each cluster
'Output: cluster_size(1 to k), number of points  assigned to each cluster
Sub kMean_Clustering(x() As Double, k As Long, Optional iterate_max As Long = 100, Optional strType As String = "EUCLIDEAN")
Dim i As Long, j As Long, m As Long, n As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, tmp_min As Double
Dim n_raw As Long, n_dimension As Long
Dim n_chg As Long
Dim x_d2() As Double
Dim x_seeds() As Long

strType = UCase(strType)
n_raw = UBound(x, 1)
n_dimension = UBound(x, 2)
px = x
pk = k
ReDim px_cluster(1 To n_raw)
ReDim pcluster_size(1 To pk)
ReDim pcluster_mean(1 To pk, 1 To n_dimension)

'=== Initialization with k-means++ scheme
Randomize
ReDim x_seeds(0 To 0)
ReDim x_d2(1 To n_raw)
Call Append_1D(x_seeds, Int(Rnd() * n_raw) + 1) 'First seed randomly picked
'Select next seed base on distances from existing seeds
For n = 2 To k
    tmp_x = 0
    For i = 1 To n_raw
        x_d2(i) = Dist2Seeds(i, x_seeds, x, strType)
        tmp_x = tmp_x + x_d2(i)
    Next i
    For i = 1 To n_raw
        x_d2(i) = x_d2(i) / tmp_x
    Next i
    i = Random_Integer_Prob(x_d2)
    Call Append_1D(x_seeds, i)
Next n
For n = 1 To k
    For m = 1 To n_dimension
        pcluster_mean(n, m) = x(x_seeds(n), m)
    Next m
Next n
Erase x_d2, x_seeds
'=========================================

'=== Iterate until convergence
n_chg = 0
For iterate = 1 To iterate_max
    DoEvents
    Application.StatusBar = "k-Means Clustering..." & iterate & "/" & iterate_max
    
    'Assign each point to the closest center
    ReDim pcluster_size(1 To k)
    For i = 1 To n_raw
        tmp_min = Exp(70)
        n = 1
        For j = 1 To k
            tmp_x = Dist2Center(x, i, j, strType)
            If tmp_x < tmp_min Then
                tmp_min = tmp_x
                n = j
            End If
        Next j
        n_chg = n_chg + Abs(n - px_cluster(i))
        px_cluster(i) = n
        pcluster_size(n) = pcluster_size(n) + 1
    Next i
    
    'Recalculate mean of each assigned cluster
    ReDim pcluster_mean(1 To k, 1 To n_dimension)
    For i = 1 To n_raw
        n = px_cluster(i)
        For m = 1 To n_dimension
            pcluster_mean(n, m) = pcluster_mean(n, m) + x(i, m)
        Next m
    Next i
    
    For n = 1 To k
        For m = 1 To n_dimension
            pcluster_mean(n, m) = pcluster_mean(n, m) / pcluster_size(n)
        Next m
    Next n
    
    'check for convergence over the past 5 iterations
    If (iterate - 1) Mod 5 = 0 Then
        If n_chg = 0 Then
            Exit For
        Else
            n_chg = 0
        End If
    End If
Next iterate
'=========================================

Application.StatusBar = False

End Sub


'Find distance between point i and closest seed value
Private Function Dist2Seeds(i As Long, seed() As Long, x() As Double, Optional strType As String = "EUCLIDEAN") As Double
Dim j As Long, d As Long, k As Long
Dim n_dimension As Long, n_seed As Long
Dim tmp_x As Double, tmp_min As Double
Dim avg1 As Double, avg2 As Double, sd1 As Double, sd2 As Double
n_dimension = UBound(x, 2)
n_seed = UBound(seed)
tmp_min = Exp(70)
For k = 1 To n_seed
    j = seed(k)
    If i = j Then
        tmp_min = 0
        Exit For
    Else
        tmp_x = 0
        If strType = "EUCLIDEAN" Then
            For d = 1 To n_dimension
                tmp_x = tmp_x + (x(i, d) - x(j, d)) ^ 2
            Next d
        ElseIf strType = "CORREL" Then
            avg1 = 0
            avg2 = 0
            sd1 = 0
            sd2 = 0
            For d = 1 To n_dimension
                avg1 = avg1 + x(i, d)
                avg2 = avg2 + x(j, d)
                sd1 = sd1 + x(i, d) ^ 2
                sd2 = sd2 + x(j, d) ^ 2
            Next d
            avg1 = avg1 / n_dimension
            avg2 = avg2 / n_dimension
            sd1 = Sqr(sd1 / n_dimension - avg1 ^ 2)
            sd2 = Sqr(sd2 / n_dimension - avg2 ^ 2)
            For d = 1 To n_dimension
                tmp_x = tmp_x + (x(i, d) - avg1) * (x(j, d) - avg2)
            Next d
            tmp_x = 2 * (1 - tmp_x / (sd1 * sd2 * n_dimension))
         Else
            Debug.Print "k-means clustering: Invalid distance metric"
            Exit Function
        End If
        
        If tmp_x < tmp_min Then tmp_min = tmp_x
    End If
Next k
Dist2Seeds = tmp_min
End Function


'Find distance between point i to cluster center k
Private Function Dist2Center(x() As Double, i As Long, k As Long, Optional strType As String = "EUCLIDEAN") As Double
Dim j As Long, d As Long, n_dimension As Long
Dim tmp_x As Double
Dim avg1 As Double, avg2 As Double, sd1 As Double, sd2 As Double
    n_dimension = UBound(x, 2)
    tmp_x = 0
    If strType = "EUCLIDEAN" Then
        For d = 1 To n_dimension
            tmp_x = tmp_x + (x(i, d) - pcluster_mean(k, d)) ^ 2
        Next d
    ElseIf strType = "CORREL" Then
        avg1 = 0
        avg2 = 0
        sd1 = 0
        sd2 = 0
        For d = 1 To n_dimension
            avg1 = avg1 + x(i, d)
            avg2 = avg2 + pcluster_mean(k, d)
            sd1 = sd1 + x(i, d) ^ 2
            sd2 = sd2 + pcluster_mean(k, d) ^ 2
        Next d
        avg1 = avg1 / n_dimension
        avg2 = avg2 / n_dimension
        sd1 = Sqr(sd1 / n_dimension - avg1 ^ 2)
        sd2 = Sqr(sd2 / n_dimension - avg2 ^ 2)
        For d = 1 To n_dimension
            tmp_x = tmp_x + (x(i, d) - avg1) * (pcluster_mean(k, d) - avg2)
        Next d
        tmp_x = 2 * (1 - tmp_x / (sd1 * sd2 * n_dimension))
    Else
        Debug.Print "k-means clustering: Invalid distance metric"
        Exit Function
    End If
    Dist2Center = tmp_x
End Function


'Add tgt to the end of a 1D vector x(0 to n)
Private Sub Append_1D(x As Variant, tgt As Variant)
Dim n As Long
    n = UBound(x) + 1
    ReDim Preserve x(0 To n)
    x(n) = tgt
End Sub


'Pick a random integer between 1 to N with given probablity distribution
Private Function Random_Integer_Prob(Prob() As Double) As Long
Dim i As Long, n_raw As Long
Dim tmp_x As Double
Dim prob_C() As Double

    n_raw = UBound(Prob)
    ReDim prob_C(0 To n_raw)
    prob_C(1) = Prob(1)
    For i = 2 To n_raw
        prob_C(i) = prob_C(i - 1) + Prob(i)
    Next i
    
    Randomize
    tmp_x = Rnd()
    For i = 1 To n_raw
        If tmp_x >= prob_C(i - 1) And tmp_x <= prob_C(i) Then
            Random_Integer_Prob = i
            Exit For
        End If
    Next i
End Function

