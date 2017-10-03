VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ckMeanSegmental"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=======================================================================
'"Modified k-mean clustering method of HMM states for initialization
'of Baum-Welch traing algorithm", Pauline Larue et al (2011)
'=======================================================================

Private pn As Long              'length of time series x(1:N,1:D)
Private pn_dimension As Long    'dimension of time series x(1:N,1:D)
Private pn_cluster As Long      'number of clusters
Private pcluster_index() As Long    'pointer to which cluster an observation belongs to
Private pcluster_means As Variant   'jagged array holding means of each cluster
Private pcluster_covars As Variant  'jagged array holding covariance matrix of each cluster
Private pcluster_size() As Long     'integer vector holding size of each cluster
Private pprob_type As String    'probability distribution, "GAUSSIA" or "LAPALCE"


Public Property Get n() As Long
    n = pn
End Property

Public Property Get n_cluster() As Long
    n_cluster = pn_cluster
End Property

Public Property Get cluster_index() As Long()
    cluster_index = pcluster_index
End Property

Public Property Get cluster_size() As Long()
    cluster_size = pcluster_size
End Property

Public Property Get cluster_mean(k As Long) As Double()
    cluster_mean = pcluster_means(k)
End Property

Public Property Get cluster_covar(k As Long) As Double()
    cluster_covar = pcluster_covars(k)
End Property

Sub Reset()
    pn = 0
    pn_cluster = 0
    Erase pcluster_index, pcluster_means, pcluster_covars, pcluster_size
End Sub


Sub Segmentation(x() As Double, n_cluster As Long, Optional seg_len As Long = 13, _
        Optional iter_max As Long = 500, _
        Optional prob_type As String = "GAUSSIAN", Optional usePLA As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_raw As Long, n_dimension As Long, n_segment As Long, min_len As Long
Dim iArr() As Long, xArr() As Double
Dim tmp_min As Double, tmp_x As Double
Dim segment_len() As Long
Dim segment_mean() As Double, segment_covar() As Double
Dim cluster_mean() As Double, cluster_covar() As Double
Dim segment_means As Variant, segment_covars As Variant
Dim segment_index() As Long, cluster_index() As Long
Dim vArr As Variant

n_raw = UBound(x, 1)
n_dimension = UBound(x, 2)
If n_dimension > 1 And prob_type = "LAPLACE" Then
    Debug.Print "ckMeanSegmental: Segmentation: Mutlivarite only implemented for GAUSSIAN."
    Exit Sub
End If
pn = n_raw
pn_dimension = n_dimension
pn_cluster = n_cluster
pprob_type = prob_type

If usePLA = False Then  'Split time series into equal length segments
    n_segment = 0
    ReDim segment_index(1 To n_raw)
    For i = 1 To n_raw Step seg_len
        n = seg_len
        If (i + n - 1) > n_raw Then n = (n_raw - i + 1)
        n_segment = n_segment + 1
        For j = 1 To n
            segment_index(i + j - 1) = n_segment
        Next j
    Next i
Else    'Split time series use PLA with mininum segment length
    ReDim xArr(1 To n_raw)
    xArr(1) = x(1, 1)
    For i = 2 To n_raw
        xArr(i) = xArr(i - 1) + x(i, 1)
    Next i
    Call mPLA.BottomUp(xArr, n_segment, Int(n_raw / seg_len), seg_len, segment_index)
End If

Call Find_segment_mean_covar(x, segment_index, n_segment, segment_means, segment_covars, segment_len)

'=== Initialization with k-means++ scheme
Randomize
'Randomly pick first center
ReDim pcluster_means(1 To 1)
ReDim pcluster_covars(1 To 1)
k = Int(Rnd() * n_segment) + 1
pcluster_means(1) = segment_means(k)
pcluster_covars(1) = segment_covars(k)
'Select next center base on distances from existing centers
For n = 2 To n_cluster
    tmp_x = 0
    ReDim xArr(1 To n_segment)
    For i = 1 To n_segment
        segment_mean = segment_means(i)
        segment_covar = segment_covars(i)
        xArr(i) = Dist2Centers(segment_mean, segment_covar, pcluster_means, pcluster_covars, k)
        tmp_x = tmp_x + xArr(i)
    Next i
    For i = 1 To n_segment
        xArr(i) = xArr(i) / tmp_x
    Next i
    ReDim Preserve pcluster_means(1 To n)
    ReDim Preserve pcluster_covars(1 To n)
    i = modMath.Random_Integer_Prob(xArr)
    pcluster_means(n) = segment_means(i)
    pcluster_covars(n) = segment_covars(i)
Next n
'=========================================

ReDim iArr(1 To n_segment)
For iterate = 1 To iter_max
    
    If iterate Mod 20 = 0 Then
        DoEvents
        Application.StatusBar = "Segmental k-mean: " & iterate & "/" & iter_max
    End If
    
    'Find out the closest cluster to each segment
    ReDim cluster_index(1 To n_segment)
    For i = 1 To n_segment
        segment_mean = segment_means(i)
        segment_covar = segment_covars(i)
        tmp_min = Dist2Centers(segment_mean, segment_covar, pcluster_means, pcluster_covars, n)
        cluster_index(i) = n
    Next i
    
    'Recalculate mean and covariance base on cluster assignment
    For k = 1 To n_cluster
        ReDim cluster_mean(1 To n_dimension)
        ReDim cluster_covar(1 To n_dimension, 1 To n_dimension)
        
        n = 0
        For i = 1 To n_segment
            If cluster_index(i) = k Then
                n = n + segment_len(i)
                segment_mean = segment_means(i)
                segment_covar = segment_covars(i)
                For j = 1 To n_dimension
                    cluster_mean(j) = cluster_mean(j) + segment_mean(j) * segment_len(i)
                Next j
                Call covar_update(segment_covar, segment_mean, cluster_covar, segment_len(i))
            End If
        Next i
        For j = 1 To n_dimension
            cluster_mean(j) = cluster_mean(j) / n
            cluster_covar(j, j) = cluster_covar(j, j) / n - cluster_mean(j) * cluster_mean(j)
        Next j
        For i = 1 To n_dimension - 1
            For j = i + 1 To n_dimension
                cluster_covar(i, j) = cluster_covar(i, j) / n - cluster_mean(i) * cluster_mean(j)
                cluster_covar(j, i) = cluster_covar(i, j)
            Next j
        Next i
        pcluster_means(k) = cluster_mean
        pcluster_covars(k) = cluster_covar
        
    Next k
    
    'Terminate on convergence
    n = 0
    For i = 1 To n_segment
        If cluster_index(i) = iArr(i) Then n = n + 1
    Next i
    If n = n_segment Then
        Debug.Print "Segmental k-mean: Convergence achieved at " & iterate & "/" & iter_max
        Exit For
    End If
    iArr = cluster_index
    
Next iterate

'Rearrange cluster index by first dimension in ascending order
ReDim xArr(1 To n_cluster)
For k = 1 To n_cluster
    xArr(k) = pcluster_means(k)(1)
Next k
Call modMath.Sort_Quick_A(xArr, 1, n_cluster, iArr)
For i = 1 To n_segment
    For j = 1 To n_cluster
        If iArr(j) = cluster_index(i) Then
            cluster_index(i) = j
            Exit For
        End If
    Next j
Next i
vArr = pcluster_means
For j = 1 To n_cluster
    pcluster_means(j) = vArr(iArr(j))
Next j
vArr = pcluster_covars
For j = 1 To n_cluster
    pcluster_covars(j) = vArr(iArr(j))
Next j
Erase vArr

'Assign cluster index to each observation in original series
ReDim pcluster_index(1 To n_raw)
ReDim pcluster_size(1 To n_segment)
For i = 1 To n_raw
    j = cluster_index(segment_index(i))
    pcluster_index(i) = j
    pcluster_size(j) = pcluster_size(j) + 1
Next i
Erase iArr, segment_index, xArr, segment_covar, segment_mean, segment_covars, segment_means
Application.StatusBar = False
End Sub


Private Sub Find_segment_mean_covar(x() As Double, segment_index() As Long, n_segment As Long, _
    segment_means As Variant, segment_covars As Variant, segment_len() As Long)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, jj As Long
Dim n_raw As Long, n_dimension As Long
Dim segment_mean() As Double, segment_covar() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim segment_means(1 To n_segment)
    ReDim segment_covars(1 To n_segment)
    ReDim segment_len(1 To n_segment)
    For k = 1 To n_segment
        ReDim segment_mean(1 To n_dimension)
        ReDim segment_covar(1 To n_dimension, 1 To n_dimension)
        n = 0
        For i = 1 To n_raw
            If segment_index(i) = k Then
                segment_len(k) = segment_len(k) + 1
                n = n + 1
                For j = 1 To n_dimension
                    segment_mean(j) = segment_mean(j) + x(i, j)
                    segment_covar(j, j) = segment_covar(j, j) + x(i, j) ^ 2
                    For jj = j + 1 To n_dimension
                        segment_covar(j, jj) = segment_covar(j, jj) + x(i, j) * x(i, jj)
                    Next jj
                Next j
                
            End If
        Next i
        For j = 1 To n_dimension
            segment_mean(j) = segment_mean(j) / n
            segment_covar(j, j) = segment_covar(j, j) / n - segment_mean(j) ^ 2
        Next j
        For j = 1 To n_dimension - 1
            For jj = j + 1 To n_dimension
                segment_covar(j, jj) = segment_covar(j, jj) / n - segment_mean(j) * segment_mean(jj)
                segment_covar(jj, j) = segment_covar(j, jj)
            Next jj
        Next j
        segment_means(k) = segment_mean
        segment_covars(k) = segment_covar
    Next k
    Erase segment_mean, segment_covar
End Sub


Private Sub covar_update(x_covar() As Double, x_mean() As Double, cluster_covar() As Double, segment_len As Long)
Dim i As Long, j As Long, n As Long
    n = UBound(x_mean, 1)
    'For i = 1 To n
    '    For j = 1 To n
    '        cluster_covar(i, j) = cluster_covar(i, j) + x_covar(i, j) + x_mean(i) * x_mean(j) - cluster_mean(i) * cluster_mean(j)
    '    Next j
    'Next i
    For i = 1 To n
        For j = 1 To n
            cluster_covar(i, j) = cluster_covar(i, j) + (x_covar(i, j) + x_mean(i) * x_mean(j)) * segment_len
        Next j
    Next i
End Sub

'=== "Renyi Divergence Measures for Commonly Used Univariate Continuous Distributions"
'=== M. Gil, F. Alajaji (2013)
'=== D(Px||Py)
Private Function KL_Divergence(x_mean() As Double, x_covar() As Double, y_mean() As Double, y_covar() As Double) As Double
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double, lambda_x As Double, lambda_y As Double
Dim y_covar_i() As Double
    n_dimension = UBound(x_mean, 1)
    
    If pprob_type = "GAUSSIAN" Then
    
        y_covar_i = modMath.Matrix_Inverse(y_covar)
        tmp_x = Log(modMath.LUPDeterminant(y_covar) / modMath.LUPDeterminant(x_covar))
        For i = 1 To n_dimension
            For j = 1 To n_dimension
                tmp_x = tmp_x + (y_mean(i) - x_mean(i)) * y_covar_i(i, j) * (y_mean(j) - x_mean(j)) _
                    + y_covar_i(i, j) * x_covar(j, i)
            Next j
        Next i
        KL_Divergence = (tmp_x - n_dimension) / 2
        
    ElseIf pprob_type = "LAPLACE" Then
        
        tmp_x = Abs(x_mean(1) - y_mean(1))
        lambda_x = Sqr(x_covar(1, 1) / 2)
        lambda_y = Sqr(y_covar(1, 1) / 2)
        KL_Divergence = Log(lambda_y / lambda_x) + tmp_x / lambda_y + (lambda_x / lambda_y) * Exp(-tmp_x / lambda_x) - 1
    
    Else
        Debug.Print "ckMeanSegmental: KL_Divergence: prob dist not implemented."
    End If
End Function


Private Sub Find_mean_covar(x() As Double, x_mean() As Double, x_covar() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim tmp_x As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim x_mean(1 To n_dimension)
    ReDim x_covar(1 To n_dimension, 1 To n_dimension)
    For k = 1 To n_dimension
        For i = 1 To n_raw
            x_mean(k) = x_mean(k) + x(i, k)
        Next i
        x_mean(k) = x_mean(k) / n_raw
        
        tmp_x = 0
        For i = 1 To n_raw
            tmp_x = tmp_x + (x(i, k) - x_mean(k)) ^ 2
        Next i
        x_covar(k, k) = tmp_x / n_raw
    Next k
    For m = 1 To n_dimension - 1
        For n = m + 1 To n_dimension
            tmp_x = 0
            For i = 1 To n_raw
                tmp_x = tmp_x + (x(i, m) - x_mean(m)) * (x(i, n) - x_mean(n))
            Next i
            x_covar(m, n) = tmp_x / n_raw
            x_covar(n, m) = x_covar(m, n)
        Next n
    Next m
End Sub



'K-L Divergence of a segment from closest cluster center
'k is return as the cluster index that is the closest cluster
Private Function Dist2Centers(segment_mean() As Double, segment_covar() As Double, cluster_means As Variant, cluster_covars As Variant, k As Long) As Double
Dim i As Long
Dim tmp_x As Double, tmp_min As Double
Dim cluster_mean() As Double, cluster_covar() As Double
    tmp_min = Exp(70)
    For i = 1 To UBound(cluster_means)
        cluster_mean = cluster_means(i)
        cluster_covar = cluster_covars(i)
        tmp_x = KL_Divergence(segment_mean, segment_covar, cluster_mean, cluster_covar)
        If tmp_x < tmp_min Then
            tmp_min = tmp_x
            k = i
        End If
    Next i
    Dist2Centers = tmp_min
End Function
