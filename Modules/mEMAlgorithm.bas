Attribute VB_Name = "mEMAlgorithm"
Option Explicit
'Requires: modMath, ckMeanCluster


'After getting the mixture parameters, one can now input data x()
'and find the probability density of each data p()
'Input: x(1 to N, 1 to dimension)
'Output: p(1 to N)
Function get_prob(x() As Double, mix_wgts As Variant, _
     x_means As Variant, x_covars As Variant, Optional dist_type As String = "GAUSSIAN") As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long, n_mixture As Long
Dim tmp_x As Double, tmp_y As Double
Dim w As Double, x_mean() As Double, x_covar() As Double
Dim p() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    n_mixture = UBound(mix_wgts, 1)
    ReDim p(1 To n_raw)
    For m = 1 To n_mixture
        w = mix_wgts(m)
        x_mean = x_means(m)
        x_covar = x_covars(m)
        
        If dist_type = "GAUSSIAN" Then
        
            tmp_y = w / Sqr((6.28318530717959 ^ n_dimension) * modMath.LUPDeterminant(x_covar))
            x_covar = modMath.Matrix_Inverse(x_covar)
            For i = 1 To n_raw
                tmp_x = 0
                For k = 1 To n_dimension
                    For n = 1 To n_dimension
                        tmp_x = tmp_x + (x(i, k) - x_mean(k)) * x_covar(k, n) * (x(i, n) - x_mean(n))
                    Next n
                Next k
                p(i) = p(i) + Exp(-0.5 * tmp_x) * tmp_y
            Next i
        
        ElseIf dist_type = "LAPLACE" Then

            For k = 1 To n_dimension
                tmp_x = w / Sqr(2 * x_covar(k, k))
                tmp_y = Sqr(x_covar(k, k) / 2)
                For i = 1 To n_raw
                    p(i) = p(i) + Exp(-Abs(x(i, k) - x_mean(k)) / tmp_y) * tmp_x
                Next i
            Next k

        End If
    Next m
    get_prob = p
    Erase x_mean, x_covar, p
End Function

'Input: x(1 to N, 1 to D), D-dimension data x
'Input: n_mixture, desired number of mixtures
'Input: dist_type, GAUSSIAN or LAPLACE
'Output: mix_wgts, array holding mixing weights
'Output: x_means, jagged array where each element is a D-dimensional vector
'Output: x_covars, jagged array where each element is a DxD dimensional covariance matrix
Sub Mixture(x() As Double, n_mixture As Long, mix_wgts As Variant, x_means As Variant, x_covars As Variant, _
        Optional iter_max As Long = 1000, Optional likelihood As Variant, _
        Optional init_by_kMeans As Long = 0, Optional dist_type As String = "GAUSSIAN")
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, det As Double
Dim w() As Double, x_mean() As Double, x_covar() As Double
Dim sort_index() As Long, mix_wgts_old As Variant, x_means_old As Variant, x_covars_old As Variant
Dim conv_count As Long, conv_chk As Double, conv_chk_prev As Double

    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim mix_wgts(1 To n_mixture)
    ReDim x_means(1 To n_mixture)
    ReDim x_covars(1 To n_mixture)
    
    If dist_type = "LAPLACE" And n_dimension > 1 Then
        Debug.Print "Laplace distribution only supports 1-dimensional data for now."
        Exit Sub
    End If
    
    If init_by_kMeans = 1 Then
        Call Init_kMean(x, n_mixture, mix_wgts, x_means, x_covars)
    Else
        Call Init_Random(x, n_mixture, mix_wgts, x_means, x_covars)
    End If
    
    If IsMissing(likelihood) = False Then ReDim likelihood(1 To iter_max)
    
    conv_chk_prev = -Exp(70)
    For iterate = 1 To iter_max
        
        If iterate Mod 50 = 0 Then
            DoEvents
            Application.StatusBar = "EM Algorithm: " & iterate & "/" & iter_max
        End If
        
        ReDim w(1 To n_raw, 1 To n_mixture)
        If dist_type = "GAUSSIAN" Then
            For k = 1 To n_mixture
                x_covar = x_covars(k)
                x_mean = x_means(k)
                det = Sqr((6.28318530717959 ^ n_dimension) * modMath.LUPDeterminant(x_covar))
                x_covar = modMath.Matrix_Inverse(x_covar)
                For i = 1 To n_raw
                    tmp_x = 0
                    For m = 1 To n_dimension
                        For n = 1 To n_dimension
                            tmp_x = tmp_x + (x(i, m) - x_mean(m)) * x_covar(m, n) * (x(i, n) - x_mean(n))
                        Next n
                    Next m
                    w(i, k) = Exp(-0.5 * tmp_x) / det
                Next i
            Next k
        ElseIf dist_type = "LAPLACE" Then
            For k = 1 To n_mixture
                x_covar = x_covars(k)
                x_mean = x_means(k)
                det = Sqr(2 * x_covar(1, 1))
                For i = 1 To n_raw
                    w(i, k) = Exp(-Sqr(2 / x_covar(1, 1)) * Abs(x(i, 1) - x_mean(1))) / det
                Next i
            Next k
        Else
            Debug.Print "Mixture: dist_type not defined."
        End If
        
        conv_chk = 0
        For i = 1 To n_raw
            tmp_x = 0
            For k = 1 To n_mixture
                tmp_x = tmp_x + w(i, k) * mix_wgts(k)
            Next k
            conv_chk = conv_chk + Log(tmp_x)
            For k = 1 To n_mixture
                w(i, k) = w(i, k) * mix_wgts(k) / tmp_x
            Next k
        Next i
        
        If IsMissing(likelihood) = False Then likelihood(iterate) = likelihood(iterate) + conv_chk / n_raw
        
        
        For k = 1 To n_mixture
            tmp_x = 0
            For i = 1 To n_raw
                tmp_x = tmp_x + w(i, k)
            Next i
            mix_wgts(k) = tmp_x / n_raw
            
            ReDim x_mean(1 To n_dimension)
            For n = 1 To n_dimension
                tmp_x = 0
                For i = 1 To n_raw
                    tmp_x = tmp_x + w(i, k) * x(i, n)
                Next i
                x_mean(n) = tmp_x / (mix_wgts(k) * n_raw)
            Next n
            x_means(k) = x_mean
        
            ReDim x_covar(1 To n_dimension, 1 To n_dimension)
            For m = 1 To n_dimension
                tmp_x = 0
                For i = 1 To n_raw
                    tmp_x = tmp_x + ((x(i, m) - x_mean(m)) ^ 2) * w(i, k)
                Next i
                x_covar(m, m) = tmp_x / (mix_wgts(k) * n_raw)
            Next m
            For m = 1 To n_dimension - 1
                For n = m + 1 To n_dimension
                    tmp_x = 0
                    For i = 1 To n_raw
                        tmp_x = tmp_x + (x(i, m) - x_mean(m)) * (x(i, n) - x_mean(n)) * w(i, k)
                    Next i
                    x_covar(m, n) = tmp_x / (mix_wgts(k) * n_raw)
                    x_covar(n, m) = x_covar(m, n)
                Next n
            Next m
            
            x_covars(k) = x_covar
        Next k
        
        
        If conv_chk >= conv_chk_prev Then
            conv_count = conv_count + 1
            If conv_count >= 100 And (conv_chk - conv_chk_prev) < 0.00000001 Then
                Exit For
            End If
        Else
            conv_count = 0
        End If
        conv_chk_prev = conv_chk
        
        
    Next iterate
    
    If iterate > iter_max Then iterate = iter_max
    If IsMissing(likelihood) = False Then ReDim Preserve likelihood(1 To iterate)
    
    'Sort mixtures by first dimension
    ReDim x_mean(1 To n_mixture)
    For i = 1 To n_mixture
        x_mean(i) = x_means(i)(1)
    Next i
    Call modMath.Sort_Quick_A(x_mean, 1, n_mixture, sort_index)
    mix_wgts_old = mix_wgts
    x_means_old = x_means
    x_covars_old = x_covars
    For i = 1 To n_mixture
        j = sort_index(i)
        tmp_x = mix_wgts_old(j)
        x_mean = x_means_old(j)
        x_covar = x_covars_old(j)
        mix_wgts(i) = tmp_x
        x_means(i) = x_mean
        x_covars(i) = x_covar
    Next i
    Erase mix_wgts_old, x_means_old, x_covars_old, x_mean, x_covar, sort_index
    
Application.StatusBar = False
End Sub



Private Sub Init_Random(x() As Double, n_mixture As Long, mix_wgts As Variant, x_means As Variant, x_covars As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim x_tmp() As Double
Dim x_mean() As Double, x_covar() As Double

    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim mix_wgts(1 To n_mixture)
    ReDim x_means(1 To n_mixture)
    ReDim x_covars(1 To n_mixture)
    
    Randomize
    n = Int(n_raw / n_mixture)
    For k = 1 To n_mixture
        mix_wgts(k) = 1# / n_mixture
        
        ReDim x_tmp(1 To n, 1 To n_dimension)
        ReDim x_mean(1 To n_dimension)
        For i = 1 To n
            j = Rnd() * n_raw + 1
            For m = 1 To n_dimension
                x_tmp(i, m) = x(j, m)
                x_mean(m) = x_mean(m) + x(j, m)
            Next m
        Next i
        For m = 1 To n_dimension
            x_mean(m) = x_mean(m) / n
        Next m
        x_means(k) = x_mean
    
        x_covar = modMath.Covariance_Matrix(x_tmp)
        x_covars(k) = x_covar
    Next k
    Erase x_tmp
End Sub


'Initialize by k-means clustering
Private Sub Init_kMean(x() As Double, n_mixture As Long, mix_wgts As Variant, x_means As Variant, x_covars As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim iArr() As Long, jArr() As Long
Dim x_tmp() As Double
Dim x_mean() As Double, x_covar() As Double
Dim kM1 As ckMeanCluster

    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim mix_wgts(1 To n_mixture)
    ReDim x_means(1 To n_mixture)
    ReDim x_covars(1 To n_mixture)
    
    Set kM1 = New ckMeanCluster
    With kM1
        Call .kMean_Clustering(x, n_mixture)
        iArr = .x_cluster
        jArr = .cluster_size
        For k = 1 To n_mixture
            mix_wgts(k) = jArr(k) / n_raw

            x_tmp = .cluster_mean
            ReDim x_mean(1 To n_dimension)
            For m = 1 To n_dimension
                x_mean(m) = x_tmp(k, m)
            Next m

            ReDim x_tmp(1 To jArr(k), 1 To n_dimension)
            j = 0
            For i = 1 To n_raw
                If iArr(i) = k Then
                    j = j + 1
                    For m = 1 To n_dimension
                        x_tmp(j, m) = x(i, m)
                    Next m
                End If
                If j = jArr(k) Then Exit For
            Next i
            x_covar = modMath.Covariance_Matrix(x_tmp)

            x_means(k) = x_mean
            x_covars(k) = x_covar
        Next k
        Call .Reset
    End With
    Set kM1 = Nothing
End Sub
