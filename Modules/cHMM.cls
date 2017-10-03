VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHMM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Requires: modmath, ckMeanCluster, ckMeanSegmental, mEMAlgorithm

Private pn_state As Long
Private pn_obs As Long
Private pn_mixture As Long
Private pstate_labels() As String
Private pobs_labels() As String
Private pstart_probs() As Double    '(1:N_state)
Private ptransitions() As Double    '(1:N_state, 1:N_state)
Private pn_dimension As Long
Private pmix_wgt() As Double    '(1:N_state, 1:N_mixture)
Private pmix_mean() As Double   '(1:N_state, 1:N_mixture, 1:D)
Private pmix_var() As Double    '(1:N_state, 1:N_mixture, 1:D, 1:D)
Private pcost_function() As Double
Private pprob_type As String '"GAUSSIAN" or "LAPLACE"

Public Property Get dimension() As Long
    dimension = pn_dimension
End Property

Public Property Get cost_function() As Double()
    cost_function = pcost_function
End Property

Public Property Get n_state() As Long
    n_state = pn_state
End Property

Public Property Get n_mixture() As Long
    n_mixture = pn_mixture
End Property

Sub Reset()
    Erase pstate_labels, pobs_labels
    Erase pstart_probs, ptransitions
    Erase pmix_wgt, pmix_mean, pmix_var
    Erase pcost_function
End Sub

Sub Init(n_state As Long, n_mixture As Long, n_dimension As Long, Optional prob_type As String = "GAUSSIAN")
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double
    
    pprob_type = VBA.UCase(prob_type)
    pn_state = n_state
    pn_mixture = n_mixture
    pn_dimension = n_dimension
    
    ReDim pstate_labels(1 To n_state)
    ReDim pstart_probs(1 To n_state)
    ReDim ptransitions(1 To n_state, 1 To n_state)
    ReDim pmix_wgt(1 To n_state, 1 To n_mixture)
    ReDim pmix_mean(1 To n_state, 1 To n_mixture, 1 To n_dimension)
    ReDim pmix_var(1 To n_state, 1 To n_mixture, 1 To n_dimension, 1 To n_dimension)
    
    Randomize
    For i = 1 To pn_state
        pstart_probs(i) = 1# / pn_state
        tmp_x = 0
        For j = 1 To pn_state
            ptransitions(i, j) = Rnd() + 0.01
            tmp_x = tmp_x + ptransitions(i, j)
        Next j
        For j = 1 To pn_state
            ptransitions(i, j) = ptransitions(i, j) / tmp_x
        Next j
    Next i
    
    Randomize
    For i = 1 To pn_state
        tmp_x = 0
        For k = 1 To pn_mixture
            pmix_wgt(i, k) = 1 / pn_mixture
            For m = 1 To n_dimension
                pmix_mean(i, k, m) = -0.5 + Rnd()
                pmix_var(i, k, m, m) = 0.02 + (-0.01 + Rnd() * 0.02)
            Next m
            For m = 1 To n_dimension - 1
                For n = m + 1 To n_dimension
                    pmix_var(i, k, m, n) = 0.002 + (-0.001 + Rnd() * 0.002)
                    pmix_var(i, k, n, m) = pmix_var(i, k, m, n)
                Next n
            Next m
        Next k
    Next i
    
End Sub




'Replace zeroes in probablilities by small random values
Private Sub remove_zeroes()
Dim i As Long, j As Long, k As Long
Dim tmp_x As Double, tmp_y As Double
    Randomize
    For k = 1 To pn_state
        If pstart_probs(k) = 0 Then pstart_probs(k) = 0.001 + Rnd() * 0.001
        tmp_x = tmp_x + pstart_probs(k)
        
        tmp_y = 0
        For j = 1 To pn_state
            If ptransitions(k, j) = 0 Then ptransitions(k, j) = 0.001 + Rnd() * 0.001
            tmp_y = tmp_y + ptransitions(k, j)
        Next j
        For j = 1 To pn_state
            ptransitions(k, j) = ptransitions(k, j) / tmp_y
        Next j
        
        tmp_y = 0
        For j = 1 To pn_mixture
            If pmix_wgt(k, j) = 0 Then pmix_wgt(k, j) = 0.001 + Rnd() * 0.001
            tmp_y = tmp_y + pmix_wgt(k, j)
        Next j
        For j = 1 To pn_mixture
            pmix_wgt(k, j) = pmix_wgt(k, j) / tmp_y
        Next j
    Next k
    For k = 1 To pn_state
        pstart_probs(k) = pstart_probs(k) / tmp_x
    Next k
End Sub

Public Property Let mix_wgt(x() As Double)
    pmix_wgt = x
End Property

Public Property Let mix_mean(x() As Double)
    pmix_mean = x
End Property

Public Property Let mix_var(x() As Double)
    pmix_var = x
End Property

Public Property Let start_prob(i As Long, x As Double)
    pstart_probs(i) = x
End Property

Public Property Let start_probs(x() As Double)
Dim i As Long
Dim tmp_x As Double
    pstart_probs = x
    tmp_x = 0
    Randomize
    For i = 1 To pn_state
        If pstart_probs(i) = 0.01 Then pstart_probs(i) = 0.001 + Rnd() * 0.01
        tmp_x = tmp_x + pstart_probs(i)
    Next i
    For i = 1 To pn_state
        pstart_probs(i) = pstart_probs(i) / tmp_x
    Next i
End Property

Public Property Get start_probs() As Double()
    start_probs = pstart_probs
End Property

Public Property Let transition(i As Long, j As Long, x As Double)
    ptransitions(i, j) = x
End Property

Public Property Let transitions(x() As Double)
    Dim i As Long, j As Long
    Dim tmp_x As Double
    ptransitions = x
    tmp_x = 0
    For i = 1 To pn_state
        Randomize
        tmp_x = 0
        For j = 1 To pn_state
            If ptransitions(i, j) = 0 Then ptransitions(i, j) = 0.001 + Rnd() * 0.01
            tmp_x = tmp_x + ptransitions(i, j)
        Next j
        For j = 1 To pn_state
            ptransitions(i, j) = ptransitions(i, j) / tmp_x
        Next j
    Next i
End Property

Public Property Get transitions() As Double()
    transitions = ptransitions
End Property


'Public Property Let state_label(i As Long, x As String)
'pstate_labels(i) = x
'End Property
'
'Public Property Get state_label(i As Long) As String
'state_label = pstate_labels(i)
'End Property
'
'Public Property Let obs_label(i As Long, x As String)
'pobs_labels(i) = x
'End Property
'
'Public Property Get obs_label(i As Long) As String
'obs_label = pobs_labels(i)
'End Property

'Public Property Let emission(i As Long, j As Long, x As Double)
'pemissions(i, j) = x
'End Property

'Public Property Let emissions(x() As Double)
'pemissions() = x
'End Property
'
'Public Property Get emissions() As Double()
'emissions = pemissions
'End Property
'
'Sub gen_next_state(s As Long, s_next As Long)
'    Dim p() As Double
'    Call get_vector(ptransitions, s, 1, p)
'    s_next = Random_Integer_Prob(p)
'End Sub
'
'Sub gen_obs(s As Long, k_obs As Long)
'    Dim p() As Double
'    Call get_vector(pemissions, s, 1, p)
'    k_obs = Random_Integer_Prob(p)
'End Sub
'
'Sub gen_sequence(n_T As Long, s() As Long, x() As Long)
'Dim t As Long
'ReDim s(1 To n_T)
'ReDim x(1 To n_T)
's(1) = Random_Integer_Prob(pstart_probs)
'Call Me.gen_obs(s(1), x(1))
'
'For t = 2 To n_T
'    Call Me.gen_next_state(s(t - 1), s(t))
'    Call Me.gen_obs(s(t), x(t))
'Next t
'End Sub

'Initialize HMM base on segmental k-means on selected dimension
'Input: x(1 to n, 1 to n_dimension)
Sub Init_SegmentalK(x() As Double, Optional tgt_dimension As Long = 1, _
        Optional seg_len As Long = 15, Optional usePLA As Boolean = False)
Dim t As Long, s As Long, n_T As Long
Dim x_sub() As Double, q() As Long
Dim kS1 As ckMeanSegmental
    n_T = UBound(x, 1)
    
    'extract selected dimension
    ReDim x_sub(1 To n_T, 1 To 1)
    For t = 1 To n_T
        x_sub(t, 1) = x(t, tgt_dimension)
    Next t
    
    'run segmental k-mean
    Set kS1 = New ckMeanSegmental
    With kS1
        Call .Segmentation(x_sub, pn_state, seg_len, , pprob_type, usePLA)
        q = .cluster_index
        Call .Reset
    End With
    Set kS1 = Nothing
    
    Call Init_State_Index(x, q)
    Call remove_zeroes
    Erase x_sub, q
End Sub


'Initialize HMM base on k-means clustering
'Input: x(1 to n, 1 to n_dimension)
Sub Init_k_means(x() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, s As Long
Dim x_norm() As Double, x_avg() As Double, x_sd() As Double
Dim state_index() As Long, sort_index() As Long
Dim ckMean1 As ckMeanCluster
    
    n = UBound(x, 1)
    x_norm = x
    Call normalize_data(x_norm, x_avg, x_sd)
    
    Set ckMean1 = New ckMeanCluster
    With ckMean1
        Call .kMean_Clustering(x_norm, pn_state)
        x_norm = .cluster_mean
        state_index = .x_cluster
        Call .Reset
    End With
    Set ckMean1 = Nothing
    
    ReDim x_avg(1 To pn_state)
    For k = 1 To pn_state
        x_avg(k) = x_norm(k, 1)
    Next k
    Call modMath.Sort_Quick_A(x_avg, 1, pn_state, sort_index)
    For i = 1 To n
        j = state_index(i)
        For k = 1 To pn_state
            If sort_index(k) = j Then
                state_index(i) = pn_state - k + 1
                Exit For
            End If
        Next k
    Next i
    
    Call Init_State_Index(x, state_index)
    Call remove_zeroes
    Erase x_norm, x_avg, x_sd, state_index, sort_index
End Sub


'Initialize HMM based on an intial guss of state sequence
Private Sub Init_State_Index(x() As Double, state_index() As Long)
Dim i As Long, k As Long, s As Long, n_T As Long
Dim vArr As Variant, x_sub() As Double
    n_T = UBound(state_index, 1)
    
    For i = 1 To n_T
        s = state_index(i)
        pstart_probs(s) = pstart_probs(s) + 1
        If i < n_T Then
            k = state_index(i + 1)
            ptransitions(s, k) = ptransitions(s, k) + 1
        End If
    Next i
    
    For s = 1 To pn_state
        For k = 1 To pn_state
            If ptransitions(s, k) > 0 Then ptransitions(s, k) = ptransitions(s, k) / pstart_probs(s)
        Next k
        pstart_probs(s) = pstart_probs(s) / n_T
    Next s
    
    'For each state, fit a Gaussian mixtures
    vArr = Divide_Set(x, state_index)
    For s = 1 To pn_state
        x_sub = vArr(s)
        If UBound(x_sub, 1) > 0 Then Call Fit_Mixture(x_sub, s)
    Next s
    
    Erase vArr, x_sub
End Sub


'Fit mixture model to x() using EM-algorithm
Private Sub Fit_Mixture(x() As Double, s As Long)
Dim i As Long, j As Long, m As Long
Dim x_avg() As Double, x_sd() As Double
Dim mix_wgts As Variant, x_means As Variant, x_covars As Variant, x_mean() As Double, x_covar() As Double
    Call normalize_data(x, x_avg, x_sd)
    Call mEMAlgorithm.Mixture(x, pn_mixture, mix_wgts, x_means, x_covars, , , 1, pprob_type)
    For m = 1 To pn_mixture
        pmix_wgt(s, m) = mix_wgts(m)
        x_mean = x_means(m)
        x_covar = x_covars(m)
        For i = 1 To pn_dimension
            pmix_mean(s, m, i) = x_mean(i) * x_sd(i) + x_avg(i)
            For j = 1 To pn_dimension
                pmix_var(s, m, i, j) = x_covar(i, j) * x_sd(i) * x_sd(j)
            Next j
        Next i
    Next m
    Erase x_avg, x_sd, x_mean, x_covar, x_means, x_covars
End Sub


Private Sub x_mean_covar(x() As Double, x_mean() As Double, x_covar() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim tmp_x As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim x_mean(1 To n_dimension)
    ReDim x_covar(1 To n_dimension, 1 To n_dimension)
    
    For m = 1 To n_dimension
        For i = 1 To n_raw
            x_mean(m) = x_mean(m) + x(i, m)
        Next i
        x_mean(m) = x_mean(m) / n_raw
        tmp_x = 0
        For i = 1 To n_raw
            tmp_x = tmp_x + (x(i, m) - x_mean(m)) ^ 2
        Next i
        x_covar(m, m) = tmp_x / (n_raw - 1)
    Next m
    
    For m = 1 To n_dimension - 1
        For n = m + 1 To n_dimension
            tmp_x = 0
            For i = 1 To n_raw
                tmp_x = tmp_x + (x(i, m) - x_mean(m)) * (x(i, n) - x_mean(n))
            Next i
            x_covar(m, n) = tmp_x / (n_raw - 1)
            x_covar(n, m) = x_covar(m, n)
        Next n
    Next m
End Sub


Private Function Divide_Set(x() As Double, set_idx() As Long) As Variant
Dim i As Long, j As Long, k As Long, n As Long, s As Long
Dim n_raw As Long, n_dimension As Long, n_set As Long
Dim set_size() As Long, iArr() As Long
Dim vArr As Variant, xArr() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    n_set = 0
    ReDim set_size(1 To n_raw)
    For i = 1 To n_raw
        j = set_idx(i)
        If j > n_set Then n_set = j
        set_size(j) = set_size(j) + 1
    Next i
    ReDim Preserve set_size(1 To n_set)
    
    ReDim vArr(1 To n_set)
    For s = 1 To n_set
        If set_size(s) = 0 Then
            ReDim xArr(0 To 0, 1 To n_dimension)
        Else
            ReDim xArr(1 To set_size(s), 1 To n_dimension)
        End If
        vArr(s) = xArr
    Next s
    
    ReDim iArr(1 To n_set)
    For i = 1 To n_raw
        s = set_idx(i)
        xArr = vArr(s)
        iArr(s) = iArr(s) + 1
        For j = 1 To pn_dimension
            xArr(iArr(s), j) = x(i, j)
        Next j
        vArr(s) = xArr
    Next i
    
    Divide_Set = vArr
    Erase set_size, vArr, xArr, iArr
End Function


Sub BaumWelch(x_obs As Variant, Optional iterate_max As Long = 100)
Dim i As Long, j As Long, k As Long, n As Long, t As Long
Dim iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_T As Long, n_raw As Long
Dim alphas() As Variant, betas() As Variant, zS() As Variant
Dim gammas() As Variant, lambdas() As Variant, gamma2s() As Variant
Dim pdfs() As Variant, bs() As Variant
Dim z() As Double, count As Long

n_raw = UBound(x_obs, 1)
ReDim pcost_function(1 To iterate_max)
For iterate = 1 To iterate_max

    DoEvents
    If iterate Mod 5 = 0 Then Application.StatusBar = "cHMM: BaumWelch: " & iterate & "/" & iterate_max
    
    Call remove_zeroes
    
    Call Find_all_probs(x_obs, pdfs)
    Call Find_all_b(pdfs, bs)

    Call FwdProc(alphas, zS, bs)   'Forward Procedure
    Call BkwdProc(betas, zS, bs)   'Backward Procedure

    Call Calc_Prob(gammas, lambdas, alphas, betas, zS, bs) 'Update gamma() and lambda()
    Call Calc_gamma2(gamma2s, alphas, betas, bs, pdfs)

    Call ReEstimate(gammas, lambdas)   'Re-estimate pi(), a(), b()
    Call ReEstimate2(x_obs, gamma2s)  'Re-estimate mixture wgts, means & covariance
    
    For n = 1 To n_raw
        z = zS(n)
        n_T = UBound(z, 1)
        tmp_x = 0
        For t = 1 To n_T
            tmp_x = tmp_x - Log(z(t))
        Next t
        pcost_function(iterate) = pcost_function(iterate) + tmp_x / n_T
    Next n
    pcost_function(iterate) = pcost_function(iterate) / (n_raw * pn_dimension)

Next iterate
Erase alphas, betas, zS
Erase gammas, lambdas, gamma2s, pdfs, bs, z
Application.StatusBar = False
End Sub


Private Sub FwdProc(alphas() As Variant, zS() As Variant, bs() As Variant)
Dim n As Long, n_raw As Long
Dim B() As Double, alpha() As Double, z() As Double
    n_raw = UBound(bs, 1)
    ReDim alphas(1 To n_raw)
    ReDim zS(1 To n_raw)
    For n = 1 To n_raw
        B = bs(n)
        Call FwdProc_Single(alpha, z, B)
        alphas(n) = alpha
        zS(n) = z
    Next n
End Sub


Private Sub BkwdProc(betas() As Variant, zS() As Variant, bs() As Variant)
Dim n As Long, n_raw As Long
Dim beta() As Double, B() As Double, z() As Double
    n_raw = UBound(bs, 1)
    ReDim betas(1 To n_raw)
    For n = 1 To n_raw
        z = zS(n)
        B = bs(n)
        Call BkwdProc_Single(beta, z, B)
        betas(n) = beta
    Next n
End Sub


Private Sub Calc_Prob(gammas() As Variant, Xis() As Variant, alphas() As Variant, betas() As Variant, zS() As Variant, bs() As Variant)
Dim n As Long, n_raw As Long
Dim alpha() As Double, beta() As Double, z() As Double, B() As Double
Dim gamma() As Double, xi() As Double
    n_raw = UBound(bs, 1)
    ReDim gammas(1 To n_raw)
    ReDim Xis(1 To n_raw)
    For n = 1 To n_raw
        alpha = alphas(n)
        beta = betas(n)
        z = zS(n)
        B = bs(n)
        Call Calc_Prob_Single(gamma, xi, alpha, beta, z, B)
        gammas(n) = gamma
        Xis(n) = xi
    Next n
End Sub


Private Sub Calc_gamma2(gammas() As Variant, alphas() As Variant, betas() As Variant, bs() As Variant, pdfs() As Variant)
Dim n As Long, n_raw As Long
Dim gamma() As Double, alpha() As Double, beta() As Double, B() As Double, pdf() As Double
    n_raw = UBound(bs, 1)
    ReDim gammas(1 To n_raw)
    For n = 1 To n_raw
        alpha = alphas(n)
        beta = betas(n)
        B = bs(n)
        pdf = pdfs(n)
        Call Calc_gamma2_single(gamma, alpha, beta, B, pdf)
        gammas(n) = gamma
    Next n
End Sub


Private Sub FwdProc_Single(alpha() As Double, z() As Double, B() As Double)
Dim i As Long, j As Long, t As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim n_T As Long
    n_T = UBound(B, 2)
    ReDim alpha(1 To pn_state, 1 To n_T)
    ReDim z(1 To n_T)
    
    tmp_x = 0
    For i = 1 To pn_state
        alpha(i, 1) = pstart_probs(i) * B(i, 1)
        tmp_x = tmp_x + alpha(i, 1)
    Next i
    z(1) = tmp_x
    
    For i = 1 To pn_state
        alpha(i, 1) = alpha(i, 1) / z(1)
    Next i
    
    For t = 1 To n_T - 1
        tmp_y = 0
        For i = 1 To pn_state
            tmp_x = 0
            For j = 1 To pn_state
                tmp_x = tmp_x + alpha(j, t) * ptransitions(j, i)
            Next j
            alpha(i, t + 1) = B(i, t + 1) * tmp_x
            tmp_y = tmp_y + alpha(i, t + 1)
        Next i
        z(t + 1) = tmp_y
        For i = 1 To pn_state
            alpha(i, t + 1) = alpha(i, t + 1) / tmp_y
        Next i
        
    Next t

End Sub


Private Sub BkwdProc_Single(beta() As Double, z() As Double, B() As Double)
Dim i As Long, j As Long, t As Long, n_T As Long
Dim tmp_x As Double
    n_T = UBound(B, 2)
    ReDim beta(1 To pn_state, 1 To n_T)
          
    For i = 1 To pn_state
        beta(i, n_T) = 1
    Next i
    
    For t = n_T - 1 To 1 Step -1
        For i = 1 To pn_state
            tmp_x = 0
            For j = 1 To pn_state
                tmp_x = tmp_x + ptransitions(i, j) * B(j, t + 1) * beta(j, t + 1)
            Next j
            beta(i, t) = tmp_x / z(t + 1)
        Next i
    Next t
End Sub


Private Sub Calc_Prob_Single(gamma() As Double, xi() As Double, alpha() As Double, beta() As Double, z() As Double, B() As Double)
Dim i As Long, j As Long, t As Long, n_T As Long
    n_T = UBound(B, 2)
    ReDim gamma(1 To pn_state, 1 To n_T)
    ReDim xi(1 To pn_state, 1 To pn_state, 1 To n_T - 1)

    For t = 1 To n_T
        For i = 1 To pn_state
            gamma(i, t) = alpha(i, t) * beta(i, t)
        Next i
    Next t

    For t = 1 To n_T - 1
        For i = 1 To pn_state
            For j = 1 To pn_state
                xi(i, j, t) = alpha(i, t) * ptransitions(i, j) * B(j, t + 1) * beta(j, t + 1) / z(t + 1)
            Next j
        Next i
    Next t
End Sub


Private Sub Calc_gamma2_single(gamma() As Double, alpha() As Double, beta() As Double, B() As Double, pdf() As Double)
Dim i As Long, j As Long, k As Long, t As Long, n_T As Long
Dim tmp_x As Double, tmp_y As Double
    n_T = UBound(B, 2)
    ReDim gamma(1 To pn_state, 1 To pn_mixture, 1 To n_T)
    For t = 1 To n_T
        tmp_x = 0
        For i = 1 To pn_state
            tmp_x = tmp_x + alpha(i, t) * beta(i, t)
            If B(i, t) > 0 Then
                tmp_y = alpha(i, t) * beta(i, t) / B(i, t)
                For k = 1 To pn_mixture
                    gamma(i, k, t) = tmp_y * pmix_wgt(i, k) * pdf(i, k, t)
                Next k
            End If
        Next i
        For i = 1 To pn_state
            For k = 1 To pn_mixture
                gamma(i, k, t) = gamma(i, k, t) / tmp_x
            Next k
        Next i
    Next t
End Sub


Private Sub ReEstimate(gammas() As Variant, Xis() As Variant)
Dim i As Long, j As Long, n As Long, k As Long, t As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_T As Long, n_raw As Long
Dim gamma() As Double, xi() As Double
    n_raw = UBound(Xis, 1)
    For i = 1 To pn_state

        tmp_x = 0
        tmp_y = 0
        For n = 1 To n_raw
            gamma = gammas(n)
            tmp_x = tmp_x + gamma(i, 1)
            n_T = UBound(gamma, 2)
            For t = 1 To n_T - 1
                tmp_y = tmp_y + gamma(i, t)
            Next t
        Next n
        pstart_probs(i) = tmp_x / n_raw
    
        For j = 1 To pn_state
            tmp_x = 0
            For n = 1 To n_raw
                xi = Xis(n)
                n_T = UBound(xi, 3)
                For t = 1 To n_T - 1
                    tmp_x = tmp_x + xi(i, j, t)
                Next t
            Next n
            ptransitions(i, j) = tmp_x / tmp_y
        Next j

    Next i
End Sub


Private Sub ReEstimate2(x_obs As Variant, gammas() As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, t As Long, d As Long, d2 As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, tmp_d As Double
Dim n_T As Long, n_raw As Long
Dim tmp_vec() As Double, tmp_arr() As Double, gamma() As Double, x() As Double
    n_raw = UBound(x_obs, 1)
    For j = 1 To pn_state
        tmp_d = 0
        For k = 1 To pn_mixture
            tmp_x = 0
            tmp_y = 0
            tmp_z = 0
            ReDim tmp_vec(1 To pn_dimension)
            ReDim tmp_arr(1 To pn_dimension, 1 To pn_dimension)
            
            For n = 1 To n_raw
                gamma = gammas(n)
                x = x_obs(n)
                n_T = UBound(x, 1)
                For t = 1 To n_T
                    tmp_x = tmp_x + gamma(j, k, t)
                    For d = 1 To pn_dimension
                        tmp_vec(d) = tmp_vec(d) + gamma(j, k, t) * x(t, d)
                        For d2 = 1 To pn_dimension
                            tmp_arr(d, d2) = tmp_arr(d, d2) + gamma(j, k, t) * (x(t, d) - pmix_mean(j, k, d)) * (x(t, d2) - pmix_mean(j, k, d2))
                        Next d2
                    Next d
                Next t
            Next n
            tmp_d = tmp_d + tmp_x
            pmix_wgt(j, k) = tmp_x
            For d = 1 To pn_dimension
                pmix_mean(j, k, d) = tmp_vec(d) / tmp_x
                For d2 = 1 To pn_dimension
                    pmix_var(j, k, d, d2) = tmp_arr(d, d2) / tmp_x
                Next d2
            Next d
        Next k
        For k = 1 To pn_mixture
            pmix_wgt(j, k) = pmix_wgt(j, k) / tmp_d
        Next k
    Next j
End Sub



Sub Viterbi(x_obs As Variant, qs As Variant, likelihoods() As Double)
Dim i As Long, j As Long, k As Long, n As Long, t As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_T As Long, n_raw As Long
Dim A() As Double, pi() As Double, bs() As Variant, pdfs() As Variant
Dim q() As Long, B() As Double

    n_raw = UBound(x_obs, 1)
    
    'Log Transform probabilities
    pi = pstart_probs
    A = ptransitions
    For i = 1 To pn_state
        pi(i) = Log(pi(i))
        For j = 1 To pn_state
            A(i, j) = Log(A(i, j))
        Next j
    Next i
    
    'Pre-calculate all gaussian functions
    Call Find_all_probs(x_obs, pdfs)
    Call Find_all_b(pdfs, bs)
    Erase pdfs
    
    ReDim qs(1 To n_raw)
    ReDim likelihoods(1 To n_raw)
    For n = 1 To n_raw
        B = bs(n)
        Call Viterbi_Single(B, pi, A, q, tmp_x)
        qs(n) = q
        likelihoods(n) = tmp_x
    Next n
End Sub


Private Sub Viterbi_Single(B() As Double, pi() As Double, A() As Double, q() As Long, p As Double)
Dim i As Long, j As Long, k As Long, n As Long, t As Long, n_T As Long
Dim tmp_x As Double, tmp_y As Double
Dim tmp_max As Double
Dim i_max As Long
Dim delta() As Double
Dim pSi() As Long
Dim LOGZERO As Double

    LOGZERO = -Exp(70)
    n_T = UBound(B, 2)
    
    ReDim delta(1 To pn_state, 1 To n_T)
    ReDim pSi(1 To pn_state, 1 To n_T)
    ReDim q(1 To n_T)
    
    'Initialization
    For i = 1 To pn_state
        If B(i, 1) = 0 Then
            delta(i, 1) = LOGZERO
        Else
            delta(i, 1) = pi(i) + Log(B(i, 1))
        End If
        pSi(i, 1) = 0
    Next i
    
    For t = 2 To n_T
        For j = 1 To pn_state
            tmp_max = delta(1, t - 1) + A(1, j)
            i_max = 1
            For i = 1 To pn_state
                tmp_x = delta(i, t - 1) + A(i, j)
                If tmp_x > tmp_max Then
                    tmp_max = tmp_x
                    i_max = i
                End If
            Next i
            If B(j, t) = 0 Then
                delta(j, t) = LOGZERO
            Else
                delta(j, t) = tmp_max + Log(B(j, t))
            End If
            pSi(j, t) = i_max
        Next j
    Next t
    
    tmp_max = delta(1, n_T)
    i_max = 1
    For i = 1 To pn_state
        If delta(i, n_T) > tmp_max Then
            tmp_max = delta(i, n_T)
            i_max = i
        End If
    Next i
    p = tmp_max / n_T
    q(n_T) = i_max
    For t = n_T - 1 To 1 Step -1
        q(t) = pSi(q(t + 1), t + 1)
    Next t

End Sub

Sub ViterbiTraining(x_obs() As Long, iterate_max As Long, cost_function() As Double)
Dim i As Long, j As Long, k As Long, n As Long, t As Long
Dim iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_T As Long, n_raw As Long
Dim q() As Long
Dim p() As Double

n_T = UBound(x_obs, 1)
n_raw = UBound(x_obs, 2)
ReDim cost_function(1 To iterate_max)

For iterate = 1 To iterate_max

    DoEvents
    If iterate Mod 1 = 0 Then Application.StatusBar = iterate & "/" & iterate_max
    
    Call Me.Viterbi(x_obs, q, p)
    
    tmp_x = 0
    For n = 1 To n_raw
        tmp_x = tmp_x + p(n)
    Next n
    cost_function(iterate) = tmp_x / n_raw
    
    For i = 1 To pn_state
        tmp_x = 0
        For n = 1 To n_raw
            If q(1, n) = i Then tmp_x = tmp_x + 1
        Next n
        pstart_probs(i) = tmp_x / n_raw
    Next i

    For i = 1 To pn_state
        For j = 1 To pn_state
            tmp_x = 0
            tmp_y = 0
            For n = 1 To n_raw
                For t = 1 To n_T - 1
                    If q(t, n) = i Then
                        tmp_y = tmp_y + 1
                        If q(t + 1, n) = j Then tmp_x = tmp_x + 1
                    End If
                Next t
            Next n
            ptransitions(i, j) = tmp_x / tmp_y
        Next j
    Next i
    
    For j = 1 To pn_state
        For k = 1 To pn_obs
            tmp_x = 0
            tmp_y = 0
            For n = 1 To n_raw
                For t = 1 To n_T
                    If q(t, n) = j Then
                        tmp_y = tmp_y + 1
                        If x_obs(t, n) = k Then tmp_x = tmp_x + 1
                    End If
                Next t
            Next n
            pemissions(j, k) = tmp_x / tmp_y
        Next k
    Next j

Next iterate

End Sub


Private Sub Find_all_b(pdfs() As Variant, bs() As Variant)
Dim i As Long, j As Long, k As Long, t As Long, s As Long, n As Long, m As Long
Dim n_T As Long, n_raw As Long
Dim tmp_x As Double, tmp_y As Double
Dim pdf() As Double, B() As Double
n_raw = UBound(pdfs, 1)
ReDim bs(1 To n_raw)
For n = 1 To n_raw
    pdf = pdfs(n)
    n_T = UBound(pdf, 3)
    ReDim B(1 To pn_state, 1 To n_T)
    For j = 1 To pn_state
        For t = 1 To n_T
            tmp_x = 0
            For m = 1 To pn_mixture
                tmp_x = tmp_x + pmix_wgt(j, m) * pdf(j, m, t)
            Next m
            B(j, t) = tmp_x
        Next t
    Next j
    bs(n) = B
Next n
End Sub


Private Sub Find_all_probs(x_obs As Variant, probs() As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, t As Long, d As Long, s As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_T As Long, n_raw As Long
Dim o() As Double, x_mean() As Double, x_covar() As Double, x() As Double, Prob() As Double
Dim covar_det As Double

n_raw = UBound(x_obs, 1)
ReDim o(1 To pn_dimension)
ReDim x_mean(1 To pn_dimension)
ReDim x_covar(1 To pn_dimension, 1 To pn_dimension)
ReDim probs(1 To n_raw)
For n = 1 To n_raw
    x = x_obs(n)
    n_T = UBound(x, 1)
    ReDim Prob(1 To pn_state, 1 To pn_mixture, 1 To n_T)
    For s = 1 To pn_state
        For m = 1 To pn_mixture
    
            For i = 1 To pn_dimension
                x_mean(i) = pmix_mean(s, m, i)
                For j = 1 To pn_dimension
                    x_covar(i, j) = pmix_var(s, m, i, j)
                Next j
            Next i
            
            If pprob_type = "GAUSSIAN" Then
            
                covar_det = Sqr(modMath.LUPDeterminant(x_covar) * (6.283185307 ^ pn_dimension))
                x_covar = modMath.Matrix_Inverse(x_covar)
                For t = 1 To n_T
                    For d = 1 To pn_dimension
                        o(d) = x(t, d)
                    Next d
                    Prob(s, m, t) = Gaussian(o, x_mean, x_covar, covar_det)
                Next t
            
            ElseIf pprob_type = "LAPLACE" Then

                For t = 1 To n_T
                    For d = 1 To pn_dimension
                        o(d) = x(t, d)
                    Next d
                    Prob(s, m, t) = Laplace(o, x_mean, x_covar)
                Next t

            End If
            
        Next m
    Next s
    probs(n) = Prob
Next n
End Sub


Private Function Laplace(x() As Double, x_mean() As Double, x_covar() As Double) As Double
    Laplace = Exp(-Sqr(2 / x_covar(1, 1)) * Abs(x(1) - x_mean(1))) / Sqr(2 * x_covar(1, 1))
End Function

Private Function Gaussian(x() As Double, x_mean() As Double, x_covar_i() As Double, covar_det As Double) As Double
Dim i As Long, j As Long
Dim tmp_x As Double
    'Inverse and determinant of covariance matrix are pre-calculated to save time
    'Dim x_covar_i() As Double, covar_det As Double
    'covar_det = Sqr(modmath.LUPDeterminant(x_covar) * (6.283185307 ^ pn_dimension))
    'x_covar_i = modmath.Matrix_Inverse(x_covar)
    tmp_x = 0
    For i = 1 To pn_dimension
        For j = 1 To pn_dimension
            tmp_x = tmp_x + (x(i) - x_mean(i)) * x_covar_i(i, j) * (x(j) - x_mean(j))
        Next j
    Next i
    Gaussian = Exp(-0.5 * tmp_x) / covar_det
End Function


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


'Input: x(1 to M, 1 to N), 2D Matrix
'Output: y(), vector from the k-th row/column of x()
Private Function get_vector(x() As Double, k As Long, idim As Long) As Double()
Dim i As Long, n As Long
Dim y() As Double
If idim = 1 Then
    n = UBound(x, 2)
    ReDim y(1 To n)
    For i = 1 To n
        y(i) = x(k, i)
    Next i
ElseIf idim = 2 Then
    n = UBound(x, 1)
    ReDim y(1 To n)
    For i = 1 To n
        y(i) = x(i, k)
    Next i
End If
get_vector = y
Erase y
End Function


'# of states
'# of mixtures
'# of dimensions
'probability distribution
'state probability  || state transition matrix
'state              || mixture                      || mixture weight
'mixture means      || mixture covariance matrix
Sub Print_Model(vRng As Range)
Dim i As Long, j As Long, k As Long, m As Long, d As Long, d2 As Long
With vRng
    .Value = pn_state
    .Offset(1, 0).Value = pn_mixture
    .Offset(2, 0).Value = pn_dimension
    .Offset(3, 0).Value = pprob_type
    For i = 1 To pn_state
        .Offset(3 + i, 0).Value = pstart_probs(i)
        For j = 1 To pn_state
            .Offset(3 + i, 1 + j).Value = ptransitions(i, j)
        Next j
    Next i
    
    j = 0
    For k = 1 To pn_state
        For m = 1 To pn_mixture
            j = j + 1
            i = 4 + pn_state + (j - 1) * (pn_dimension + 2)
            .Offset(i, 0).Value = k
            .Offset(i, 1).Value = m
            .Offset(i, 2).Value = pmix_wgt(k, m)
            For d = 1 To pn_dimension
                .Offset(i + d, 0).Value = pmix_mean(k, m, d)
                For d2 = 1 To pn_dimension
                    .Offset(i + d, 1 + d2).Value = pmix_var(k, m, d, d2)
                Next d2
            Next d
        Next m
    Next k
    
End With
End Sub


Sub Read_Model(vRng As Range)
Dim i As Long, j As Long, k As Long, m As Long, d As Long, d2 As Long
With vRng

    pn_state = .Value
    pn_mixture = .Offset(1, 0).Value
    pn_dimension = .Offset(2, 0).Value
    pprob_type = .Offset(3, 0).Value
    ReDim pstart_probs(1 To pn_state)
    ReDim ptransitions(1 To pn_state, 1 To pn_state)
    ReDim pmix_wgt(1 To pn_state, 1 To pn_mixture)
    ReDim pmix_mean(1 To pn_state, 1 To pn_mixture, 1 To pn_dimension)
    ReDim pmix_var(1 To pn_state, 1 To pn_mixture, 1 To pn_dimension, 1 To pn_dimension)
    
    For i = 1 To pn_state
        pstart_probs(i) = .Offset(3 + i, 0).Value
        For j = 1 To pn_state
            ptransitions(i, j) = .Offset(3 + i, 1 + j).Value
        Next j
    Next i
    
    j = 0
    For k = 1 To pn_state
        For m = 1 To pn_mixture
            j = j + 1
            i = 4 + pn_state + (j - 1) * (pn_dimension + 2)
            pmix_wgt(k, m) = .Offset(i, 2).Value
            For d = 1 To pn_dimension
                pmix_mean(k, m, d) = .Offset(i + d, 0).Value
                For d2 = 1 To pn_dimension
                    pmix_var(k, m, d, d2) = .Offset(i + d, 1 + d2).Value
                Next d2
            Next d
        Next m
    Next k
    
End With
End Sub


Private Sub normalize_data(x As Variant, x_avg As Variant, x_sd As Variant)
Dim i As Long, d As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_raw As Long, n_dimension As Long
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim x_avg(1 To n_dimension)
    ReDim x_sd(1 To n_dimension)
    For d = 1 To n_dimension
        tmp_x = 0
        tmp_y = 0
        For i = 1 To n_raw
            tmp_x = tmp_x + x(i, d)
            tmp_y = tmp_y + x(i, d) ^ 2
        Next i
        x_avg(d) = tmp_x / n_raw
        x_sd(d) = Sqr((tmp_y / n_raw - (x_avg(d)) ^ 2) / (n_raw - 1))
        For i = 1 To n_raw
            x(i, d) = (x(i, d) - x_avg(d)) / x_sd(d)
        Next i
    Next d
End Sub


Function Print_Prob_Dist(Optional n_bin As Long = 50) As Double()
Dim i As Long, j As Long, k As Long, m As Long, d As Long, s As Long
Dim x_mean As Double, x_var As Double, dx As Double
Dim x_max As Double, x_min As Double
Dim tmp_x As Double, tmp_y As Double
Dim p() As Double, x() As Double, pdf() As Double, B() As Double, x_tmp() As Double, sort_index() As Long
Dim xS() As Variant, pdfs() As Variant, bs() As Variant
If pn_dimension > 1 Then
    ReDim p(0 To 0)
    Print_Prob_Dist = p
    Debug.Print "probability function is " & pn_dimension & "-dimensional. Cannot visualize."
    Exit Function
End If

d = 1
ReDim x(1 To n_bin * pn_state, 1 To pn_dimension)

x_max = -999999999
x_min = 999999999
For s = 1 To pn_state
    For m = 1 To pn_mixture
        tmp_x = pmix_mean(s, m, d) + 3 * Sqr(pmix_var(s, m, d, d))
        tmp_y = pmix_mean(s, m, d) - 3 * Sqr(pmix_var(s, m, d, d))
        If tmp_x > x_max Then x_max = tmp_x
        If tmp_y < x_min Then x_min = tmp_y
    Next m
    dx = (x_max - x_min) / (n_bin - 1)
    For i = 1 To n_bin
        x(i + (s - 1) * n_bin, d) = x_min + (i - 1) * dx
    Next i
Next s

ReDim xS(1 To 1)
xS(1) = x

Call Find_all_probs(xS, pdfs)
Call Find_all_b(pdfs, bs)
B = bs(1)
ReDim x_tmp(1 To n_bin * pn_state)
For i = 1 To n_bin * pn_state
    x_tmp(i) = x(i, d)
Next i
Call modMath.Sort_Quick_A(x_tmp, 1, n_bin * pn_state, sort_index)

ReDim p(1 To n_bin * pn_state, 1 To pn_state + 1)
For i = 1 To n_bin * pn_state
    j = sort_index(i)
    p(i, 1) = x(j, d)
    For s = 1 To pn_state
        p(i, 1 + s) = B(s, j)
    Next s
Next i

Print_Prob_Dist = p
Erase p, B, bs, pdfs, x_tmp, sort_index
End Function
