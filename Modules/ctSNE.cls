VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ctSNE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'===============================================================
't-Distributed Stochastic Neighbor Embedding (t-SNE)
'===============================================================
'Requires: mkdTree, cqtree, cqtree_point
'===============================================================
'Project high dimensional data to a 2/3-D for visualization
'Main reference:
'"Visualizing High-Dimensional Data Using t-SNE", Laurens van der Maaten (2008)
'"Accelerating t-SNE using Tree-Based Algorithms", Laurens van der Maaten (2014)
'Implementations can be found on the suthor's page: https://lvdmaaten.github.io/tsne/
'===============================================================

Private px() As Double  'Original data
Private py() As Double  'Projection
Private poutput_dimension As Long   'Number of dimensions in projections
Private pcost_function() As Double

Sub Reset()
    poutput_dimension = 0
    Erase px, py, pcost_function
End Sub

Public Property Get raw_data() As Double()
raw_data = px
End Property

Public Property Get cost_function(Optional show_index As Boolean = False) As Double()
Dim i As Long
If show_index = False Then
    cost_function = pcost_function
ElseIf show_index = True Then
    Dim y() As Double
    ReDim y(1 To UBound(pcost_function), 1 To 2)
    For i = 1 To UBound(pcost_function)
        y(i, 1) = i
        y(i, 2) = pcost_function(i)
    Next i
    cost_function = y
End If
End Property

Public Property Get Output() As Double()
Output = py
End Property


Private Function early_stopping(iterate As Long, conv_count As Long, conv_chk As Double) As Boolean
Dim tmp_x As Double
    early_stopping = False
    If pcost_function(iterate) < pcost_function(iterate - 1) Then
        conv_count = conv_count + 1
        If conv_count = 10 Then
            tmp_x = pcost_function(iterate) - conv_chk
            If tmp_x <= 0 And Abs(tmp_x) < 0.00001 Then early_stopping = True
            conv_chk = pcost_function(iterate)
        End If
    Else
        conv_count = 0
        conv_chk = pcost_function(iterate)
    End If
End Function

'Input: input_dist=FALSE: x(1 to n_raw, 1 to n_dimension), N x D feature vectors
'Input: input_dist=TRUE: x(1 to n_raw, 1 to n_raw), symmetric dissimarilty matrix, diagonals are zero
'Call tSNE1.tSNE(x2, 2, 30, 0.0001, 100, 0.5, 1000)
Sub tSNE(x() As Double, tgt_dimension As Long, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001, _
            Optional learn_rate As Double = 100, Optional momentum As Double = 0.5, Optional max_iterate As Long = 1000, _
            Optional input_dist As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_raw As Long
Dim prob_cond() As Double
Dim sigma As Double, sigma_min As Double, sigma_max As Double
Dim dist() As Double, y_dist() As Double
Dim q() As Double, dCdy() As Double, tmp_vec() As Double, tmp_vec2() As Double
Dim y_chg() As Double
Dim exaggerate As Double, exaggerate_step As Long
Dim p() As Double
Dim cost_const As Double
Dim gains() As Double
Dim conv_chk As Double, conv_count As Long

px = x
poutput_dimension = tgt_dimension
n_raw = UBound(x, 1)
ReDim py(1 To n_raw, 1 To tgt_dimension)
If input_dist = False Then
    dist = Calc_Euclidean_Dist(x)
ElseIf input_dist = True Then
    dist = x
End If

'Calculate pairwise similarities
p = Similarities(dist, perplexity, perp_err)

'Constant part of the cost function
cost_const = 0
For i = 1 To n_raw - 1
    For j = i + 1 To n_raw
        If p(j, i) > 0 Then cost_const = cost_const + p(j, i) * Log(p(j, i))
    Next j
Next i

'Initialize y() to random small values
Randomize
For i = 1 To n_raw
    For j = 1 To tgt_dimension
        py(i, j) = Rnd() / 10000
    Next j
Next i

exaggerate_step = 100
exaggerate = 4

'Pre-allocate memory
ReDim pcost_function(1 To max_iterate)
ReDim q(1 To n_raw, 1 To n_raw)
ReDim dCdy(1 To n_raw, 1 To tgt_dimension)
ReDim y_chg(1 To n_raw, 1 To tgt_dimension)
ReDim gains(1 To n_raw, 1 To tgt_dimension)
ReDim tmp_vec(1 To tgt_dimension)
ReDim tmp_vec2(1 To n_raw)
For i = 1 To n_raw
    For j = 1 To tgt_dimension
        gains(i, j) = 1
    Next j
Next i

conv_chk = Exp(70)
conv_count = 0
For iterate = 1 To max_iterate
    
    If iterate Mod 20 = 0 Then
        DoEvents
        Application.StatusBar = "t-SNE: " & iterate & "/" & max_iterate
    End If
    
    'Exaggerate probabilty during initial stage
    If iterate > exaggerate_step Then exaggerate = 1
    
    y_dist = Calc_Euclidean_Dist(py)
    
    tmp_x = 0
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            q(i, j) = 1# / (1 + y_dist(i, j))
            tmp_x = tmp_x + q(i, j)
        Next j
    Next i
    tmp_x = 2 * tmp_x
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            q(i, j) = q(i, j) / tmp_x
            q(j, i) = q(i, j)
        Next j
    Next i
    
    'Cost function of current iteration
    tmp_x = 0
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            tmp_x = tmp_x - p(i, j) * Log(q(i, j))
        Next j
    Next i
    pcost_function(iterate) = 2 * (cost_const + tmp_x)
    
    'Compute gradient
    For i = 1 To n_raw
        For j = 1 To n_raw
            tmp_vec2(j) = (p(i, j) - q(i, j)) / (1 + y_dist(i, j))
        Next j
        For k = 1 To tgt_dimension
            tmp_x = 0
            tmp_y = py(i, k)
            For j = 1 To n_raw
                'tmp_x = tmp_x + tmp_vec2(j) * (py(i, k) - py(j, k))
                tmp_x = tmp_x + tmp_vec2(j) * (tmp_y - py(j, k))
            Next j
            dCdy(i, k) = 4 * tmp_x
        Next k
    Next i

    'Update y() with adaptive learning rate
    For k = 1 To tgt_dimension
        For i = 1 To n_raw
            If Sgn(dCdy(i, k)) <> Sgn(y_chg(i, k)) Then
                gains(i, k) = gains(i, k) + 0.2
            ElseIf Sgn(dCdy(i, k)) = Sgn(y_chg(i, k)) Then
                gains(i, k) = gains(i, k) * 0.8
            End If
            If gains(i, k) < 0.01 Then gains(i, k) = 0.01
            y_chg(i, k) = -learn_rate * gains(i, k) * dCdy(i, k) + momentum * y_chg(i, k)
            py(i, k) = py(i, k) + y_chg(i, k)
        Next i
    Next k
    
    'check for convergence
    If iterate > 1 Then
        If early_stopping(iterate, conv_count, conv_chk) = True Then Exit For
    End If
    
Next iterate

If iterate < max_iterate Then ReDim Preserve pcost_function(1 To iterate)

Application.StatusBar = False
End Sub


Private Function Calc_Euclidean_Dist(x() As Double) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim tmp_x As Double, tmp_y As Double
Dim n_raw As Long, n_dimension As Long
Dim dist() As Double, x1() As Double
n_raw = UBound(x, 1)
n_dimension = UBound(x, 2)
ReDim dist(1 To n_raw, 1 To n_raw)
ReDim x1(1 To n_dimension)
For i = 1 To n_raw - 1
    For k = 1 To n_dimension
        x1(k) = x(i, k)
    Next k
    For j = i + 1 To n_raw
        tmp_x = 0
        For k = 1 To n_dimension
            tmp_x = tmp_x + (x1(k) - x(j, k)) ^ 2
        Next k
        dist(i, j) = tmp_x
        dist(j, i) = tmp_x
    Next j
Next i
Calc_Euclidean_Dist = dist
End Function




'Input: input_dist=FALSE: x(1 to n_raw, 1 to n_dimension), N x D feature vectors
'Input: input_dist=TRUE: x(1 to n_raw, 1 to n_raw), symmetric dissimarilty matrix, diagonals are zero
Sub tSNE_multi(x() As Double, tgt_dimension As Long, n_map As Long, _
            Optional perplexity As Double, Optional perp_err As Double, _
            Optional learn_rate As Double, Optional momentum As Double, Optional learn_rate_w As Double, Optional max_iterate As Long, _
            Optional input_dist As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double
Dim n_raw As Long, n_dimension As Long
Dim prob_cond() As Double, Prob() As Double
Dim sigma As Double, sigma_min As Double, sigma_max As Double
Dim dist() As Double, y_dist() As Double
Dim q() As Double, QQ() As Double
Dim y_chg() As Double
Dim exaggerate As Double, exaggerate_step As Long
Dim p() As Double, d() As Double, h As Double
Dim cost_const As Double
Dim gains() As Double
Dim proportion() As Double, weights() As Double
Dim dCdP() As Double, dCdW() As Double, dCdD() As Double, dCdy() As Double

If input_dist = False Then
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    dist = Calc_Euclidean_Dist(x)
ElseIf input_dist = True Then
    n_raw = UBound(x, 1)
    dist = x
End If
ReDim py(1 To n_raw, 1 To tgt_dimension, 1 To n_map)

ReDim prob_cond(1 To n_raw, 1 To n_raw)
For i = 1 To n_raw
    DoEvents
    Application.StatusBar = "Calculating perplexity... " & i & "/" & n_raw
    sigma_min = 0
    sigma_max = 50
    ReDim d(1 To n_raw - 1)
    ReDim p(1 To n_raw - 1)
    k = 0
    For j = 1 To n_raw
        If i <> j Then
            k = k + 1
            d(k) = dist(i, j)
        End If
    Next j
    
    'Binary search for sigma that gives the desired perplexity
    iterate = 0
    Do
        DoEvents
        iterate = iterate + 1
        sigma = (sigma_min + sigma_max) * 0.5
        
        tmp_x = 0
        tmp_y = 0
        For j = 1 To n_raw - 1
            p(j) = Exp(-d(j) / sigma)
            tmp_x = tmp_x + p(j)
            tmp_y = tmp_y + d(j) * p(j)
        Next j
        
        h = tmp_y / (tmp_x * sigma) + Log(tmp_x)
        
        For j = 1 To n_raw - 1
            p(j) = p(j) / tmp_x
        Next j
        
        tmp_x = h - Log(perplexity)
        If tmp_x > perp_err Then
            sigma_max = sigma
        ElseIf tmp_x < -perp_err Then
            sigma_min = sigma
        ElseIf Abs(tmp_x) <= perp_err Then
            Exit Do
        End If
    Loop While iterate <= 1000
    
    k = 0
    For j = 1 To n_raw
        If i <> j Then
            k = k + 1
            prob_cond(j, i) = p(k)
        End If
    Next j
    Debug.Print i & "," & Exp(h)
Next i

'Symmetrize joint probabilities
ReDim Prob(1 To n_raw, 1 To n_raw)
For i = 1 To n_raw - 1
    For j = i + 1 To n_raw
        Prob(i, j) = (prob_cond(j, i) + prob_cond(i, j)) / (2 * n_raw)
        Prob(j, i) = Prob(i, j)
    Next j
Next i
Erase prob_cond

'Constant part of the cost function
cost_const = 0
For i = 1 To n_raw - 1
    For j = i + 1 To n_raw
        cost_const = cost_const + Prob(j, i) * Log(Prob(j, i))
    Next j
Next i

'Initialize y() to random small values
Randomize
For i = 1 To n_raw
    For j = 1 To tgt_dimension
        For m = 1 To n_map
            py(i, j, m) = Rnd() / 1000
        Next m
    Next j
Next i

'Begin Gradient descent
exaggerate_step = 100
exaggerate = 4
ReDim pcost_function(1 To max_iterate)
ReDim dCdP(1 To n_raw, 1 To n_map)
ReDim dCdW(1 To n_raw, 1 To n_map)
ReDim dCdD(1 To n_raw, 1 To n_raw, 1 To n_map)
ReDim dCdy(1 To n_raw, 1 To tgt_dimension, 1 To n_map)
ReDim y_dist(1 To n_raw, 1 To n_raw, 1 To n_map)
ReDim y_chg(1 To n_raw, 1 To tgt_dimension, 1 To n_map)
ReDim weights(1 To n_raw, 1 To n_map)
ReDim proportion(1 To n_raw, 1 To n_map)
For i = 1 To n_raw
    For m = 1 To n_map
        weights(i, m) = 1# / n_map
    Next m
Next i

For iterate = 1 To max_iterate
    
    If iterate Mod 1 = 0 Then
        DoEvents
        Application.StatusBar = "t-SNE (multi-map): " & iterate & "/" & max_iterate
    End If
    
    'Exaggerate probabilty during initial stage
    If iterate > exaggerate_step Then exaggerate = 1
    
    'Compute importance from weights
    For i = 1 To n_raw
        tmp_x = 0
        For m = 1 To n_map
            proportion(i, m) = Exp(-weights(i, m))
            tmp_x = tmp_x + proportion(i, m)
        Next m
        For m = 1 To n_map
            proportion(i, m) = proportion(i, m) / tmp_x
        Next m
    Next i
    
    'Euclidean distance of y() in each map
    For m = 1 To n_map
        For i = 1 To n_raw - 1
            For j = i + 1 To n_raw
                tmp_x = 0
                For k = 1 To tgt_dimension
                    tmp_x = tmp_x + (py(i, k, m) - py(j, k, m)) ^ 2
                Next k
                y_dist(i, j, m) = tmp_x
                y_dist(j, i, m) = tmp_x
            Next j
        Next i
    Next m

    ReDim q(1 To n_raw, 1 To n_raw)
    tmp_z = 0
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            tmp_x = 0
            For m = 1 To n_map
                tmp_x = tmp_x + proportion(i, m) * proportion(j, m) / (1 + y_dist(i, j, m))
            Next m
            tmp_z = tmp_z + tmp_x
            q(i, j) = tmp_x
        Next j
    Next i
    tmp_z = 2 * tmp_z
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            q(i, j) = q(i, j) / tmp_z
            q(j, i) = q(i, j)
        Next j
    Next i
    
    'Cost function of current iteration
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            pcost_function(iterate) = pcost_function(iterate) - Prob(j, i) * Log(q(j, i))
        Next j
    Next i
    pcost_function(iterate) = 2 * (cost_const + pcost_function(iterate))
    
    'Compute gradients
    'w.r.t. importance
    For i = 1 To n_raw
        For m = 1 To n_map
            tmp_x = 0
            For j = 1 To n_raw
                If j <> i Then tmp_x = tmp_x + (Prob(i, j) - q(i, j)) * proportion(j, m) / _
                    ((1 + y_dist(i, j, m)) * q(i, j))
            Next j
            dCdP(i, m) = -2 * tmp_x / tmp_z
        Next m
    Next i
    
    'w.r.t. weights
    For i = 1 To n_raw
        For m = 1 To n_map
            tmp_x = 0
            For j = 1 To n_map
                tmp_x = tmp_x + proportion(i, j) * dCdP(i, j)
            Next j
            dCdW(i, m) = proportion(i, m) * (tmp_x - dCdP(i, m))
        Next m
    Next i
        
    'w.r.t. distance()
    For i = 1 To n_raw - 1
        For j = i + 1 To n_raw
            tmp_x = (Prob(i, j) - q(i, j)) / (q(i, j) * tmp_z)
            For m = 1 To n_map
                dCdD(i, j, m) = tmp_x * proportion(i, m) * proportion(j, m) / _
                    ((1 + y_dist(i, j, m)) ^ 2)
                dCdD(j, i, m) = dCdD(i, j, m)
            Next m
        Next j
    Next i
    
    'w.r.t. y()
    For i = 1 To n_raw
        For k = 1 To tgt_dimension
            For m = 1 To n_map
                tmp_x = 0
                For j = 1 To n_raw
                    tmp_x = tmp_x + dCdD(i, j, m) * (py(i, k, m) - py(j, k, m))
                Next j
                dCdy(i, k, m) = 4 * tmp_x
            Next m
        Next k
    Next i

    'Update y() and weights()
    For i = 1 To n_raw
        For k = 1 To tgt_dimension
            For m = 1 To n_map
                y_chg(i, k, m) = -learn_rate * dCdy(i, k, m) + momentum * y_chg(i, k, m)
                py(i, k, m) = py(i, k, m) + y_chg(i, k, m)
            Next m
        Next k
        
        For m = 1 To n_map
            weights(i, m) = weights(i, m) - learn_rate_w * dCdW(i, m)
        Next m
    Next i
    
Next iterate

Application.StatusBar = False
End Sub




Sub tSNE_BarnesHut(x() As Double, tgt_dimension As Long, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001, _
            Optional learn_rate As Double = 100, Optional momentum As Double = 0.5, Optional max_iterate As Long = 1000, _
            Optional input_dist As Boolean = False)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, k_max As Long, n_raw As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, d As Double, logz As Double
Dim sigma As Double, sigma_min As Double, sigma_max As Double
Dim p() As Double, q() As Double, dCdy() As Double, tmp_vec() As Double, y_chg() As Double
Dim exaggerate As Double, exaggerate_step As Long
Dim cost_const As Double
Dim gains() As Double
Dim conv_chk As Double, conv_count As Long
Dim k_idx() As Long, k_dist() As Double
Dim quadtree As cqtree

If tgt_dimension <> 2 Then
    Debug.Print "ctSNE: tSNE_BarnesHut only supports 2D output at the moment."
    Exit Sub
End If

px = x
poutput_dimension = tgt_dimension
n_raw = UBound(x, 1)
k_max = Int(3 * perplexity)
ReDim py(1 To n_raw, 1 To tgt_dimension)
p = Similarities_Tree(x, k_idx, perplexity, perp_err, input_dist)

'Constant part of the cost function
cost_const = 0
For i = 1 To n_raw
    For k = 1 To k_max
        j = k_idx(i, k)
        If p(j, i) > 0 Then cost_const = cost_const + p(j, i) * Log(p(j, i))
    Next k
Next i

'Initialize y() to random small values
Randomize
For i = 1 To n_raw
    For j = 1 To tgt_dimension
        py(i, j) = Rnd() / 10000
    Next j
Next i

exaggerate_step = 100
exaggerate = 4

'Pre-allocate memory
ReDim pcost_function(1 To max_iterate)
ReDim dCdy(1 To n_raw, 1 To tgt_dimension)
ReDim y_chg(1 To n_raw, 1 To tgt_dimension)
ReDim gains(1 To n_raw, 1 To tgt_dimension)
ReDim tmp_vec(1 To k_max)
For i = 1 To n_raw
    For j = 1 To tgt_dimension
        gains(i, j) = 1
    Next j
Next i

'Begin Gradient descent
conv_count = 0
conv_chk = Exp(70)
For iterate = 1 To max_iterate

    If iterate Mod 20 = 0 Then
        DoEvents
        Application.StatusBar = "t-SNE (Barnes-Hut): " & iterate & "/" & max_iterate
    End If

    Set quadtree = New cqtree
    q = quadtree.tSNE_Force(py, tmp_z)

    'Exaggerate probabilty during initial stage
    If iterate > exaggerate_step Then exaggerate = 1

    'Compute gradient
    tmp_y = 0
    logz = Log(tmp_z)
    For i = 1 To n_raw
        For j = 1 To k_max
            d = y_ij(i, k_idx(i, j))
            tmp_vec(j) = p(i, k_idx(i, j)) / (1 + d)
            tmp_y = tmp_y + p(i, k_idx(i, j)) * (Log(1 + d) + logz)
        Next j
        For k = 1 To tgt_dimension
            tmp_x = 0
            For j = 1 To k_max
                tmp_x = tmp_x + tmp_vec(j) * (py(i, k) - py(k_idx(i, j), k))
            Next j
            dCdy(i, k) = 4 * (tmp_x - q(i, k) / tmp_z)
        Next k
    Next i
    pcost_function(iterate) = (cost_const + tmp_y) 'Cost function of current iteration

    'Update y() with adaptive learning rate
    For k = 1 To tgt_dimension
        For i = 1 To n_raw
            If Sgn(dCdy(i, k)) <> Sgn(y_chg(i, k)) Then
                gains(i, k) = gains(i, k) + 0.2
            ElseIf Sgn(dCdy(i, k)) = Sgn(y_chg(i, k)) Then
                gains(i, k) = gains(i, k) * 0.8
            End If
            If gains(i, k) < 0.01 Then gains(i, k) = 0.01
            y_chg(i, k) = -learn_rate * gains(i, k) * dCdy(i, k) + momentum * y_chg(i, k)
            py(i, k) = py(i, k) + y_chg(i, k)
        Next i
    Next k
    
    'check for convergence
    If iterate > 1 Then
        If early_stopping(iterate, conv_count, conv_chk) = True Then Exit For
    End If
    
Next iterate
If iterate < max_iterate Then ReDim Preserve pcost_function(1 To iterate)
Erase dCdy, y_chg, gains, tmp_vec, p, q, k_idx
Set quadtree = Nothing
Application.StatusBar = False
End Sub


Private Function y_ij(i As Long, j As Long) As Double
Dim k As Long
For k = 1 To poutput_dimension
    y_ij = y_ij + (py(i, k) - py(j, k)) ^ 2
Next k
End Function

'Input: N x N distance (sum of square) matrix
'Output:  N x N  pairwise similarities matrix
Private Function Similarities(dist() As Double, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, n_raw As Long
Dim tmp_x As Double, tmp_y As Double
Dim prob_cond() As Double
Dim sigma As Double, sigma_min As Double, sigma_max As Double
Dim p() As Double, d() As Double, h As Double

n_raw = UBound(dist, 1)
ReDim prob_cond(1 To n_raw, 1 To n_raw)
For i = 1 To n_raw
    
    If i Mod 100 = 0 Then
        DoEvents
        Application.StatusBar = "ctSNE: Similarities: " & i & "/" & n_raw
    End If
    
    sigma_min = 0
    sigma_max = 50
    ReDim d(1 To n_raw - 1)
    ReDim p(1 To n_raw - 1)
    k = 0
    For j = 1 To n_raw
        If i <> j Then
            k = k + 1
            d(k) = dist(i, j)
        End If
    Next j
    
    'Binary search for sigma that gives the desired perplexity
    iterate = 0
    Do
        iterate = iterate + 1
        sigma = (sigma_min + sigma_max) * 0.5
        
        tmp_x = 0
        tmp_y = 0
        For j = 1 To n_raw - 1
            p(j) = Exp(-d(j) / sigma)
            tmp_x = tmp_x + p(j)
            tmp_y = tmp_y + d(j) * p(j)
        Next j
        
        h = tmp_y / (tmp_x * sigma) + Log(tmp_x)
        
        For j = 1 To n_raw - 1
            p(j) = p(j) / tmp_x
        Next j
        
        tmp_x = h - Log(perplexity)
        If tmp_x > perp_err Then
            sigma_max = sigma
        ElseIf tmp_x < -perp_err Then
            sigma_min = sigma
        ElseIf Abs(tmp_x) <= perp_err Then
            Exit Do
        End If
    Loop While iterate <= 1000
    
    k = 0
    For j = 1 To n_raw
        If i <> j Then
            k = k + 1
            prob_cond(j, i) = p(k)
        End If
    Next j

Next i

'Symmetrize joint probabilities
ReDim p(1 To n_raw, 1 To n_raw)
For i = 1 To n_raw - 1
    For j = i + 1 To n_raw
        p(i, j) = (prob_cond(j, i) + prob_cond(i, j)) / (2 * n_raw)
        p(j, i) = p(i, j)
    Next j
Next i
Erase prob_cond, d
Similarities = p
Erase p
End Function


'Input: x(1 to N, 1 to D), N rows of D dimensional vector
'Output: N X N pairwise similarity matrix
'Output: k_idx(1 to N, 1 to [3*perplexity]), adjaceny list of k-nearest neighors graph
Private Function Similarities_Tree(x() As Double, k_idx() As Long, _
            Optional perplexity As Double = 30, Optional perp_err As Double = 0.0001, _
            Optional input_dist As Boolean = False) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, k_max As Long, n_raw As Long
Dim tmp_x As Double, tmp_y As Double, tmp_z As Double, d As Double
Dim prob_cond() As Double
Dim sigma As Double, sigma_min As Double, sigma_max As Double
Dim p() As Double, h As Double
Dim k_dist() As Double

n_raw = UBound(x, 1)
k_max = Int(3 * perplexity)

'Build k-nearest neighbor graph
If input_dist = False Then
    Call mkdTree.kNN_All(k_idx, k_dist, x, k_max, 0)
Else
    Call kNN_Graph(x, k_idx, k_dist, k_max)
End If

ReDim p(1 To k_max)
ReDim prob_cond(1 To n_raw, 1 To n_raw)
For i = 1 To n_raw
    
    If i Mod 100 = 0 Then
        DoEvents
        Application.StatusBar = "ctSNE: Similarities_Tree: " & i & "/" & n_raw
    End If
    
    sigma_min = 0
    sigma_max = 50
    
    'Binary search for sigma that gives the desired perplexity
    iterate = 0
    Do
        iterate = iterate + 1
        sigma = (sigma_min + sigma_max) * 0.5

        tmp_x = 0
        tmp_y = 0
        For j = 1 To k_max
            d = k_dist(i, j) ^ 2
            p(j) = Exp(-d / sigma)
            tmp_x = tmp_x + p(j)
            tmp_y = tmp_y + d * p(j)
        Next j

        h = tmp_y / (tmp_x * sigma) + Log(tmp_x)

        For j = 1 To k_max
            p(j) = p(j) / tmp_x
        Next j

        tmp_x = h - Log(perplexity)
        If tmp_x > perp_err Then
            sigma_max = sigma
        ElseIf tmp_x < -perp_err Then
            sigma_min = sigma
        ElseIf Abs(tmp_x) <= perp_err Then
            Exit Do
        End If
    Loop While iterate <= 1000

    For j = 1 To k_max
        prob_cond(k_idx(i, j), i) = p(j)
    Next j

Next i

'Symmetrize joint probabilities
ReDim p(1 To n_raw, 1 To n_raw)
For i = 1 To n_raw - 1
    For j = i + 1 To n_raw
        p(i, j) = (prob_cond(j, i) + prob_cond(i, j)) / (2 * n_raw)
        p(j, i) = p(i, j)
    Next j
Next i
Erase prob_cond, k_dist
Similarities_Tree = p
Erase p
End Function


'Build k-Nearest Neighbors Graph from a N X N distance matrix
Private Sub kNN_Graph(dist() As Double, k_idx() As Long, k_dist() As Double, k As Long)
Dim i As Long, j As Long, n As Long, n_raw As Long
Dim d() As Double, idx() As Long
n_raw = UBound(dist, 1)
ReDim k_idx(1 To n_raw, 1 To k)
ReDim k_dist(1 To n_raw, 1 To k)
ReDim d(1 To n_raw - 1)
ReDim idx(1 To n_raw - 1)
For i = 1 To n_raw
    If i Mod 50 = 0 Then
        DoEvents
        Application.StatusBar = "ctSNE: kNN_Graph: " & i & "/" & n_raw
    End If
    n = 0
    For j = 1 To n_raw
        If j <> i Then
            n = n + 1
            d(n) = dist(i, j)
            idx(n) = j
        End If
    Next j
    Call modMath.Sort_Quick_A(d, 1, n, idx, 0)
    For j = 1 To k
        k_dist(i, j) = Sqr(d(j))
        k_idx(i, j) = idx(j)
    Next j
Next i
Erase d, idx
Application.StatusBar = False
End Sub
