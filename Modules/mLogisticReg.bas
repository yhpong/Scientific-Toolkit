Attribute VB_Name = "mLogisticReg"
Option Explicit

'========================================================
'Perform binary logistic regression with gradient descent
'========================================================
'Output: beta(1:D+1), regression coefficients of D-dimension, the D+1 element is the bias term
'Input:  y(1:N), binary target N observations
'        x(1:N,1:D), D-dimensional feature vector of N observations
'        learn_rate, learning rate for gradient descent
Function Train(y As Variant, x As Variant, _
        Optional learn_rate As Double = 0.001, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 2000, _
        Optional conv_max As Long = 5, Optional conv_tol As Double = 0.000001, _
        Optional loss_function As Variant, _
        Optional L2 As Double = 0, _
        Optional adaptive_learn As Boolean = True, _
        Optional show_progress As Boolean = True, _
        Optional init_beta As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long, ii As Long
Dim batch_count As Long, epoch As Long, conv_count As Long
Dim tmp_x As Double, tmp_y As Double, delta As Double, y_output As Double
Dim max_gain As Double
Dim loss() As Double, grad() As Double, grad_prev() As Double, gain() As Double
Dim beta() As Double, beta_chg() As Double
Dim iArr() As Long
    n = UBound(x, 1)            'number of observations
    n_dimension = UBound(x, 2)  'number of dimensions
    max_gain = 1 / learn_rate   's.t. max learn rate is 1
    
    'Initialize beta() to zeroes
    ReDim beta(1 To n_dimension + 1)
    ReDim beta_chg(1 To n_dimension + 1)
    
    If IsMissing(init_beta) = False Then beta = init_beta
    
    'Perform gradient descent
    conv_count = 0
    ReDim loss(1 To epoch_max)
    For epoch = 1 To epoch_max
        
        If show_progress = True Then
            If epoch Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = "mLogisticReg: Train: " & epoch & "/" & epoch_max
            End If
        End If
        
        'Shuffle data set
        iArr = modMath.index_array(1, n)
        Call modMath.Shuffle(iArr)
        
        'Reset gradients
        batch_count = 0
        ReDim grad(1 To n_dimension + 1)
        ReDim grad_prev(1 To n_dimension + 1)
        ReDim gain(1 To n_dimension + 1)
        For j = 1 To n_dimension + 1
            gain(j) = 1
        Next j
        
        'Scan through dataset
        For ii = 1 To n
            i = iArr(ii)
            
            'beta dot x
            y_output = beta(n_dimension + 1)
            For j = 1 To n_dimension
                y_output = y_output + beta(j) * x(i, j)
            Next j
            y_output = 1# / (1 + Exp(-y_output)) 'Sigmoid function
            
            'accumulate gradient
            delta = y_output - y(i)
            For j = 1 To n_dimension
                grad(j) = grad(j) + x(i, j) * delta
            Next j
            grad(n_dimension + 1) = grad(n_dimension + 1) + delta
            
            'update beta() when mini batch count is reached
            batch_count = batch_count + 1
            If batch_count = mini_batch Or ii = n Then
                For j = 1 To n_dimension + 1
                    grad(j) = grad(j) / batch_count
                Next j
                If L2 > 0 Then 'L2-regularization
                    For j = 1 To n_dimension
                        grad(j) = grad(j) + L2 * beta(j)
                    Next j
                End If
                If adaptive_learn = True Then
                    Call calc_gain(grad, grad_prev, gain, max_gain)
                End If
                For j = 1 To n_dimension + 1
                    beta_chg(j) = momentum * beta_chg(j) - learn_rate * grad(j) * gain(j)
                    beta(j) = beta(j) + beta_chg(j)
                Next j
                'reset mini batch count and gradient
                batch_count = 0
                grad_prev = grad
                ReDim grad(1 To n_dimension + 1)
            End If
            
        Next ii

        loss(epoch) = Cross_Entropy(y, Predict(beta, x, , False), False)
        
        'early terminate on convergence
        If epoch > 1 Then
            If loss(epoch) < 0.05 Then
                ReDim Preserve loss(1 To epoch)
                Exit For
            End If
            If loss(epoch) <= loss(epoch - 1) Then
                conv_count = conv_count + 1
                If conv_count > conv_max Then
                    If (loss(epoch - 1) - loss(epoch)) < conv_tol Then
                        ReDim Preserve loss(1 To epoch)
                        Exit For
                    End If
                End If
            Else
                conv_count = 0
            End If
        End If
        
    Next epoch
    
    If epoch >= epoch_max Then
        Debug.Print "mLogisticReg: Train: Failed to converge in " & epoch_max & " epochs. L2=" & L2
    End If
    
    Train = beta
    
    If IsMissing(loss_function) = False Then loss_function = loss
    Erase loss, grad, grad_prev, gain, beta_chg, iArr, beta
    Application.StatusBar = False
End Function


'========================================================
'Perform softmax regression with gradient descent
'========================================================
'Output: beta(1:D+1,1:K), regression coefficients of D-dimension, the D+1 element is the bias term
'Input:  y(1:N,1:K), binary array of N observations and K classes
'        x(1:N,1:D), D-dimensional feature vector of N observations
'        learn_rate, learning rate for gradient descent
Function Train_Multiclass(y As Variant, x As Variant, _
        Optional learn_rate As Double = 0.001, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 2000, _
        Optional conv_max As Long = 5, Optional conv_tol As Double = 0.000001, _
        Optional loss_function As Variant, _
        Optional L2 As Double = 0, _
        Optional adaptive_learn As Boolean = True, _
        Optional show_progress As Boolean = True, _
        Optional init_beta As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long, n_class As Long, ii As Long
Dim batch_count As Long, epoch As Long, conv_count As Long
Dim tmp_x As Double, tmp_y As Double, delta As Double
Dim max_gain As Double
Dim loss() As Double, grad() As Double, grad_prev() As Double, gain() As Double
Dim beta() As Double, beta_chg() As Double, xi() As Double, y_output() As Double
Dim iArr() As Long
    n = UBound(x, 1)            'number of observations
    n_dimension = UBound(x, 2)  'number of dimensions
    n_class = UBound(y, 2)      'number of output class
    max_gain = 1 / learn_rate   's.t. max learn rate is 1
    
    'Initialize beta() to zeroes
    ReDim beta(1 To n_dimension + 1, 1 To n_class)
    ReDim beta_chg(1 To n_dimension + 1, 1 To n_class)
    
    If IsMissing(init_beta) = False Then beta = init_beta
    
    'Perform gradient descent
    conv_count = 0
    ReDim loss(1 To epoch_max)
    For epoch = 1 To epoch_max
        
        If show_progress = True Then
            If epoch Mod 100 = 0 Then
                DoEvents
                Application.StatusBar = "mLogisticReg: Train_Multiclass: " & epoch & "/" & epoch_max
            End If
        End If
        
        'Shuffle data set
        iArr = modMath.index_array(1, n)
        Call modMath.Shuffle(iArr)
        
        'Reset gradients
        batch_count = 0
        ReDim grad(1 To n_dimension + 1, 1 To n_class)
        ReDim grad_prev(1 To n_dimension + 1, 1 To n_class)
        ReDim gain(1 To n_dimension + 1, 1 To n_class)
        For i = 1 To n_class
            For j = 1 To n_dimension + 1
                gain(j, i) = 1
            Next j
        Next i
        
        'Scan through dataset
        ReDim xi(1 To n_dimension)
        For ii = 1 To n
            i = iArr(ii)
            For j = 1 To n_dimension
                xi(j) = x(i, j)
            Next j
            'Softmax output
            tmp_y = 0
            ReDim y_output(1 To n_class)
            For k = 1 To n_class
                tmp_x = beta(n_dimension + 1, k)
                For j = 1 To n_dimension
                    tmp_x = tmp_x + xi(j) * beta(j, k)
                Next j
                y_output(k) = Exp(tmp_x)
                tmp_y = tmp_y + y_output(k)
            Next k
            For k = 1 To n_class
                y_output(k) = y_output(k) / tmp_y
            Next k
            'Accumulate gradient
            For k = 1 To n_class
                delta = y_output(k) - y(i, k)
                For j = 1 To n_dimension
                    grad(j, k) = grad(j, k) + xi(j) * delta
                Next j
                grad(n_dimension + 1, k) = grad(n_dimension + 1, k) + delta
            Next k
            
            'update beta() when mini batch count is reached
            batch_count = batch_count + 1
            If batch_count = mini_batch Or ii = n Then
                For k = 1 To n_class
                    For j = 1 To n_dimension + 1
                        grad(j, k) = grad(j, k) / batch_count
                    Next j
                Next k
                If L2 > 0 Then 'L2-regularization
                    For k = 1 To n_class
                        For j = 1 To n_dimension
                            grad(j, k) = grad(j, k) + L2 * beta(j, k)
                        Next j
                    Next k
                End If
                If adaptive_learn = True Then
                    Call calc_gain(grad, grad_prev, gain, max_gain, True)
                End If
                For k = 1 To n_class
                    For j = 1 To n_dimension + 1
                        beta_chg(j, k) = momentum * beta_chg(j, k) - learn_rate * grad(j, k) * gain(j, k)
                        beta(j, k) = beta(j, k) + beta_chg(j, k)
                    Next j
                Next k
                
                'reset mini batch count and gradient
                batch_count = 0
                grad_prev = grad
                ReDim grad(1 To n_dimension + 1, 1 To n_class)
            End If
            
        Next ii

        loss(epoch) = Cross_Entropy(y, Predict(beta, x, , True), True)
        
        'early terminate on convergence
        If epoch > 1 Then
            If loss(epoch) < 0.05 Then
                ReDim Preserve loss(1 To epoch)
                Exit For
            End If
            If loss(epoch) <= loss(epoch - 1) Then
                conv_count = conv_count + 1
                If conv_count > conv_max Then
                    If (loss(epoch - 1) - loss(epoch)) < conv_tol Then
                        ReDim Preserve loss(1 To epoch)
                        Exit For
                    End If
                End If
            Else
                conv_count = 0
            End If
        End If
        
    Next epoch
    
    If epoch >= epoch_max Then
        Debug.Print "mLogisticReg: Train_Multiclass: Failed to converge in " & epoch_max & " epochs. L2=" & L2
    End If
    
    Train_Multiclass = beta
    
    If IsMissing(loss_function) = False Then loss_function = loss
    Erase loss, grad, grad_prev, gain, beta_chg, iArr, beta
    Application.StatusBar = False
End Function


'========================================================
'Perform K-fold cross-validation to find optimal L2 regularization
'========================================================
Function Train_CV(y As Variant, x As Variant, Optional k_fold As Long = 10, _
        Optional learn_rate As Double = 0.001, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 2000, _
        Optional conv_max As Long = 5, Optional conv_tol As Double = 0.000001, _
        Optional loss_function As Variant, _
        Optional L2_max As Double = 1, Optional L2_best As Variant, _
        Optional adaptive_learn As Boolean = True, _
        Optional isMulticlass As Boolean = False) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long, n_class As Long
Dim i_cv As Long, ii As Long, jj As Long, ii_max As Long, jj_max As Long
Dim tmp_x As Double, L2 As Double
Dim loss() As Double, loss_min As Double
Dim y_output() As Double
Dim iArr() As Long, i_validate() As Long, i_train() As Long
Dim x_train() As Double, x_validate() As Double
Dim y_train() As Double, y_validate() As Double
Dim beta_prev() As Double, beta() As Double
Dim L2_list() As Double

    n = UBound(x, 1)            'number of observations
    n_dimension = UBound(x, 2)  'number of dimensions
    If isMulticlass = True Then n_class = UBound(y, 2) 'number of class
    
    'Shuffle data set
    iArr = modMath.index_array(1, n)
    Call modMath.Shuffle(iArr)
    
    jj_max = 0: L2 = 0
    If L2_max > 0 Then jj_max = 6
    
    'ReDim accur(0 To jj_max)
    ReDim loss(0 To jj_max)
    If isMulticlass = False Then
        ReDim beta_prev(1 To n_dimension + 1)
    Else
        ReDim beta_prev(1 To n_dimension + 1, 1 To n_class)
    End If
    
    'List of L2 to try
    ReDim L2_list(0 To jj_max)
    For jj = 1 To jj_max
        L2_list(jj) = L2_max * (5 ^ (jj - jj_max))
    Next jj
    
    'Outer loop for different L2 values
    For jj = jj_max To 0 Step -1
        L2 = L2_list(jj)
        'K-fold cross-validation
        For i_cv = 1 To k_fold
            DoEvents
            Application.StatusBar = "Train_CV: " & (jj_max - jj) & "/" & jj_max & ";" & i_cv & "/" & k_fold
            Call modMath.CrossValidate_set(i_cv, k_fold, iArr, i_validate, i_train)
            Call modMath.Filter_Array(y, y_validate, i_validate)
            Call modMath.Filter_Array(x, x_validate, i_validate)
            Call modMath.Filter_Array(y, y_train, i_train)
            Call modMath.Filter_Array(x, x_train, i_train)
            If isMulticlass = False Then
                beta = Train(y_train, x_train, learn_rate, momentum, _
                    mini_batch, epoch_max, conv_max, conv_tol, , L2, adaptive_learn, False, beta_prev)
            Else
                beta = Train_Multiclass(y_train, x_train, learn_rate, momentum, _
                    mini_batch, epoch_max, conv_max, conv_tol, , L2, adaptive_learn, False, beta_prev)
            End If
            y_output = Predict(beta, x_validate, , isMulticlass)
            'loss(jj) = loss(jj) - Accuracy(y_validate, y_output, isMulticlass) * UBound(y_validate) / n
            loss(jj) = loss(jj) + Cross_Entropy(y_validate, y_output, isMulticlass) / k_fold
            beta_prev = beta
        Next i_cv
    Next jj
    Erase x_train, y_train, x_validate, y_validate
    Erase iArr, i_validate, i_train
    
    'Find L2 that gives lowest loss
    loss_min = Exp(70)
    For jj = 0 To jj_max
        Debug.Print "L2 & loss, " & L2_list(jj) & ", " & loss(jj)
        If loss(jj) < loss_min Then
            loss_min = loss(jj)
            L2 = L2_list(jj)
        End If
    Next jj
    Debug.Print "Train_CV: Best L2 = " & L2 & ", loss=" & loss_min
    If IsMissing(L2_best) = False Then L2_best = L2
    
    'Use selected L2 to train on whole dataset again
    If isMulticlass = False Then
        beta = Train(y, x, learn_rate, momentum, mini_batch, epoch_max, _
                    conv_max, conv_tol, loss, L2, adaptive_learn, True)
    Else
        beta = Train_Multiclass(y, x, learn_rate, momentum, mini_batch, epoch_max, _
                conv_max, conv_tol, loss, L2, adaptive_learn, True)
    End If
    If IsMissing(loss_function) = False Then loss_function = loss
    Train_CV = beta
    Erase beta, beta_prev
    Erase loss, y_output
    Application.StatusBar = False
End Function


'===========================
'Evaluate Model performance
'===========================

'Calculate accuracy
Function Accuracy(y_tgt As Variant, y As Variant, Optional isMulticlass As Boolean = False) As Double
Dim i As Long, j As Long, k As Long, n As Long, n_class As Long
Dim tmp_x As Double, tmp_max As Double
    n = UBound(y, 1)
    tmp_x = 0
    If isMulticlass = False Then
        For i = 1 To n
            If y_tgt(i) >= 0.5 Then
                If y(i) >= 0.5 Then tmp_x = tmp_x + 1
            ElseIf y_tgt(i) < 0.5 Then
                If y(i) < 0.5 Then tmp_x = tmp_x + 1
            End If
        Next i
    Else
        n_class = UBound(y, 2)
        For i = 1 To n
            tmp_max = y(i, 1)
            j = 1
            For k = 2 To n_class
                If y(i, k) > tmp_max Then
                    tmp_max = y(i, k)
                    j = k
                End If
            Next k
            If y_tgt(i, j) = 1 Then tmp_x = tmp_x + 1
        Next i
    End If
    Accuracy = tmp_x / n
End Function

'Calculate cross entropy
Function Cross_Entropy(y_tgt As Variant, y As Variant, Optional isMulticlass As Boolean = False) As Double
Dim i As Long, k As Long, n As Long
Dim tmp_x As Double, tmp_y As Double
    n = UBound(y, 1)
    tmp_x = 0
    If isMulticlass = False Then
        For i = 1 To n
            tmp_y = log_clip(y(i))
            tmp_x = tmp_x - y_tgt(i) * Log(tmp_y) - (1 - y_tgt(i)) * Log(1 - tmp_y)
        Next i
    Else
        For k = 1 To UBound(y, 2)
            For i = 1 To n
                tmp_y = log_clip(y(i, k))
                tmp_x = tmp_x - y_tgt(i, k) * Log(tmp_y)
            Next i
        Next k
    End If
    Cross_Entropy = tmp_x / n
End Function

Private Function log_clip(y) As Double
    log_clip = y
    If log_clip > 0.999999999 Then
        log_clip = 0.999999999
    ElseIf log_clip < 0.000000001 Then
        log_clip = 0.000000001
    End If
End Function

'===========================
'Make Predictions
'===========================
'Output: y(1:N), binary classification output
'        or y(1:N,1:K), if isMultclass is set to true
'Input:  beta(1:D+1), regression coefficients of D-dimension, the D+1 element is the bias term
'        or beta(1:D+1.1:K), if isMultclass is set to true
'        x(1:N,1:D), D-dimensional feature vector of N observations
'        force_binary, if set to true then y() will be set to exactly 0 or 1.
Function Predict(beta() As Double, x As Variant, _
    Optional force_binary As Boolean = False, _
    Optional isMulticlass As Boolean = False) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long, n_class As Long
Dim tmp_x As Double, tmp_max As Double
Dim y() As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    y = modMath.M_Dot(x, beta)
    If isMulticlass = False Then
        For i = 1 To n
            y(i) = 1# / (1 + Exp(-(y(i) + beta(n_dimension + 1))))
        Next i
        If force_binary = True Then
            For i = 1 To n
                If y(i) >= 0.5 Then
                    y(i) = 1
                Else
                    y(i) = 0
                End If
            Next i
        End If
    Else
        n_class = UBound(beta, 2)
        For i = 1 To n
            tmp_x = 0
            For k = 1 To n_class
                y(i, k) = Exp(y(i, k) + beta(n_dimension + 1, k))
                tmp_x = tmp_x + y(i, k)
            Next k
            For k = 1 To n_class
                y(i, k) = y(i, k) / tmp_x
            Next k
        Next i
        If force_binary = True Then
            For i = 1 To n
                tmp_max = y(i, 1)
                j = 1
                For k = 2 To n_class
                    If y(i, k) >= tmp_max Then
                        tmp_max = y(i, k)
                        j = k
                    End If
                Next k
                For k = 1 To n_class
                    y(i, k) = 0
                Next k
                y(i, j) = 1
            Next i
        End If
    End If
    Predict = y
    Erase y
End Function


'Adjust beta to reflect scales of x() before transformation
Sub Rescale_beta(beta() As Double, x_mean() As Double, x_sd() As Double, Optional isMulticlass As Boolean = False)
Dim i As Long, k As Long, n As Long
    n = UBound(beta, 1) - 1
    If isMulticlass = False Then
        For i = 1 To n
            beta(i) = beta(i) / x_sd(i)
            beta(n + 1) = beta(n + 1) - beta(i) * x_mean(i)
        Next i
    Else
        For k = 1 To UBound(beta, 2)
            For i = 1 To n
                beta(i, k) = beta(i, k) / x_sd(i)
                beta(n + 1, k) = beta(n + 1, k) - beta(i, k) * x_mean(i)
            Next i
        Next k
    End If
End Sub


'If current update is in the same direction as previous update, increase
'learn rate since there is more confidence in its direction.
Private Sub calc_gain(grad() As Double, grad_prev() As Double, gain() As Double, max_gain As Double, _
            Optional isMulticlass As Boolean = False)
Dim i As Long, k As Long, n As Long, n_class As Long
    n = UBound(grad)
    If isMulticlass = False Then
        For i = 1 To n
            If Sgn(grad(i)) = Sgn(grad_prev(i)) Then
                gain(i) = gain(i) * 1.1
            Else
                gain(i) = gain(i) * 0.9
            End If
            If gain(i) > max_gain Then gain(i) = max_gain
            If gain(i) < 0.01 Then gain(i) = 0.01
        Next i
    Else
        n_class = UBound(grad, 2)
        For k = 1 To n_class
            For i = 1 To n
                If Sgn(grad(i, k)) = Sgn(grad_prev(i, k)) Then
                    gain(i, k) = gain(i, k) * 1.1
                Else
                    gain(i, k) = gain(i, k) * 0.9
                End If
                If gain(i, k) > max_gain Then gain(i, k) = max_gain
                If gain(i, k) < 0.01 Then gain(i, k) = 0.01
            Next i
        Next k
    End If
End Sub
