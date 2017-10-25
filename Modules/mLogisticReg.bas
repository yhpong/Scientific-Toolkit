Attribute VB_Name = "mLogisticReg"
Option Explicit

'========================================================
'Perform binary logistic regression with gradient descent
'========================================================
'Output: beta(1:D+1), regression coefficients of D-dimension, the D+1 element is the bias term
'Input: y(1:N), binary target N observations
'       x(1:N,1:D), D-dimensional feature vector of N observations
'       learn_rate, learning rate for gradient descent
Function Binary_Train(y As Variant, x As Variant, _
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
                Application.StatusBar = "mLogisticReg: Binary_Train: " & epoch & "/" & epoch_max
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
'                If L1 > 0 Then 'L1-regularization
'                    For j = 1 To n_dimension
'                        grad(j) = grad(j) + L1 * Sgn(beta(j))
'                    Next j
'                End If
                If L2 > 0 Then 'L2-regularization
                    For j = 1 To n_dimension
                        grad(j) = grad(j) + L2 * beta(j)
                    Next j
                End If
                
                If adaptive_learn = True Then
                    Call calc_gain(grad, grad_prev, gain, max_gain)
                End If
                
                For j = 1 To n_dimension + 1
                    beta_chg(j) = momentum * beta_chg(j) - grad(j) * learn_rate * gain(j)
                    beta(j) = beta(j) + beta_chg(j)
                Next j
                
                'reset mini batch count and gradient
                batch_count = 0
                grad_prev = grad
                ReDim grad(1 To n_dimension + 1)
            End If
            
        Next ii

        'loss(epoch) = loss(epoch) / n
        loss(epoch) = Cross_Entropy(y, Binary_Predict(beta, x))
        
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
        Debug.Print "mLogisticReg: Binary_Train: Failed to converge in " & epoch_max & " epochs. L2=" & L2
    End If
    
    Binary_Train = beta
    
    If IsMissing(loss_function) = False Then loss_function = loss
    Erase loss, grad, grad_prev, gain, beta_chg, iArr, beta
    Application.StatusBar = False
End Function


'========================================================
'Perform K-fold crossvalidation to find optimal L2 regularization
'========================================================
Function Binary_Train_CV(y As Variant, x As Variant, Optional k_fold As Long = 10, _
        Optional learn_rate As Double = 0.001, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 2000, _
        Optional conv_max As Long = 5, Optional conv_tol As Double = 0.000001, _
        Optional loss_function As Variant, _
        Optional L2_max As Double = 1, Optional L2_best As Variant, _
        Optional adaptive_learn As Boolean = True) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long
Dim i_cv As Long, ii As Long, jj As Long, ii_max As Long, jj_max As Long
Dim tmp_x As Double, L1 As Double, L2 As Double
Dim loss() As Double, loss_min As Double
Dim accur() As Double, accur_max As Double
Dim y_output() As Double
Dim iArr() As Long, i_validate() As Long, i_train() As Long
Dim x_train() As Double, x_validate() As Double
Dim y_train() As Double, y_validate() As Double
Dim beta_prev() As Double, beta() As Double
Dim L2_list() As Double

    n = UBound(x, 1)            'number of observations
    n_dimension = UBound(x, 2)  'number of dimensions
    
    'Shuffle data set
    iArr = modMath.index_array(1, n)
    Call modMath.Shuffle(iArr)
    
    jj_max = 0: L2 = 0
    If L2_max > 0 Then jj_max = 6
    
    'ReDim accur(0 To jj_max)
    ReDim loss(0 To jj_max)
    ReDim beta_prev(1 To n_dimension + 1)
    
    'List of L2 to try
    ReDim L2_list(0 To jj_max)
    For jj = 1 To jj_max
        L2_list(jj) = L2_max * (5 ^ (jj - jj_max))
    Next jj
    
    'Outer loop for different L1 & L2 values
    For jj = jj_max To 0 Step -1
        L2 = L2_list(jj)
        
        'K-fold cross-validation
        For i_cv = 1 To k_fold
        
            DoEvents
            Application.StatusBar = "Binary_Train_CV: " & (jj_max - jj) & "/" & jj_max & ";" & i_cv & "/" & k_fold
                    
            Call modMath.CrossValidate_set(i_cv, k_fold, iArr, i_validate, i_train)
            Call modMath.Filter_Array(y, y_validate, i_validate)
            Call modMath.Filter_Array(x, x_validate, i_validate)
            Call modMath.Filter_Array(y, y_train, i_train)
            Call modMath.Filter_Array(x, x_train, i_train)
            
            beta = Binary_Train(y_train, x_train, learn_rate, momentum, _
                mini_batch, epoch_max, conv_max, conv_tol, , L2, adaptive_learn, False, beta_prev)
                
            y_output = Binary_Predict(beta, x_validate)
            'accur(jj) = accur(jj) + Accuracy(y_validate, y_output) * UBound(y_validate) / n
            loss(jj) = loss(jj) + Cross_Entropy(y_validate, y_output) / k_fold
            
            beta_prev = beta
        Next i_cv
    Next jj
    
    'Find L2 that gives lowest loss
    loss_min = Exp(70)
    For jj = 0 To jj_max
        If loss(jj) < loss_min Then
            loss_min = loss(jj)
            L2 = L2_list(jj)
        End If
    Next jj
    Debug.Print "Binary_Train_CV: Best L2 = " & L2 & ", loss=" & loss_min
    If IsMissing(L2_best) = False Then L2_best = L2
    
    'Find L2 that gives highest accuracy
'    accur_max = -Exp(70)
'        For jj = 0 To jj_max
'            If accur(jj) > accur_max Then
'                accur_max = accur(jj)
'                L2 = L2_list(jj)
'            End If
'        Next jj
'    Debug.Print "Binary_Train_CV: Best L2= " & L2 & ", accuracy=" & Format(accur_max, "0.0%")
    
    'Use selected L2 to train on whole data set
    beta = Binary_Train(y, x, learn_rate, momentum, _
            mini_batch, epoch_max, conv_max, conv_tol, loss, L2, adaptive_learn, True)
            
    If IsMissing(loss_function) = False Then loss_function = loss
    
    Binary_Train_CV = beta
    
    Erase x_train, x_validate, y_train, y_validate, beta, beta_prev
    Erase iArr, i_validate, i_train
    Erase loss, accur, y_output
    Application.StatusBar = False
End Function


'Calculate accuracy
Function Accuracy(y_tgt As Variant, y As Variant) As Double
Dim i As Long, n As Long
Dim tmp_x As Double
    n = UBound(y, 1)
    tmp_x = 0
    For i = 1 To n
        If y_tgt(i) >= 0.5 Then
            If y(i) >= 0.5 Then tmp_x = tmp_x + 1
        ElseIf y_tgt(i) < 0.5 Then
            If y(i) < 0.5 Then tmp_x = tmp_x + 1
        End If
    Next i
    Accuracy = tmp_x / n
End Function


'Calculate cross entropy
Function Cross_Entropy(y_tgt As Variant, y As Variant) As Double
Dim i As Long, n As Long
Dim tmp_x As Double
    n = UBound(y, 1)
    tmp_x = 0
    For i = 1 To n
        tmp_x = tmp_x - y_tgt(i) * Log(y(i)) - (1 - y_tgt(i)) * Log(1 - y(i))
    Next i
    Cross_Entropy = tmp_x / n
End Function


'===========================
'Output from logistic model
'===========================
'Output: y(1:N), binary classification output
'Input:  beta(1:D+1), regression coefficients of D-dimension, the D+1 element is the bias term
'        x(1:N,1:D), D-dimensional feature vector of N observations
'        force_binary, if set to true then y() will be rounded to exactly 0 or 1.
Function Binary_Predict(beta() As Double, x As Variant, Optional force_binary As Boolean = False) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim tmp_x As Double
Dim y() As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim y(1 To n)
    For i = 1 To n
        tmp_x = beta(n_dimension + 1)
        For j = 1 To n_dimension
            tmp_x = tmp_x + beta(j) * x(i, j)
        Next j
        y(i) = 1# / (1 + Exp(-tmp_x))
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
    Binary_Predict = y
    Erase y
End Function


Sub Rescale_beta(beta() As Double, x_mean() As Double, x_sd() As Double)
Dim i As Long, n As Long
    n = UBound(beta) - 1
    For i = 1 To n
        beta(i) = beta(i) / x_sd(i)
        beta(n + 1) = beta(n + 1) - beta(i) * x_mean(i)
    Next i
End Sub


Private Sub calc_gain(grad() As Double, grad_prev() As Double, gain() As Double, max_gain As Double)
Dim i As Long, n As Long
    n = UBound(grad)
    For i = 1 To n
        If Sgn(grad(i)) = Sgn(grad_prev(i)) Then
            gain(i) = gain(i) * 1.1
        Else
            gain(i) = gain(i) * 0.9
        End If
        If gain(i) > max_gain Then gain(i) = max_gain
        If gain(i) < 0.01 Then gain(i) = 0.01
    Next i
End Sub
