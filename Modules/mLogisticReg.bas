Attribute VB_Name = "mLogisticReg"
Option Explicit

'========================================================
'Perform binary logistic regression with gradient descent
'========================================================
'Output: beta(1:D+1), regression coefficients of D-dimension, the D+1 element is the bias term
'Input: y(1:N), binary target N observations
'       x(1:N,1:D), D-dimensional feature vector of N observations
'       learn_rate, learning rate for gradient descent
Sub Binary_Train(beta() As Double, y As Variant, x As Variant, _
        Optional learn_rate As Double = 0.001, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 1000, _
        Optional conv_max As Long = 5, Optional conv_tol As Double = 0.0000001, _
        Optional loss_function As Variant, Optional L1 As Double = 0, Optional L2 As Double = 0, Optional show_progress As Boolean = True)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long, ii As Long
Dim batch_count As Long, epoch As Long, conv_count As Long
Dim tmp_x As Double, tmp_y As Double, delta As Double, y_output As Double
Dim loss() As Double, grad() As Double, beta_chg() As Double
Dim iArr() As Long
    n = UBound(x, 1)            'number of observations
    n_dimension = UBound(x, 2)  'number of dimensions
    
    'Randomly initialize beta()
    Randomize
    ReDim beta(1 To n_dimension + 1)
    ReDim beta_chg(1 To n_dimension + 1)
    For j = 1 To n_dimension + 1
        beta(j) = -1 + 2 * Rnd()
    Next j
    
    'Perform gradient descent
    conv_count = 0
    ReDim loss(1 To epoch_max)
    For epoch = 1 To epoch_max
        
        If show_progress = True Then
            If epoch Mod 10 = 0 Then
                DoEvents
                Application.StatusBar = "mLogisticReg: Binary_Train: " & epoch & "/" & epoch_max
            End If
        End If
        
        'Shuffle data set
        iArr = modMath.index_array(1, n)
        Call modMath.Shuffle(iArr)
        
        'Scan through dataset
        batch_count = 0
        ReDim grad(1 To n_dimension + 1)
        For ii = 1 To n
            i = iArr(ii)
            
            'beta dot x
            tmp_x = beta(n_dimension + 1)
            For j = 1 To n_dimension
                tmp_x = tmp_x + beta(j) * x(i, j)
            Next j
            
            y_output = 1# / (1 + Exp(-tmp_x)) 'Sigmoid function
            loss(epoch) = loss(epoch) - y(i) * Log(y_output) - (1 - y(i)) * Log(1 - y_output) 'accumulate loss function
            
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
                If L1 > 0 Then 'L1-regularization
                    For j = 1 To n_dimension
                        grad(j) = grad(j) + L1 * Sgn(beta(j))
                    Next j
                End If
                If L2 > 0 Then 'L2-regularization
                    For j = 1 To n_dimension
                        grad(j) = grad(j) + L2 * beta(j)
                    Next j
                End If
                
                For j = 1 To n_dimension + 1
                    beta_chg(j) = momentum * beta_chg(j) - grad(j) * learn_rate
                    beta(j) = beta(j) + beta_chg(j)
                Next j
                
                'reset mini batch count and gradient
                batch_count = 0
                ReDim grad(1 To n_dimension + 1)
            End If
            
        Next ii
        
        loss(epoch) = loss(epoch) / n
        
        If L1 > 0 Then
            tmp_x = 0
            For j = 1 To n_dimension
                tmp_x = tmp_x + Abs(beta(j))
            Next j
            loss(epoch) = loss(epoch) + L1 * tmp_x
        End If
        
        If L2 > 0 Then
            tmp_x = 0
            For j = 1 To n_dimension
                tmp_x = tmp_x + beta(j) ^ 2
            Next j
            loss(epoch) = loss(epoch) + L2 * tmp_x
        End If
        
        'early terminate on convergence
        If epoch > 1 Then
            If loss(epoch) < loss(epoch - 1) Then
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
    
    If IsMissing(loss_function) = False Then loss_function = loss
    Erase loss, grad, beta_chg, iArr
    Application.StatusBar = False
End Sub


'========================================================
'Perform K-fold crossvalidation to find optimal L1 & L2 regularization
'========================================================
Sub Binary_Train_CV(beta() As Double, y As Variant, x As Variant, Optional K_fold As Long = 10, _
        Optional learn_rate As Double = 0.001, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 1000, _
        Optional conv_max As Long = 5, Optional conv_tol As Double = 0.0000001, _
        Optional loss_function As Variant, _
        Optional L1_max As Double = 0.01, Optional L2_max As Double = 2)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long
Dim i_cv As Long, ii As Long, jj As Long
Dim n_train As Long, n_validate As Long
Dim tmp_x As Double, L1 As Double, L2 As Double
Dim loss() As Double, tmp_vec() As Double, y_output() As Double
Dim accur() As Double, accur_max As Double
Dim iArr() As Long, i_validate() As Long, i_train() As Long
Dim x_train() As Double, x_validate() As Double
Dim y_train() As Double, y_validate() As Double
    n = UBound(x, 1)            'number of observations
    n_dimension = UBound(x, 2)  'number of dimensions
    n_validate = n \ K_fold
    n_train = n - n_validate
    
    'Shuffle data set
    iArr = modMath.index_array(1, n)
    Call modMath.Shuffle(iArr)

    ReDim accur(0 To 5, 0 To 5)
    
    For ii = 0 To 5
        L1 = ii * L1_max / 5 'Try L1 values
        For jj = 0 To 5
            L2 = jj * L2_max / 5  'Try L2 values
            
            'K-fold cross-validation
            For i_cv = 1 To K_fold
            
                DoEvents
                Application.StatusBar = "Binary_Train_CV: " & ii & "/" & 5 & _
                        " ; " & jj & "/" & 5 & ";" & i_cv & "/" & K_fold
                
                ReDim i_validate(1 To n_validate)
                ReDim i_train(1 To n_train)
                For i = 1 To n_validate
                    i_validate(i) = iArr((i_cv - 1) * n_validate + i)
                Next i
                j = 0
                For i = 1 To n
                    If i <= ((i_cv - 1) * n_validate) Or _
                        i > (i_cv * n_validate) Then
                        j = j + 1
                        i_train(j) = i
                    End If
                Next i
                If i_cv = K_fold And (i_cv * n_validate) < n Then
                    m = n - (i_cv * n_validate)
                    ReDim Preserve i_validate(1 To n_validate + m)
                    ReDim Preserve i_train(1 To n_train - m)
                    For i = 1 To m
                        i_validate(n_validate + i) = iArr(i_cv * n_validate + i)
                    Next i
                End If
                Call modMath.Filter_Array(y, y_validate, i_validate)
                Call modMath.Filter_Array(x, x_validate, i_validate)
                Call modMath.Filter_Array(y, y_train, i_train)
                Call modMath.Filter_Array(x, x_train, i_train)
                
                Call Binary_Train(tmp_vec, y_train, x_train, learn_rate, momentum, _
                    mini_batch, epoch_max, conv_max, conv_tol, , L1, L2, False)
                    
                y_output = Binary_InOut(tmp_vec, x_validate)
                accur(ii, jj) = accur(ii, jj) + Accuracy(y_validate, y_output) * UBound(y_validate) / n
                
            Next i_cv
        Next jj
    Next ii
    
    'Find L1 & L2 that gives highest accuracy
    accur_max = -Exp(70)
    For ii = 0 To 5
        For jj = 0 To 5
            If accur(ii, jj) > accur_max Then
                accur_max = accur(ii, jj)
                L1 = ii * L1_max / 5
                L2 = jj * L2_max / 5
            End If
        Next jj
    Next ii
    Debug.Print "Binary_Train_CV: Best(L1,L2)= (" & L1 & ", " & L2; "), accuracy=" & Format(accur_max, "0.0%")
    
    'Use selected L1 & L2 to train on whole data set
    Call Binary_Train(beta, y, x, learn_rate, momentum, _
            mini_batch, epoch_max, conv_max, conv_tol, loss, L1, L2, True)
            
    If IsMissing(loss_function) = False Then loss_function = loss
    
    Application.StatusBar = False
End Sub


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
Function Binary_InOut(beta() As Double, x As Variant, Optional force_binary As Boolean = False) As Double()
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
    Binary_InOut = y
    Erase y
End Function
