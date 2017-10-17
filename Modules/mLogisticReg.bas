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
        Optional learn_rate As Double = 0.01, Optional momentum As Double = 0.5, _
        Optional mini_batch As Long = 5, _
        Optional epoch_max As Long = 1000, _
        Optional conv_max As Long = 10, Optional conv_tol As Double = 0.0000001, _
        Optional loss_function As Variant)
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long, ii As Long
Dim batch_count As Long, epoch As Long, conv_count As Long
Dim tmp_x As Double, delta As Double, y_output As Double
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
    
        If epoch Mod 10 = 0 Then
            DoEvents
            Application.StatusBar = "mLogisticReg: Binary_Train: " & epoch & "/" & epoch_max
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
                    beta_chg(j) = momentum * beta_chg(j) - grad(j) * learn_rate / batch_count
                    beta(j) = beta(j) + beta_chg(j)
                Next j
                'reset mini batch count and gradient
                batch_count = 0
                ReDim grad(1 To n_dimension + 1)
            End If
            
        Next ii
        
        loss(epoch) = loss(epoch) / n
        
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


'===========================
'Output from logistic model
'===========================
'Output: y(1:N), binary classification output
'Input:  beta(1:D+1), regression coefficients of D-dimension, the D+1 element is the bias term
'        x(1:N,1:D), D-dimensional feature vector of N observations
Function Binary_InOut(beta() As Double, x As Variant) As Double()
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
    Binary_InOut = y
    Erase y
End Function
