Attribute VB_Name = "mENET"
Option Explicit

'*******************************************
'Elastic Net implemented with coordinate descent
'*******************************************
'Main reference is "Regularization Paths for Generalized Linear Models via Coordinate Descent, Friedman 2008"
'*******************************************

'==== Elastic Net solved with coordinate descent
'Input: y(1:N), N observatons of response variable
'       x(1:N,1:D), N observations of D-dimensional predictors
'       lambda, total regularization strength
'       alpha, strength of L1 relative to total regularization, real value from 0 to 1
'       iter_max, maximum number of iterations in coordinates descent
'       isNormalize, set to TRUE if y() and x() already have zero mean and unit variance
'       init_beta, initial guess of beta, if skipped will be set to zero
'Output: beta(1:D+1), vector of regression coefficients of length D+1, where the last dimension is the intercept.
'        or beta(1:D) if isNormalize is set to TRUE so intercept is assumed to be zero.
Function Fit(ByVal y As Variant, ByVal x As Variant, Optional lambda As Double = 1, Optional alpha As Double = 0.5, _
        Optional iter_max As Long = 1000, Optional isNormalize As Boolean = False, _
        Optional init_beta As Variant, _
        Optional x_covar As Variant, Optional yx_covar As Variant, Optional isCalc As Variant, _
        Optional debugprint As Boolean = True) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim x_mean() As Double, x_sd() As Double
Dim y_mean As Double, y_sd As Double
Dim beta() As Double, beta_prev() As Double
Dim L1 As Double, L2 As Double

    L1 = lambda * alpha
    L2 = lambda * (1 - alpha)
    
    n = UBound(y, 1)
    If modMath.getDimension(x) = 1 Then
        n_dimension = 1
        Call modMath.Promote_Vec(x, x)
    Else
        n_dimension = UBound(x, 2)
    End If

    'scale to zero mean and variance
    If isNormalize = False Then
        Call normalize(y, x, y_mean, y_sd, x_mean, x_sd)
    End If
    
    'Use supplied beta() if avaiable
    If IsMissing(init_beta) = True Then
        ReDim beta(1 To n_dimension)
    Else
        beta = init_beta
    End If
    
    'Pre-calculate covariance of y vs x and also variance of x
    If IsMissing(yx_covar) = True Then Call Calc_Covar_yx(y, x, yx_covar)
    If IsMissing(x_covar) = True Then
        ReDim x_covar(1 To n_dimension, 1 To n_dimension)
        Call Calc_var_xx(x, x_covar)
    End If
    If IsMissing(isCalc) = True Then ReDim isCalc(1 To n_dimension)
    'Start coordinate descent
    
    For iterate = 1 To iter_max
        beta_prev = beta
        
        'Calculate covariance of active set
        For k = 1 To n_dimension
            If beta(k) <> 0 And isCalc(k) = 0 Then
                Call Calc_Covar_xx(x, x_covar, k, isCalc)
                isCalc(k) = 1
            End If
        Next k
        
        For k = 1 To n_dimension
            tmp_x = yx_covar(k)
            For j = 1 To n_dimension
                If j <> k And beta(j) <> 0 Then
                    tmp_x = tmp_x - beta(j) * x_covar(j, k)
                End If
            Next j
            If tmp_x > L1 Then
                beta(k) = (tmp_x - L1) / (x_covar(k, k) + L2)
            ElseIf tmp_x < (-L1) Then
                beta(k) = (tmp_x + L1) / (x_covar(k, k) + L2)
            Else
                beta(k) = 0
            End If
        Next k
        'terminate if incremental change in beta is smaller than tolerance
        k = 0
        For j = 1 To n_dimension
            If Abs(beta_prev(j) - beta(j)) < 0.0000001 Then k = k + 1
        Next j
        If k = n_dimension Then Exit For
    Next iterate
    
    If iterate >= iter_max Then
        If debugprint = True Then Debug.Print "ENET: failed to converge after " & (iterate - 1) & " iterations. lambda=" & lambda
    End If
    
    'Add intercept term if data is not normalized
    If isNormalize = False Then
        ReDim Preserve beta(1 To n_dimension + 1)
        Call restore_intercept(beta, y_mean, y_sd, x_mean, x_sd, n_dimension)
    End If
    
    Fit = beta
    Erase beta, beta_prev, x_mean, x_sd
End Function



'Returns the elastic net path with lambda from [0,lambda_max] with a given alpha between [0,1]
'Output: lambda(1:n_lambda), values of lambda tested
'        betas(1:n_lambda,1:D or 1:D+1), lasso path
Sub Fit_path(lambda() As Double, betas() As Double, alpha As Double, ByVal y As Variant, ByVal x As Variant, _
        Optional iter_max As Long = 1000, Optional isNormalize As Boolean = False, _
        Optional lambda_max As Double = 1, Optional n_lambda As Long = 10, Optional lambda_step As Double = 2)
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim x_mean() As Double, x_sd() As Double
Dim y_mean As Double, y_sd As Double
Dim beta() As Double, beta_tmp() As Double
Dim yx() As Double, z() As Double, isCalc() As Long
    n = UBound(y, 1)
    If modMath.getDimension(x) = 1 Then
        n_dimension = 1
        Call modMath.Promote_Vec(x, x)
    Else
        n_dimension = UBound(x, 2)
    End If

    'scale to zero mean and variance
    If isNormalize = False Then
        Call normalize(y, x, y_mean, y_sd, x_mean, x_sd)
    End If
    
    'values of lambda to try
    ReDim lambda(1 To n_lambda)
    For iterate = 2 To n_lambda
        lambda(iterate) = lambda_max * (lambda_step ^ (iterate - n_lambda))
    Next iterate
    
    'Initialize beta
    ReDim beta(1 To n_dimension)
    If isNormalize = False Then
        ReDim betas(1 To n_lambda, 1 To n_dimension + 1)
    Else
        ReDim betas(1 To n_lambda, 1 To n_dimension)
    End If
    
    'pre-calculate covariance of y vs x and variance of x
    ReDim isCalc(1 To n_dimension)
    ReDim z(1 To n_dimension, 1 To n_dimension)
    Call Calc_Covar_yx(y, x, yx)
    Call Calc_var_xx(x, z)
    
    For iterate = n_lambda To 1 Step -1
        beta = Fit(y, x, lambda(iterate), alpha, iter_max, True, beta, z, yx, isCalc)
        beta_tmp = beta
        If isNormalize = False Then
            ReDim Preserve beta_tmp(1 To n_dimension + 1)
            Call restore_intercept(beta_tmp, y_mean, y_sd, x_mean, x_sd, n_dimension)
        End If
        For j = 1 To n_dimension
            betas(iterate, j) = beta_tmp(j)
        Next j
        If isNormalize = False Then betas(iterate, n_dimension + 1) = beta_tmp(n_dimension + 1)
    Next iterate

    Erase beta, beta_tmp, x_mean, x_sd, z, yx
End Sub



'Returns the optimal lambda_best and corresponding beta using K-fold cross-validation
Function Fit_CV(ByVal y As Variant, ByVal x As Variant, lambda_best As Double, alpha_best As Double, _
        Optional k_fold As Long = 10, _
        Optional iter_max As Long = 1000, _
        Optional lambda_max As Double = 1, Optional n_lambda As Long = 10, Optional lambda_step As Double = 2, _
        Optional alpha_min As Double = 0, Optional alpha_max As Double = 1, Optional n_alpha As Long = 5, _
        Optional n_shuffle As Long = 1) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim iCV As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim iArr() As Long, iTrain() As Long, iTest() As Long
Dim y_train() As Double, y_test() As Double, y_out() As Double
Dim x_train() As Double, x_test() As Double
Dim x_mean() As Double, x_sd() As Double, y_mean As Double, y_sd As Double
Dim MSE_CV() As Double
Dim lambda() As Double, alpha() As Double, betas() As Double, beta() As Double

    n = UBound(y, 1)
    If modMath.getDimension(x) = 1 Then
        n_dimension = 1
        Call modMath.Promote_Vec(x, x)
    Else
        n_dimension = UBound(x, 2)
    End If
    
    'list of alpha to try
    If alpha_max > alpha_min Then
        ReDim alpha(1 To n_alpha)
        For i = 1 To n_alpha
            alpha(i) = ((n_alpha - i) * alpha_min + (i - 1) * alpha_max) / (n_alpha - 1)
        Next i
    ElseIf alpha_max < alpha_min Then
        Debug.Print "ENET:Fit_CV: alpha_max needs to be larger than alpha_min."
        Exit Function
    Else
        ReDim alpha(1 To 1)
        alpha(1) = alpha_min
        n_alpha = 1 'overide n_alpha
    End If
    
    ReDim MSE_CV(1 To n_lambda, 1 To n_alpha)
    For iterate = 1 To n_shuffle
        iArr = modMath.index_array(1, n)
        Call modMath.Shuffle(iArr)
        
        For iCV = 1 To k_fold
            DoEvents
            Application.StatusBar = "ENET:Fit_CV: " & iterate & "/" & n_shuffle & "; " & iCV & "/" & k_fold
            'Extract training set and validation set
            Call modMath.CrossValidate_set(iCV, k_fold, iArr, iTest, iTrain)
            Call modMath.Filter_Array(y, y_train, iTrain)
            Call modMath.Filter_Array(x, x_train, iTrain)
            Call modMath.Filter_Array(y, y_test, iTest)
            Call modMath.Filter_Array(x, x_test, iTest)
            
            For j = 1 To n_alpha
                'Fit path of current training set at current alpha
                Call Fit_path(lambda, betas, alpha(j), y_train, x_train, iter_max, False, lambda_max, n_lambda, lambda_step)
                'Test path on validation set
                For i = 1 To n_lambda
                    Call modMath.get_vector(betas, i, 1, beta)
                    y_out = Predict(beta, x_test, True, y_test, tmp_x)
                    MSE_CV(i, j) = MSE_CV(i, j) + tmp_x / (k_fold * n_shuffle)
                Next i
            Next j
        Next iCV
    Next iterate
    
    'Find lambda and alpha that gives smallest error
    tmp_x = Exp(70)
    For i = 1 To n_lambda
        For j = 1 To n_alpha
            If MSE_CV(i, j) < tmp_x Then
                lambda_best = lambda(i)
                alpha_best = alpha(j)
                tmp_x = MSE_CV(i, j)
            End If
        Next j
    Next i
    
    
    Fit_CV = Fit(y, x, lambda_best, alpha_best, iter_max, False)
    
    Application.StatusBar = False
    Erase betas, beta, iArr, iTrain, iTest, MSE_CV
    Erase x_train, x_test, y_train, y_test, y_out
End Function



Function Predict(beta() As Double, x As Variant, Optional hasIntercept As Boolean = True, _
        Optional y_tgt As Variant, Optional mse As Variant) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim y() As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim y(1 To n)
    If hasIntercept = True Then
         For i = 1 To n
            y(i) = beta(n_dimension + 1)
         Next i
    End If
    For i = 1 To n
        For j = 1 To n_dimension
            y(i) = y(i) + beta(j) * x(i, j)
        Next j
    Next i
    
    If IsMissing(y_tgt) = False Then
        mse = 0
        For i = 1 To n
            mse = mse + (y(i) - y_tgt(i)) ^ 2
        Next i
        mse = mse / n
    End If
    
    Predict = y
    Erase y
End Function


'Normalize y and x to zero mean and unit variance
Sub normalize(y As Variant, x As Variant, y_mean As Double, _
        y_sd As Double, x_mean() As Double, x_sd() As Double, Optional isKnown As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    If isKnown = True Then
        For i = 1 To n
            y(i) = (y(i) - y_mean) / y_sd
            For j = 1 To n_dimension
                x(i, j) = (x(i, j) - x_mean(j)) / x_sd(j)
            Next j
        Next i
        Exit Sub
    End If
    
    y_mean = 0
    y_sd = 0
    ReDim x_mean(1 To n_dimension)
    ReDim x_sd(1 To n_dimension)
    For j = 1 To n_dimension
        For i = 1 To n
            x_mean(j) = x_mean(j) + x(i, j)
            x_sd(j) = x_sd(j) + x(i, j) ^ 2
        Next i
        x_mean(j) = x_mean(j) / n
        x_sd(j) = Sqr(x_sd(j) / n - x_mean(j) ^ 2)
    Next j
    For i = 1 To n
        y_mean = y_mean + y(i)
        y_sd = y_sd + y(i) ^ 2
    Next i
    y_mean = y_mean / n
    y_sd = Sqr(y_sd / n - y_mean ^ 2)
    For i = 1 To n
        y(i) = (y(i) - y_mean) / y_sd
        For j = 1 To n_dimension
            x(i, j) = (x(i, j) - x_mean(j)) / x_sd(j)
        Next j
    Next i
End Sub



Sub normalize_x(x As Variant, x_mean() As Double, x_sd() As Double, Optional isKnown As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    If isKnown = True Then
        For i = 1 To n
            For j = 1 To n_dimension
                x(i, j) = (x(i, j) - x_mean(j)) / x_sd(j)
            Next j
        Next i
        Exit Sub
    End If
    ReDim x_mean(1 To n_dimension)
    ReDim x_sd(1 To n_dimension)
    For j = 1 To n_dimension
        For i = 1 To n
            x_mean(j) = x_mean(j) + x(i, j)
            x_sd(j) = x_sd(j) + x(i, j) ^ 2
        Next i
        x_mean(j) = x_mean(j) / n
        x_sd(j) = Sqr(x_sd(j) / n - x_mean(j) ^ 2)
    Next j
    For i = 1 To n
        For j = 1 To n_dimension
            x(i, j) = (x(i, j) - x_mean(j)) / x_sd(j)
        Next j
    Next i
End Sub

Sub restore_intercept(beta() As Double, _
    y_mean As Double, y_sd As Double, _
    x_mean() As Double, x_sd() As Double, n_dimension As Long)
Dim j As Long
    beta(n_dimension + 1) = y_mean
    For j = 1 To n_dimension
        beta(j) = beta(j) * y_sd / x_sd(j)
        beta(n_dimension + 1) = beta(n_dimension + 1) - beta(j) * x_mean(j)
    Next j
End Sub


Private Sub Calc_Covar_yx(y As Variant, x As Variant, yx_covar As Variant)
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim tmp_x As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim yx_covar(1 To n_dimension)
    For j = 1 To n_dimension
        tmp_x = 0
        For i = 1 To n
            tmp_x = tmp_x + x(i, j) * y(i)
        Next i
        yx_covar(j) = tmp_x / n
    Next j
End Sub

Private Sub Calc_Covar_xx(x As Variant, x_covar As Variant, k As Long, isCalc As Variant)
Dim i As Long, j As Long, n As Long, n_dimension As Long
Dim tmp_x As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    For j = 1 To n_dimension
        If j <> k And isCalc(j) = 0 Then
            tmp_x = 0
            For i = 1 To n
                tmp_x = tmp_x + x(i, j) * x(i, k)
            Next i
            x_covar(k, j) = tmp_x / n
            x_covar(j, k) = x_covar(k, j)
        End If
    Next j
End Sub

Private Sub Calc_var_xx(x As Variant, x_covar As Variant)
Dim i As Long, j As Long, n As Long, n_dimension As Long
Dim tmp_x As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    For j = 1 To n_dimension
        tmp_x = 0
        For i = 1 To n
            tmp_x = tmp_x + x(i, j) ^ 2
        Next i
        x_covar(j, j) = tmp_x / n
    Next j
End Sub

Private Function beta_conv(beta() As Double, beta_prev() As Double, Optional tol As Double = 0.000001) As Boolean
Dim i As Long, k As Long, m As Long
    k = 0
    m = UBound(beta)
    For i = 1 To m
        If Abs(beta(i) - beta_prev(i)) < tol Then k = k + 1
    Next i
    If k = m Then
        beta_conv = True
    Else
        beta_conv = False
    End If
End Function


Sub Test_ENET()
Dim i As Long, j As Long, n As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double, s_optimal As Double, lambda_optimal As Double
Dim x() As Double, y() As Double, x_mean() As Double, x_sd() As Double, y_mean As Double, y_sd As Double
Dim beta() As Double, tmp_vec() As Double
Dim A() As Long, lambda() As Double
Dim strFactor() As String
Dim mywkbk As Workbook

    msgbox "Testing Elastic Net A"

    Set mywkbk = ActiveWorkbook
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    '=== Prostate Cancer Data
    With mywkbk.Sheets("Data_Prostate") 'Diabetes") '
        n = .Range("A100000").End(xlUp).Row - 1
        n_dimension = .Range("IV1").End(xlToLeft).Column - 1
        ReDim strFactor(1 To n_dimension, 1 To 1)
        ReDim x(1 To n, 1 To n_dimension)
        ReDim y(1 To n)
        For i = 1 To n
            y(i) = .Cells(1 + i, n_dimension + 1)
            For j = 1 To n_dimension
                x(i, j) = .Cells(1 + i, j).Value
            Next j
        Next i
        For j = 1 To n_dimension
            strFactor(j, 1) = .Cells(1, j).Text
        Next j
    End With
    '=============

    Call mENET.normalize(y, x, y_mean, y_sd, x_mean, x_sd)

    beta = mENET.Fit_CV(y, x, tmp_x, tmp_y, 10, , 2, 80, 1.1, 0, 1, 10, 10)
    Debug.Print "best (Lambda, alpha)=" & Format(tmp_x, "0.0000") & ", " & tmp_y
    For j = 1 To UBound(beta)
        Debug.Print j & ", " & beta(j)
    Next j

    Call mENET.Fit_path(lambda, beta, 0.5, y, x, , False, 2, 80, 1.1)
    With mywkbk.Sheets("ENET")
        .Range("A3:J10000").Clear
        .Range("A3").Resize(UBound(lambda), 1).Value = modMath.wkshtTranspose(lambda)
        .Range("B3").Resize(UBound(beta, 1), UBound(beta, 2)).Value = beta
    End With

    Set mywkbk = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


'=========================
'Logistic regionsion for binary response y()

Function Logistic_Fit(ByVal y As Variant, ByVal x As Variant, _
        Optional lambda As Double = 1, Optional alpha As Double = 0.5, _
        Optional iter_max As Long = 1000, Optional init_beta As Variant) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long, iterate As Long
Dim tmp_x As Double, p As Double, w As Double
Dim x_mean() As Double, x_sd() As Double, z() As Double, xx() As Double
Dim beta() As Double, beta_prev() As Double
    
    n = UBound(y, 1)
    If modMath.getDimension(x) = 1 Then
        n_dimension = 1
        Call modMath.Promote_Vec(x, x)
    Else
        n_dimension = UBound(x, 2)
    End If
    
    'Use supplied beta() if available
    If IsMissing(init_beta) = True Then
        ReDim beta(1 To n_dimension + 1)
    Else
        beta = init_beta
    End If
    
    beta_prev = beta
    For iterate = 1 To iter_max
        
        'caculate logit response with current beta
        ReDim z(1 To n)
        ReDim xx(1 To n, 1 To n_dimension)
        For i = 1 To n
            tmp_x = beta(n_dimension + 1)
            For j = 1 To n_dimension
                tmp_x = tmp_x + x(i, j) * beta(j)
            Next j
            p = 1 / (1 + Exp(-tmp_x))
            If p < 0.00001 Then
                p = 0
                w = 0.00001
            ElseIf p > 0.99999 Then
                p = 1
                w = 0.00001
            Else
                w = p * (1 - p)
            End If
            
            z(i) = Sqr(w) * (tmp_x + (y(i) - p) / w)
            For j = 1 To n_dimension
                xx(i, j) = Sqr(w) * x(i, j)
            Next j
        Next i
        
        beta = Fit(z, xx, lambda, alpha, iter_max, False, beta_prev, , , , False)
        
        Erase z, xx
        
        'terminate if incremental change in beta is smaller than tolerance
        k = 0
        For j = 1 To n_dimension + 1
            If Abs(beta_prev(j) - beta(j)) < 0.000001 Then k = k + 1
        Next j
        If k = n_dimension + 1 Then Exit For

        beta_prev = beta
        
    Next iterate
    
    If iterate >= iter_max Then
        Debug.Print "ENET: Logistic_Fit: failed to converge after " & (iterate - 1) & " iterations."
    End If
    
    Logistic_Fit = beta
    Erase beta, beta_prev, x_mean, x_sd
End Function


Sub Logistic_Fit_path(lambda() As Double, betas() As Double, alpha As Double, ByVal y As Variant, ByVal x As Variant, _
        Optional iter_max As Long = 1000, _
        Optional lambda_min As Double = 0.001, Optional lambda_max As Double = 1, Optional n_lambda As Long = 20)
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim x_mean() As Double, x_sd() As Double
Dim y_mean As Double, y_sd As Double
Dim beta() As Double, beta_tmp() As Double
Dim yx() As Double, z() As Double, isCalc() As Long
    n = UBound(y, 1)
    If modMath.getDimension(x) = 1 Then
        n_dimension = 1
        Call modMath.Promote_Vec(x, x)
    Else
        n_dimension = UBound(x, 2)
    End If

    'values of lambda to try
    ReDim lambda(1 To n_lambda)
    lambda(1) = lambda_min
    lambda(n_lambda) = lambda_max
    tmp_x = (Log(lambda_max) - Log(lambda_min)) / (n_lambda - 1)
    For iterate = 2 To n_lambda - 1
        lambda(iterate) = Exp(Log(lambda_min) + (iterate - 1) * tmp_x)
    Next iterate
    
    'Initialize beta
    ReDim beta(1 To n_dimension + 1)
    ReDim betas(1 To n_lambda, 1 To n_dimension + 1)
    
    For iterate = n_lambda To 1 Step -1
        beta = Logistic_Fit(y, x, lambda(iterate), alpha, iter_max, beta)
        For j = 1 To n_dimension + 1
            betas(iterate, j) = beta(j)
        Next j
    Next iterate

    Erase beta, beta_tmp, x_mean, x_sd, z, yx
End Sub


Function Logistic_Predict(beta() As Double, x As Variant, Optional force_binary = False, _
        Optional y_tgt As Variant, Optional accuracy As Variant, Optional cross_entropy As Variant) As Double()
Dim i As Long, j As Long, k As Long, n As Long, n_dimension As Long
Dim tmp_x As Double
Dim y() As Double

    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim y(1 To n)

    For i = 1 To n
        tmp_x = beta(n_dimension + 1)
        For j = 1 To n_dimension
            tmp_x = tmp_x + x(i, j) * beta(j)
        Next j
        y(i) = 1 / (1 + Exp(-tmp_x))
    Next i
    
    If IsMissing(y_tgt) = False Then
        If IsMissing(accuracy) = False Then
            accuracy = 0
            For i = 1 To n
                If y_tgt(i) >= 0.5 Then
                    If y(i) >= 0.5 Then accuracy = accuracy + 1
                Else
                    If y(i) < 0.5 Then accuracy = accuracy + 1
                End If
            Next i
            accuracy = accuracy / n
        End If
        If IsMissing(cross_entropy) = False Then
            cross_entropy = 0
            For i = 1 To n
                cross_entropy = cross_entropy - y_tgt(i) * Log(y(i)) - (1 - y_tgt(i)) * Log(1 - y(i))
            Next i
            cross_entropy = cross_entropy / n
        End If
    End If
    
    If force_binary = True Then
         For i = 1 To n
            If y(i) >= 0.5 Then
                y(i) = 1
            Else
                y(i) = 0
            End If
         Next i
    End If
    
    Logistic_Predict = y
    Erase y
    
End Function




'Returns the optimal lambda_best and corresponding beta using K-fold cross-validation
Function Logistic_Fit_CV(ByVal y As Variant, ByVal x As Variant, lambda_best As Double, alpha_best As Double, _
        Optional k_fold As Long = 10, _
        Optional iter_max As Long = 1000, _
        Optional lambda_min As Double = 0.001, Optional lambda_max As Double = 1, Optional n_lambda As Long = 20, _
        Optional alpha_min As Double = 0, Optional alpha_max As Double = 1, Optional n_alpha As Long = 5, _
        Optional n_shuffle As Long = 1) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, n_dimension As Long
Dim iCV As Long, iterate As Long
Dim tmp_x As Double, tmp_y As Double
Dim iArr() As Long, iTrain() As Long, iTest() As Long
Dim y_train() As Double, y_test() As Double, y_out() As Double
Dim x_train() As Double, x_test() As Double
Dim x_mean() As Double, x_sd() As Double, y_mean As Double, y_sd As Double
Dim LOSS_CV() As Double, ACC_CV() As Double
Dim lambda() As Double, alpha() As Double, betas() As Double, beta() As Double

    n = UBound(y, 1)
    If modMath.getDimension(x) = 1 Then
        n_dimension = 1
        Call modMath.Promote_Vec(x, x)
    Else
        n_dimension = UBound(x, 2)
    End If
    
    'list of alpha to try
    If alpha_max > alpha_min Then
        ReDim alpha(1 To n_alpha)
        For i = 1 To n_alpha
            alpha(i) = ((n_alpha - i) * alpha_min + (i - 1) * alpha_max) / (n_alpha - 1)
        Next i
    ElseIf alpha_max < alpha_min Then
        Debug.Print "ENET:Logistic_Fit_CV: alpha_max needs to be larger than alpha_min."
        Exit Function
    Else
        ReDim alpha(1 To 1)
        alpha(1) = alpha_min
        n_alpha = 1 'overide n_alpha
    End If
    
    ReDim LOSS_CV(1 To n_lambda, 1 To n_alpha)
    ReDim ACC_CV(1 To n_lambda, 1 To n_alpha)
    For iterate = 1 To n_shuffle
        iArr = modMath.index_array(1, n)
        Call modMath.Shuffle(iArr)
        
        For iCV = 1 To k_fold
            DoEvents
            Application.StatusBar = "ENET:Logistic_Fit_CV: " & iterate & "/" & n_shuffle & "; " & iCV & "/" & k_fold
            'Extract training set and validation set
            Call modMath.CrossValidate_set(iCV, k_fold, iArr, iTest, iTrain)
            Call modMath.Filter_Array(y, y_train, iTrain)
            Call modMath.Filter_Array(x, x_train, iTrain)
            Call modMath.Filter_Array(y, y_test, iTest)
            Call modMath.Filter_Array(x, x_test, iTest)
            m = UBound(iTest)
            For j = 1 To n_alpha
                'Fit path of current training set at current alpha
                Call Logistic_Fit_path(lambda, betas, alpha(j), y_train, x_train, iter_max, _
                    lambda_min, lambda_max, n_lambda)
                'Test path on validation set
                For i = 1 To n_lambda
                    Call modMath.get_vector(betas, i, 1, beta)
                    y_out = Logistic_Predict(beta, x_test, False, y_test, tmp_x, tmp_y)
                    LOSS_CV(i, j) = LOSS_CV(i, j) + tmp_y / (k_fold * n_shuffle)
                    ACC_CV(i, j) = ACC_CV(i, j) + (m * tmp_x / n) / (n_shuffle)
                Next i
            Next j
        Next iCV
    Next iterate
    
    'Find lambda and alpha that gives smallest error
    tmp_x = Exp(70)
    For i = 1 To n_lambda
        For j = 1 To n_alpha
            If LOSS_CV(i, j) < tmp_x Then
                lambda_best = lambda(i)
                alpha_best = alpha(j)
                tmp_x = LOSS_CV(i, j)
            End If
        Next j
    Next i
    
    
    Logistic_Fit_CV = Logistic_Fit(y, x, lambda_best, alpha_best, iter_max)
    
    Application.StatusBar = False
    Erase betas, beta, iArr, iTrain, iTest, LOSS_CV, ACC_CV
    Erase x_train, x_test, y_train, y_test, y_out
End Function


Sub Test_Logistic()
Dim i As Long, j As Long, k As Long, n As Long
Dim n_dimension As Long, n_train As Long, n_test As Long
Dim factor_label() As String
Dim x_mean() As Double, x_sd() As Double
Dim x_train() As Double, x_test() As Double
Dim y_train() As Double, y_test() As Double
Dim beta() As Double, conv_chk() As Double
Dim mywkbk As Workbook
Dim strDir As String
Dim tmp_x As Double, tmp_y As Double, lambda As Double, alpha As Double
Dim lambdas() As Double, betas() As Double

    Set mywkbk = ActiveWorkbook
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    'Read training & test data
    strDir = "C:\Users\yp73275\Desktop\Logistic_Reg\"
    Call modMath.Read_csv(strDir & "Online_Courses_training.csv", y_train, x_train, , , factor_label)
    Call modMath.Read_csv(strDir & "Online_Courses_testing.csv", y_test, x_test)
    n_dimension = UBound(x_train, 2)
    n_train = UBound(y_train, 1)
    n_test = UBound(y_test, 1)
    ReDim Preserve factor_label(1 To n_dimension + 1)
    factor_label(n_dimension + 1) = "bias"
    
    'Normalize data to zero mean and unit variance
    Call modMath.normalize_x(x_train, x_mean, x_sd)
    Call modMath.normalize_x(x_test, x_mean, x_sd, , True)
    
    msgbox "AA"
    
    'Test Cross-validation
    Call mENET.Logistic_Fit_CV(y_train, x_train, lambda, alpha, 10, , 0.01, 2, , , , , 1)
    Debug.Print "Best (lambda,alpha)= ( " & Format(lambda, "0.0000") & ", " & alpha & " )"
    
'    alpha = 0.3
'    lambda = 0.001
    
    Call mENET.Logistic_Fit_path(lambdas, betas, alpha, y_train, x_train, lambda_min:=0.01, lambda_max:=2, n_lambda:=50) 'Test Path
    beta = mENET.Logistic_Fit(y_train, x_train, lambda, alpha) 'Train selected model
    
    'Output results
    With mywkbk.Sheets("LOGISTIC")
        'Output path
        .Range("Q3:AI50000").ClearContents
        .Range("Q3").Resize(UBound(lambdas, 1), 1).Value = modMath.wkshtTranspose(lambdas)
        .Range("R2").Resize(1, UBound(factor_label, 1)).Value = factor_label
        .Range("R3").Resize(UBound(betas, 1), UBound(betas, 2)).Value = betas
        
        'Output predictions and selected model
        .Range("A3:H50000").ClearContents
        .Range("E3").Resize(n_dimension + 1, 1).Value = modMath.wkshtTranspose(factor_label)
        .Range("F3").Resize(n_dimension + 1, 1).Value = modMath.wkshtTranspose(beta)
        .Range("G3").Resize(n_dimension, 1).Value = modMath.wkshtTranspose(x_mean)
        .Range("H3").Resize(n_dimension, 1).Value = modMath.wkshtTranspose(x_sd)
        
        .Range("A3").Resize(n_train, 1).Value = modMath.wkshtTranspose(y_train)
        .Range("B3").Resize(n_train, 1).Value = modMath.wkshtTranspose(mENET.Logistic_Predict(beta, x_train))
    
        .Range("C3").Resize(n_test, 1).Value = modMath.wkshtTranspose(y_test)
        .Range("D3").Resize(n_test, 1).Value = modMath.wkshtTranspose(mENET.Logistic_Predict(beta, x_test))
    End With
    
    Erase x_mean, x_sd, x_train, x_test, y_train, y_test
    Set mywkbk = Nothing
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub
