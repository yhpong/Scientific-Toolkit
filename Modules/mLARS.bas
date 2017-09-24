Attribute VB_Name = "mLARS"
Option Explicit
'Requires: modMath



'====================================================
'Least Angle Regression (LARS)
'====================================================
'Main Reference: "Least Angle Regression", Efron (2003)
'https://web.stanford.edu/~hastie/Papers/LARS/LeastAngle_2002.pdf
'Implementation from http://www.ece.ubc.ca/~xiaohuic/code/LARS/lars.m
'Reference below shows how the constraint can be expressed in max-norm form or Lagrange from
'https://stats.stackexchange.com/questions/207484/lasso-regularisation-parameter-from-lars-algorithm
'Input:     x() is a NxD predictor variables with zero mean and uni length, i.e. sum(x^2)=1
'           y() is a length N vector of response variable
'           LASSO, TRUE if LASSO regression needs to be performed
'           norm1, if left empty, LARS return the whole solution path same as beta()
'               if given, LARS returns beta as a vector as the solution that satisfy:
'                   if norm_rel is set to 1, |beta|=norm1*|beta_ols|
'                   if norm_rel is set to 0, |beta|=norm1
'Output:    beta(), solution path of dimension (1 to n_dimension, 1 to number of steps)
'           A(),    order of which each dimension is added to the model
Function LARS(x() As Double, y() As Double, beta() As Double, A() As Long, _
        Optional LASSO As Boolean = False, Optional norm1 As Variant, Optional norm_rel As Long = 1) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long
Dim n_raw As Long, n_dimension As Long, m_dimension As Long
Dim addVar As Long, n_A As Long, n_Ac As Long, signOK As Long
Dim tmp_x As Double, tmp_y As Double, c_max As Double, INFINITY As Double, gamma2 As Double, tol As Double
Dim beta_tmp() As Double, correl() As Double
Dim mu() As Double, mu_prev() As Double, gamma() As Double
Dim G_A() As Double, s_A() As Double, L_A As Double, w_A() As Double, u_A() As Double
Dim aa() As Double, x_A() As Double
Dim Ac() As Long
Dim t_prev As Double, t_now As Double, beta_t() As Double, t_tgt As Double
Dim tmp_vec() As Double
Dim x_mean() As Double, x_scale() As Double, y_mean As Double

    tol = 0.0000000001
    INFINITY = Exp(70)
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    m_dimension = min2(n_dimension, n_raw - 1)
    
    If IsMissing(norm1) = False Then
        If norm_rel = 1 Then
            Call modMath.linear_regression(y, x, beta_t)
            tmp_x = 0
            For i = 1 To n_dimension
                tmp_x = tmp_x + Abs(beta_t(i))
            Next i
            t_tgt = norm1 * tmp_x
        Else
            t_tgt = norm1
        End If
    End If
    
    ReDim beta(1 To n_dimension, 1 To 1)
    ReDim gamma(0 To 0)
    ReDim mu(1 To n_raw)
    ReDim mu_prev(1 To n_raw)
    ReDim x_A(1 To n_raw, 1 To 1)
    ReDim A(0 To 0)

    t_prev = 0
    ReDim beta_t(1 To n_dimension)
    n_A = 0
    signOK = 1
    iterate = 0
    
    Do While n_A < m_dimension
        iterate = iterate + 1
    
        Call LARS_Correl(correl, c_max, j, x, y, mu)
        If c_max < tol Then Exit Do
        If iterate = 1 Then addVar = j
        
        If signOK = 1 Then
            Call Append_1D(A, addVar)
            n_A = n_A + 1
        End If
        
        Ac = InActiveSet(A, n_dimension)
        n_Ac = UBound(Ac)
    
        ReDim s_A(1 To n_A)
        ReDim x_A(1 To n_raw, 1 To n_A)
        For i = 1 To n_A
            s_A(i) = Sgn(correl(A(i)))
        Next i
        For i = 1 To n_raw
            For j = 1 To n_A
                x_A(i, j) = x(i, A(j))
            Next j
        Next i

        G_A = modMath.M_Dot(x_A, x_A, 1, 0)
        G_A = modMath.Matrix_Inverse_Cholesky(G_A) 'How to speed up by reusing cholesky factorization?
        L_A = 1 / Sqr(modMath.VV_dot(s_A, modMath.M_Dot(G_A, s_A)))
        w_A = modMath.M_scalar_dot(modMath.M_Dot(G_A, s_A), L_A)
        u_A = modMath.M_Dot(x_A, w_A)
        aa = modMath.M_Dot(x, u_A, 1, 0)

        If n_A = m_dimension Then
            Call Append_1D(gamma, c_max / L_A)
        Else
            Call Append_1D(gamma, INFINITY)
            For i = 1 To n_Ac
                j = Ac(i)
                tmp_x = (c_max - correl(j)) / (L_A - aa(j))
                tmp_y = (c_max + correl(j)) / (L_A + aa(j))
                If tmp_x <= 0 Then tmp_x = INFINITY
                If tmp_y <= 0 Then tmp_y = INFINITY
                tmp_x = min2(tmp_x, tmp_y)
                If tmp_x < gamma(iterate) Then
                    gamma(iterate) = tmp_x
                    addVar = j
                End If
            Next i
        End If
        
        ReDim beta_tmp(1 To n_dimension)
        For i = 1 To n_A
            beta_tmp(A(i)) = beta(A(i), iterate) + gamma(iterate) * w_A(i)
        Next i
        
        If LASSO = True Then
            signOK = 1
            ReDim gammatest(1 To n_A)
            For i = 1 To n_A
                gammatest(i) = -beta(A(i), iterate) / w_A(i)
            Next i
            gamma2 = INFINITY
            For i = 1 To n_A
                If gammatest(i) > 0 Then
                    If gammatest(i) < gamma2 Then
                        gamma2 = gammatest(i)
                        k = A(i)
                    End If
                End If
            Next i
            If gamma2 < gamma(iterate) Then
                gamma(iterate) = gamma2
                For i = 1 To n_A
                    beta_tmp(A(i)) = beta(A(i), iterate) + gamma(iterate) * w_A(i)
                Next i
                beta_tmp(k) = 0
                Call Eject_1D(A, k)
                n_A = n_A - 1
                signOK = 0
            End If
        End If
        
        For i = 1 To n_raw
            mu(i) = mu_prev(i) + gamma(iterate) * u_A(i)
        Next i
        mu_prev = mu
        ReDim Preserve beta(1 To n_dimension, 1 To iterate + 1)
        For i = 1 To n_dimension
            beta(i, iterate + 1) = beta_tmp(i)
        Next i
        
        If IsMissing(norm1) = False Then
            t_now = 0
            For i = 1 To n_A
                t_now = t_now + Abs(beta_tmp(A(i)))
            Next i
            If t_prev < t_tgt And t_tgt <= t_now Then
                For i = 1 To n_A
                    beta_t(A(i)) = beta(A(i), iterate) + L_A * (t_tgt - t_prev) * w_A(i)
                Next i
                Exit Do
            End If
            t_prev = t_now
        End If
        
    Loop
    
    If IsMissing(norm1) = False Then
        LARS = beta_t
    Else
'        LARS = beta
        Call modMath.Filter_Array(beta, LARS, , UBound(beta, 2))
    End If
    
    Erase x_A, G_A, s_A, u_A, w_A, mu, mu_prev, gammatest, beta_tmp, beta_t, Ac
End Function



'=============================================
'Elastic Net Regression
'=============================================
'Main Reference: "Regularization and variable selection via elastic net", Hui Zou (2004)
'https://web.stanford.edu/~hastie/Papers/B67.2%20(2005)%20301-320%20Zou%20&%20Hastie.pdf
'Input:     x() is a NxD predictor variables with zero mean and unit length (i.e. sum(x^2)=1)
'           y() is a length N vector of response variable
'           lambda2, L2 regularization term
'Output:    beta(), solution path of dimension (1 to n_dimension, 1 to number of steps)
'           A(),    order of which each dimension is added to the model
Function ENET(x() As Double, y() As Double, beta() As Double, A() As Long, lambda2 As Double, _
            Optional norm1 As Variant, Optional norm_rel As Long = 1) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim lambda1 As Double, tmp_x As Double, t_tgt As Double
Dim x2() As Double, y2() As Double, x_mean() As Double, x_scale() As Double, y_mean As Double
Dim mse() As Double, tmp_vec() As Double
Dim Gram() As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    'Find norm-1 of ridge regression solution
    If IsMissing(norm1) = False Then
        If norm_rel = 1 Then
            Dim u() As Double, d() As Double, v() As Double
            Call modMath.Matrix_SVD(x, u, d, v)
            For i = 1 To UBound(d)
                d(i) = d(i) / (d(i) ^ 2 + lambda2)
            Next i
            d = modMath.mDiag(d)
            tmp_vec = modMath.M_Dot(modMath.M_Dot(v, modMath.M_Dot(d, u, 0, 1)), y)
            t_tgt = beta_norm(tmp_vec) * norm1
'            t_tgt = t_tgt * (1 + lambda2)
            Erase u, d, v, tmp_vec
        Else
            t_tgt = norm1
        End If
    End If

    Gram = modMath.M_Dot(x, x, 1, 0)
    For i = 1 To n_dimension
        Gram(i, i) = Gram(i, i) + lambda2
    Next i

    If IsMissing(norm1) = False Then
        tmp_vec = LARS_EN(x, y, beta, A, Gram, lambda2, t_tgt)
    Else
        tmp_vec = LARS_EN(x, y, beta, A, Gram, lambda2)
    End If
    beta = modMath.M_scalar_dot(beta, (1 + lambda2))
    tmp_vec = modMath.M_scalar_dot(tmp_vec, (1 + lambda2))
    ENET = tmp_vec
    Erase tmp_vec, Gram
End Function



'Elastic net with LARS
'http://www2.imm.dtu.dk/pubdb/views/publication_details.php?id=3897
Private Function LARS_EN(x() As Double, y() As Double, beta() As Double, A() As Long, Gram() As Double, _
        lambda2 As Double, Optional norm1 As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, iterate As Long, iter_max As Long
Dim n_raw As Long, n_dimension As Long, m_dimension As Long
Dim addVar As Long, n_A As Long, n_Ac As Long, dropIdx As Long
Dim lassoCond As Long, stopCond As Long
Dim tmp_x As Double, tmp_y As Double, INFINITY As Double, tol As Double
Dim c_max As Double, gamma As Double, gamma2 As Double, t_prev As Double, t_now As Double
Dim correl() As Double, mu() As Double
Dim G_A() As Double, x_A() As Double
Dim beta_t() As Double, b_OLS() As Double
Dim d() As Double, cd() As Double
Dim Ac() As Long

    tol = 0.0000000001
    INFINITY = Exp(70)
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    If lambda2 < tol Then
        m_dimension = min2(n_raw, n_dimension)
    Else
        m_dimension = n_dimension
    End If
    iter_max = 8 * m_dimension
    ReDim beta(1 To n_dimension, 1 To 1)
    ReDim mu(1 To n_raw)
    ReDim A(0 To 0)
    ReDim Ac(0 To n_dimension)
    For i = 1 To n_dimension
        Ac(i) = i
    Next i
    
'    If lambda2 > 0 And IsMissing(norm1) = False Then
'        If norm1 > 0 Then norm1 = norm1 / (1 + lambda2)
'    End If
    
    lassoCond = 0
    stopCond = 0
    iterate = 1
    
    Do While UBound(A, 1) < m_dimension And stopCond = 0 And iterate < iter_max
    
        Call LARS_Correl(correl, c_max, j, x, y, mu)
        c_max = -INFINITY
        For i = 1 To UBound(Ac, 1)
            j = Ac(i)
            If Abs(correl(j)) > c_max Then
                c_max = Abs(correl(j))
                addVar = j
            End If
        Next i

        If lassoCond = 0 Then
            Call Append_1D(A, addVar)
            Call Eject_1D(Ac, addVar)
        Else
            lassoCond = 0
        End If
    
        n_A = UBound(A, 1)
        n_Ac = UBound(Ac, 1)
        ReDim x_A(1 To n_raw, 1 To n_A)
        ReDim G_A(1 To n_A, 1 To n_A)
        For i = 1 To n_A
            For k = 1 To n_raw
                x_A(k, i) = x(k, A(i))
            Next k
            For j = 1 To n_A
                G_A(i, j) = Gram(A(i), A(j))
            Next j
        Next i
        
        G_A = modMath.Matrix_Inverse_Cholesky(G_A)
        b_OLS = modMath.M_Dot(G_A, modMath.M_Dot(x_A, y, 1, 0))
        d = modMath.M_Dot(x_A, b_OLS)
        For i = 1 To n_raw
            d(i) = d(i) - mu(i)
        Next i
        
        gamma2 = INFINITY
        dropIdx = 0
        For i = 1 To n_A - 1
            tmp_x = beta(A(i), iterate) / (beta(A(i), iterate) - b_OLS(i))
            If tmp_x > 0 Then
                If tmp_x < gamma2 Then
                    gamma2 = tmp_x
                    dropIdx = i
                End If
            End If
        Next i
        
        If n_Ac = 0 Then
            gamma = 1
        Else
            gamma = INFINITY
            cd = modMath.M_Dot(x, d, 1, 0)
            For i = 1 To n_Ac
                j = Ac(i)
                tmp_x = (correl(j) - c_max) / (cd(j) - c_max)
                tmp_y = (correl(j) + c_max) / (cd(j) + c_max)
                If tmp_x <= 0 Then tmp_x = INFINITY
                If tmp_y <= 0 Then tmp_y = INFINITY
                tmp_x = min2(tmp_x, tmp_y)
                If tmp_x < gamma Then
                    gamma = tmp_x
                End If
            Next i
        End If
        
        If gamma2 < gamma Then
            lassoCond = 1
            gamma = gamma2
        End If
        
        ReDim Preserve beta(1 To n_dimension, 1 To iterate + 1)
        For i = 1 To n_A
            beta(A(i), iterate + 1) = (1 - gamma) * beta(A(i), iterate) + gamma * b_OLS(i)
        Next i
        For i = 1 To n_raw
            mu(i) = mu(i) + gamma * d(i)
        Next i
        iterate = iterate + 1
        
        If IsMissing(norm1) = False Then
            If norm1 > 0 Then
                t_now = 0
                For i = 1 To n_dimension
                    t_now = t_now + Abs(beta(i, iterate))
                Next i
                If t_now >= norm1 Or UBound(A, 1) = m_dimension Then
                    t_prev = 0
                    For i = 1 To n_dimension
                        t_prev = t_prev + Abs(beta(i, iterate - 1))
                    Next i
                    tmp_x = (norm1 - t_prev) / (t_now - t_prev)
                    ReDim beta_t(1 To n_dimension)
                    For i = 1 To n_dimension
                        beta_t(i) = (1 - tmp_x) * beta(i, iterate - 1) + tmp_x * beta(i, iterate)
                        beta(i, iterate) = beta_t(i)
                    Next i
                    stopCond = 1
                End If
            End If
        End If
        
        If lassoCond = 1 Then
            dropIdx = A(dropIdx)
            Call Append_1D(Ac, dropIdx)
            Call Eject_1D(A, dropIdx)
        End If
        
        If IsMissing(norm1) = False Then
            If norm1 < 0 Then
                If UBound(A, 1) >= (-norm1) Then
                    stopCond = 1
                Else
                    stopCond = 0
                End If
            End If
        End If
        
    Loop
    
    If iterate >= iter_max Then
        Debug.Print "LARS_EN: Error: Max iteration reached."
        Exit Function
    End If

    If IsMissing(norm1) = False Then
        LARS_EN = beta_t
    Else
        'LARS_EN = beta
        Call modMath.Filter_Array(beta, LARS_EN, , UBound(beta, 2))
    End If

End Function




'Perform LASSO with K-fold cross validation to find the optimal value of s,
'which is the norm of optimal beta relative to OLS solution
Function LASSO_CV(x() As Double, y() As Double, beta() As Double, A() As Long, s_optimal As Double, _
        Optional K_Fold As Long = 10, Optional n_iterate As Long = 1, Optional CV_Curve As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long, n_test As Long, iterate As Long
Dim n_s As Long, iCV As Long
Dim tmp_x As Double
Dim CV_err() As Double, s_test() As Double, s() As Double, beta_tmp() As Double
Dim x_test() As Double, x_train() As Double, y_test() As Double, y_train() As Double, y_est() As Double
Dim x_mean() As Double, x_scale() As Double, y_mean As Double
Dim iShuffle() As Long
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    n_s = 100                        'number of values of s to test, s in (0,1]
    n_test = n_raw \ K_Fold          'size of test set
    ReDim s_test(1 To n_s)
    For i = 1 To n_s
        s_test(i) = i * 0.01
    Next i
    
    ReDim CV_err(1 To n_s)
    For iterate = 1 To n_iterate
        
        DoEvents
        Application.StatusBar = "LASSO_CV: iterate: " & iterate & "/" & n_iterate
        
        'Shuffle original data set
        iShuffle = modMath.index_array(1, n_raw)
        Call modMath.Shuffle(iShuffle)
        
        For iCV = 1 To K_Fold
            'Split into training set and test set
            Call Split_Data(x, y, iShuffle, iCV, n_test, x_test, y_test, x_train, y_train, x_mean, x_scale, y_mean)
            
            beta_tmp = LARS(x_train, y_train, beta, A, True) 'Generate LASSO path
            
            Call calc_s(beta, s)    'norm of each beta relative to OLS solution

            For i = 1 To n_s
                Call Intrapolate_beta(s, beta, s_test(i), beta_tmp)    'Intrapolate beta
                Call Predict(x_test, beta_tmp, y_est, y_test, tmp_x)   'Find prediction error on test set
                CV_err(i) = CV_err(i) + tmp_x / (K_Fold * n_iterate)   'Accumulate error
            Next i
        Next iCV
    
    Next iterate
    
    tmp_x = Exp(70)
    For i = 1 To n_s
        If CV_err(i) < tmp_x Then
            tmp_x = CV_err(i)
            s_optimal = s_test(i)
        End If
    Next i
    
    beta_tmp = LARS(x, y, beta, A, True, s_optimal)
    LASSO_CV = beta_tmp
    If IsMissing(CV_Curve) = False Then CV_Curve = CV_err
    Erase x_train, x_test, y_train, y_test, y_est, x_mean, x_scale, CV_err
    Application.StatusBar = False
End Function


Function ENET_CV(x() As Double, y() As Double, beta() As Double, A() As Long, lambda_optimal As Double, s_optimal As Double, _
        Optional K_Fold As Long = 10, Optional n_iterate As Long = 1, Optional CV_Curve As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long, n_test As Long, iterate As Long
Dim n_s As Long, n_lambda As Long, iLambda As Long, iCV As Long
Dim tmp_x As Double, tmp_y As Double
Dim tmp_vec() As Double, lambda2 As Double
Dim CV_err() As Double, lamba_test() As Double, s_test() As Double
Dim x_test() As Double, x_train() As Double, y_test() As Double, y_train() As Double, y_est() As Double
Dim x_mean() As Double, x_scale() As Double, y_mean As Double
Dim beta_tmp() As Double, s() As Double
Dim iShuffle() As Long
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    n_lambda = 7
    n_s = 100
    n_test = n_raw \ K_Fold
    ReDim s_test(1 To n_s)
    ReDim lambda_test(1 To n_lambda)
    For i = 1 To n_s
        s_test(i) = i * 0.01
    Next i
    For i = 2 To n_lambda
        lambda_test(i) = 10 ^ (i - 4)
    Next i
    
    ReDim CV_err(1 To n_lambda, 1 To n_s)
    For iterate = 1 To n_iterate
    
        iShuffle = modMath.index_array(1, n_raw)
        Call modMath.Shuffle(iShuffle)
        
        For iLambda = 1 To n_lambda
        
            DoEvents
            Application.StatusBar = "ENET_CV: " & iLambda & "/" & n_lambda & ", iterate (" & iterate & "/" & n_iterate & ")"
            
            lambda2 = lambda_test(iLambda)
            For iCV = 1 To K_Fold
                'Split data into training and test set
                Call Split_Data(x, y, iShuffle, iCV, n_test, x_test, y_test, x_train, y_train, x_mean, x_scale, y_mean)
                
                beta_tmp = ENET(x_train, y_train, beta, A, lambda2) 'Run elastic net on training set
                
                Call calc_s(beta, s) 'norm of each solution rel. to full ridge solution
    
                For i = 1 To n_s
                    Call Intrapolate_beta(s, beta, s_test(i), beta_tmp)
                    Call Predict(x_test, beta_tmp, y_est, y_test, tmp_x)
                    CV_err(iLambda, i) = CV_err(iLambda, i) + tmp_x / (K_Fold * n_iterate)
                Next i
            Next iCV
        Next iLambda
    
    Next iterate
    
    
    tmp_x = Exp(70)
    For i = 1 To n_lambda
        For k = 1 To n_s
            If CV_err(i, k) < tmp_x Then
                tmp_x = CV_err(i, k)
                lambda_optimal = lambda_test(i)
                s_optimal = s_test(k)
            End If
        Next k
    Next i
    
    beta_tmp = ENET(x, y, beta, A, lambda_optimal, s_optimal)
    ENET_CV = beta_tmp
    If IsMissing(CV_Curve) = False Then CV_Curve = CV_err
    Erase x_train, x_test, y_train, y_test, x_mean, x_scale, CV_err
    Application.StatusBar = False
End Function


Private Sub LARS_Correl(correl() As Double, c_max As Double, j_max As Long, x() As Double, y() As Double, mu() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim n_raw As Long, n_dimension As Long
Dim tmp_x As Double
    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim correl(1 To n_dimension)
    c_max = 0
    j_max = 0
    For j = 1 To n_dimension
        tmp_x = 0
        For i = 1 To n_raw
            tmp_x = tmp_x + x(i, j) * (y(i) - mu(i))
        Next i
        correl(j) = tmp_x
        If Abs(tmp_x) > c_max Then
            c_max = Abs(tmp_x)
            j_max = j
        End If
    Next j
End Sub


Sub Predict(x() As Double, beta() As Double, y_est() As Double, Optional y_tgt As Variant, Optional sse As Variant)
Dim i As Long, j As Long, n As Long, n_dimension As Long
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    ReDim y_est(1 To n)
    For i = 1 To n
        For j = 1 To n_dimension
            y_est(i) = y_est(i) + x(i, j) * beta(j)
        Next j
    Next i
    If IsMissing(y_tgt) = False And IsMissing(sse) = False Then
        sse = 0
        For i = 1 To n
            sse = sse + (y_est(i) - y_tgt(i)) ^ 2
        Next i
        sse = sse / n
    End If
End Sub


'Append tgt to end of vector x()
Private Sub Append_1D(x As Variant, tgt As Variant)
Dim n As Long
    n = UBound(x) + 1
    ReDim Preserve x(LBound(x, 1) To n)
    x(n) = tgt
End Sub

'Remove tgt from vector x()
Private Sub Eject_1D(x() As Long, tgt As Long)
Dim i As Long, j As Long, n As Long
    n = UBound(x, 1)
    If x(n) = tgt Then
        ReDim Preserve x(LBound(x, 1) To n - 1)
        Exit Sub
    End If
    For i = 1 To n - 1
        If x(i) = tgt Then
            For j = i To n - 1
                x(j) = x(j + 1)
            Next j
            ReDim Preserve x(LBound(x, 1) To n - 1)
            Exit Sub
        End If
    Next i
    Debug.Print "Eject_1D: " & tgt & " not found in set."
End Sub

'Input: A(), a subset of integers in [1,p]
'Output: Subset of integers in [1,p] excluding A()
Private Function InActiveSet(A() As Long, p As Long) As Long()
Dim i As Long, j As Long, k As Long, notA As Long
Dim Ac() As Long
    ReDim Ac(1 To p)
    k = 0
    For i = 1 To p
        notA = 1
        For j = 1 To UBound(A)
            If A(j) = i Then
                notA = 0
                Exit For
            End If
        Next j
        If notA = 1 Then
            k = k + 1
            Ac(k) = i
        End If
    Next i
    If k > 0 Then
        ReDim Preserve Ac(1 To k)
    Else
        ReDim Ac(0 To 0)
    End If
    InActiveSet = Ac
    Erase Ac
End Function


Private Function min2(x As Variant, y As Variant) As Variant
    min2 = x
    If y < x Then min2 = y
End Function


'1-norm of a vector
Private Function beta_norm(beta() As Double) As Double
Dim i As Long
    beta_norm = 0
    For i = 1 To UBound(beta, 1)
        beta_norm = beta_norm + Abs(beta(i))
    Next i
End Function


'De-mean y()
Sub Normalize_y(y() As Double, y_mean As Double, Optional apply_transform As Boolean = False)
Dim i As Long, n As Long
    n = UBound(y, 1)
    If apply_transform = False Then
        y_mean = 0
        For i = 1 To n
            y_mean = y_mean + y(i)
        Next i
        y_mean = y_mean / n
    End If
    For i = 1 To n
        y(i) = y(i) - y_mean
    Next i
End Sub

'Normalize x() to zero mean and sum(x^2)=1 or sum(x^2)=n
Sub Normalize_x(x() As Double, x_mean() As Double, x_scale() As Double, Optional apply_transform As Boolean = False)
Dim i As Long, k As Long, n As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double
    n = UBound(x, 1)
    n_dimension = UBound(x, 2)
    If apply_transform = False Then
        ReDim x_mean(1 To n_dimension)
        ReDim x_scale(1 To n_dimension)
        For k = 1 To n_dimension
            tmp_x = 0
            tmp_y = 0
            For i = 1 To n
                tmp_x = tmp_x + x(i, k)
                tmp_y = tmp_y + x(i, k) ^ 2
            Next i
            tmp_x = tmp_x / n
            tmp_y = Sqr(tmp_y - n * (tmp_x ^ 2))    'sum(x^2)=1
            'tmp_y = Sqr(tmp_y / n - (tmp_x ^ 2))    'sum(x^2)=n
            If tmp_y = 0 Then Debug.Print "Warning: sd of dimension " & k & " is zero."
            x_mean(k) = tmp_x
            x_scale(k) = tmp_y
            If x_scale(k) = 0 Then
                Debug.Print "mLARS: Normalize_x: zero variance in dimension " & k
                Exit Sub
            End If
        Next k
    End If
    For i = 1 To n
        For k = 1 To n_dimension
            x(i, k) = (x(i, k) - x_mean(k)) / x_scale(k)
        Next k
    Next i
End Sub


'Split data into training and test set at the iCV step of cross-validation
'normalize training set then apply the same transformation on test set
Private Sub Split_Data(x() As Double, y() As Double, iShuffle() As Long, iCV As Long, n_test As Long, _
    x_test() As Double, y_test() As Double, x_train() As Double, y_train() As Double, _
    x_mean() As Double, x_scale() As Double, y_mean As Double)
Dim i As Long, n As Long
Dim iTest() As Long, iTrain() As Long
    n = UBound(x, 1)
    ReDim iTest(1 To n_test)
    For i = 1 To n_test
        iTest(i) = iShuffle(min2(iCV * n_test, n) - n_test + i)
    Next i
    iTrain = InActiveSet(iTest, n)
    Call modMath.Filter_Array(x, x_test, iTest)
    Call modMath.Filter_Array(y, y_test, iTest)
    Call modMath.Filter_Array(x, x_train, iTrain)
    Call modMath.Filter_Array(y, y_train, iTrain)
    Call Normalize_x(x_train, x_mean, x_scale)
    Call Normalize_y(y_train, y_mean)
    Call Normalize_x(x_test, x_mean, x_scale, True)
    Call Normalize_y(y_test, y_mean, True)
    Erase iTest, iTrain
End Sub

Private Sub calc_s(beta() As Double, s() As Double)
Dim k As Long
Dim tmp_x As Double
Dim tmp_vec() As Double
    'norm of final LARS step (OLS or Ridge solution)
    Call modMath.Filter_Array(beta, tmp_vec, , UBound(beta, 2))
    tmp_x = beta_norm(tmp_vec)
    'norm of each beta relative to final LARS step
    ReDim s(1 To UBound(beta, 2))
    For k = 1 To UBound(beta, 2)
        Call modMath.Filter_Array(beta, tmp_vec, , k)
        s(k) = beta_norm(tmp_vec) / tmp_x
    Next k
End Sub

'Intrapolate beta at s=s_tgt
Private Sub Intrapolate_beta(s() As Double, beta() As Double, s_tgt As Double, beta_out() As Double)
Dim j As Long, k As Long
Dim tmp_x As Double
    For k = 1 To UBound(s, 1) - 1
        If s(k) < s_tgt And s_tgt <= s(k + 1) Then
            tmp_x = (s_tgt - s(k)) / (s(k + 1) - s(k))
            ReDim beta_out(1 To UBound(beta, 1))
            For j = 1 To UBound(beta, 1)
                beta_out(j) = (1 - tmp_x) * beta(j, k) + tmp_x * beta(j, k + 1)
            Next j
            Exit For
        End If
    Next k
End Sub
