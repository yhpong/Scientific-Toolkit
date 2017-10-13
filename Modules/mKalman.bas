Attribute VB_Name = "mKalman"
Option Explicit

'=== Kalman Filter (notation from "Kalman filter" on Wikipedia)
'Input: z(1:T, 1:M), M-dimensional observation for time horizon T
'       n_dimension, number of dimension of state vector
'Output: x(1:T,1:n_dimension), estimate state vector for time horizon T
'        covar(1:T,1:n_dimension), variance of x for time horizon T
Function Filter(z() As Double, n_dimension As Long, Optional covar As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, t As Long
Dim n_raw As Long, n_measure As Long, n_control As Long
Dim dt As Double, tmp_x As Double, tmp_y As Double
Dim f() As Double, B() As Double, u() As Double, h() As Double
Dim q() As Double, r() As Double
Dim x() As Double, x_prev() As Double, x_post() As Double
Dim p_prev() As Double, p_post() As Double, zt() As Double
    
    n_raw = UBound(z, 1)
    n_measure = UBound(z, 2)
    ReDim x(1 To n_raw, 1 To n_dimension)
    If IsMissing(covar) = False Then ReDim covar(1 To n_raw, 1 To n_dimension)
    
    '=====================
    'Define Model
    '=====================
    '1: x_t = F_t x_{t-1} + B_t u_t + w_t
    '2: z_t = H_t x_t + v_t
    'x() is the true-state vector, z() is observed vector
    'F() is the state transition model
    'B() is control-input model applied to control vector u()
    'w_t() is process noise with covariance matrix q()
    'H() is the observation model
    'v_t() is observation noise with covariance matrix r()
    
    'In this example of a 2D projectile subjected to gravity, state is a 4-dimensional vector
    'of position and velocity. Observation is 2-dimensional postion.
    'The equations of motion are:
    'x  = x  + vx*dt
    'y  = y  + vy*dt + 0.5*g*dt^2
    'vx = vx
    'vy = vy + g*dt
    
    n_control = 1
    dt = 0.01
    ReDim f(1 To n_dimension, 1 To n_dimension)
    For i = 1 To n_dimension
        f(i, i) = 1
    Next i
    f(1, 3) = dt
    f(2, 4) = dt
    
    ReDim B(1 To n_dimension, 1 To n_control)
    ReDim u(1 To n_control)
    B(2, 1) = 0.5 * dt * dt
    B(4, 1) = dt
    u(1) = -9.8
    
    ReDim h(1 To n_measure, 1 To n_dimension)
    h(1, 1) = 1
    h(2, 2) = 1
    
    'Process noise covariance
    ReDim q(1 To n_dimension, 1 To n_dimension)
    q(1, 1) = 0.001
    q(2, 2) = 0.001
    q(3, 3) = 0.001
    q(4, 4) = 0.001
    q(1, 3) = 0
    q(2, 4) = 0
    For i = 1 To n_dimension - 1
        For j = i + 1 To n_dimension
            q(j, i) = q(i, j)
        Next j
    Next i
    
    'Observation noise covariance
    ReDim r(1 To n_measure, 1 To n_measure)
    r(1, 1) = 0.1
    r(2, 2) = 0.1
    For i = 1 To n_measure - 1
        For j = i + 1 To n_measure
            r(j, i) = r(i, j)
        Next j
    Next i
    '==============================================
    
    '=====================
    'Initialization
    '=====================
    'Initial state and covariance
    'in this example first 2 dimensions are just the observed x & y positions, third
    'and fourth dimensions are velocity estimated using the first 5 time steps, covariance
    'are assumed to be diagonal and same as model variance
    ReDim x_post(1 To n_dimension)
    ReDim p_post(1 To n_dimension, 1 To n_dimension)
    x_post(1) = z(1, 1)
    x_post(2) = z(1, 2)
    x_post(3) = (z(6, 1) - z(1, 1)) / (dt * 5)
    x_post(4) = (z(6, 2) - z(1, 2)) / (dt * 5)
    
    For i = 1 To n_dimension
        p_post(i, i) = q(i, i)
    Next i
    For i = 1 To n_dimension - 1
        For j = i + 1 To n_dimension
            p_post(i, j) = 0
            p_post(j, i) = p_post(i, j)
        Next j
    Next i
    '==============================================
    
    ReDim zt(1 To n_measure)
    For t = 1 To n_raw
    
        For i = 1 To n_measure
            zt(i) = z(t, i)
        Next i
        x_prev = x_post
        p_prev = p_post
        
        Call Filter_Step(x_post, p_post, zt, x_prev, p_prev, f, B, u, q, h, r)
        
        'Append results
        For i = 1 To n_dimension
            x(t, i) = x_post(i)
            If IsMissing(covar) = False Then covar(t, i) = p_post(i, i)
        Next i
        
    Next t
    
    Filter = x
    Erase x, x_post, x_prev, p_post, p_prev, zt
End Function


'=== Single step of Kalman filter update
'Output: x(1:D), state estimate
'        p(1:D,1:D), covariance estimate
'Input: z(1:M) , observed vector at current time step
'       x_prev(1:D), true state vector at previous time step
'       p_prev(1:D,1:D), covariance estimate from previous time step
'       f,B,u,q,h,r, model specifications, see comments in Sub Filter()
Sub Filter_Step(x() As Double, p() As Double, z() As Double, x_prev() As Double, p_prev() As Double, _
    f() As Double, B() As Double, u() As Double, q() As Double, h() As Double, r() As Double)
Dim i As Long
Dim y() As Double, s() As Double
Dim Kalman_Gain() As Double

    'Predict
    x = Add_Vec(MDot_Vec(f, x_prev), MDot_Vec(B, u))
    p = MAdd(MDot(MDot(f, p_prev), f, True), q)
    Call Symmetrize(p)
    
    'Update
    y = MDot_Vec(h, x)
    For i = 1 To UBound(z)
        y(i) = z(i) - y(i)
    Next i
    s = MAdd(MDot(MDot(h, p), h, True), r)
    s = modMath.Matrix_Inverse(s)
    Kalman_Gain = MDot(MDot(p, h, True), s)
    
    'Output posterior state and covariance
    x = Add_Vec(x, MDot_Vec(Kalman_Gain, y))
    p = MAdd(p, MDot(MDot(Kalman_Gain, h), p), True)
    Call Symmetrize(p)
    
    Erase y, s, Kalman_Gain
End Sub



'One-sided HP-Filter formulated as Kalman filter
Function Filter_HP(z() As Double, Optional lambda As Double = 1600, Optional covar As Variant) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, t As Long, n_raw As Long
Dim f() As Double, h() As Double, q() As Double, r() As Double
Dim x() As Double, x_prior() As Double, x_post() As Double, y() As Double, s() As Double
Dim P_prior() As Double, p_post() As Double, Kalman_Gain() As Double

    n_raw = UBound(z, 1)
    ReDim x(1 To n_raw)
    If IsMissing(covar) = False Then ReDim covar(1 To n_raw)
    
    '=== Define Model
    ReDim f(1 To 2, 1 To 2)
    f(1, 1) = 2
    f(1, 2) = -1
    f(2, 1) = 1
    
    ReDim h(1 To 1, 1 To 2)
    h(1, 1) = 1
    
    'Process noise covariance
    ReDim q(1 To 2, 1 To 2)
    q(1, 1) = 1# / lambda
    
    'Observation noise covariance
    ReDim r(1 To 1, 1 To 1)
    r(1, 1) = 1
    
    ReDim y(1 To 1)
    ReDim s(1 To 1, 1 To 1)
    '==============================================
    
    '=== Initialization
    'Initial state and covariance
    ReDim x_post(1 To 2)
    ReDim p_post(1 To 2, 1 To 2)
    'Backward extrapolation with the first two datapoints
    x_post(1) = 2 * z(1) - z(2)
    x_post(2) = 3 * z(1) - 2 * z(2)
    p_post(1, 1) = 100000
    p_post(2, 2) = 100000
    '==============================================
    
    For t = 1 To n_raw
        'Predict
        x_prior = MDot_Vec(f, x_post)
        'P_prior = MAdd(MDot(MDot(F, P_post), F, True), Q)
        P_prior = MDot(MDot(f, p_post), f, True)
        P_prior(1, 1) = P_prior(1, 1) + 1# / lambda
        
        'Update
        y(1) = z(t) - x_prior(1)
        s(1, 1) = P_prior(1, 1) + r(1, 1)
        s(1, 1) = 1# / s(1, 1)
        Kalman_Gain = MDot(MDot(P_prior, h, True), s)
        x_post = Add_Vec(x_prior, MDot_Vec(Kalman_Gain, y))
        p_post = MAdd(P_prior, MDot(MDot(Kalman_Gain, h), P_prior), True)
        Call Symmetrize(p_post)
        
        'Append results
        x(t) = x_post(1)
        If IsMissing(covar) = False Then covar(t) = Sqr(p_post(1, 1))
    Next t
    
    Filter_HP = x
    Erase x, x_post, x_prior, p_post, P_prior, s, y, Kalman_Gain, f, q, r
End Function


Private Sub Symmetrize(A() As Double)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
    n = UBound(A, 1)
    For i = 1 To n - 1
        For j = i + 1 To n
            A(j, i) = A(i, j)
        Next j
    Next i
End Sub

Private Function MDot(A() As Double, B() As Double, Optional B_transpose As Boolean = False) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long, t As Long
Dim c() As Double
    m = UBound(A, 1)
    t = UBound(A, 2)
    If B_transpose = False Then
        n = UBound(B, 2)
        ReDim c(1 To m, 1 To n)
        For i = 1 To m
            For j = 1 To n
                For k = 1 To t
                    c(i, j) = c(i, j) + A(i, k) * B(k, j)
                Next k
            Next j
        Next i
    Else
        n = UBound(B, 1)
        ReDim c(1 To m, 1 To n)
        For i = 1 To m
            For j = 1 To n
                For k = 1 To t
                    c(i, j) = c(i, j) + A(i, k) * B(j, k)
                Next k
            Next j
        Next i
    End If
    MDot = c
    Erase c
End Function


Private Function MDot_Vec(A() As Double, B() As Double) As Double()
Dim i As Long, j As Long, k As Long, m As Long, n As Long
Dim c() As Double
    m = UBound(A, 1)
    n = UBound(A, 2)
    ReDim c(1 To m)
    For i = 1 To m
        For j = 1 To n
            c(i) = c(i) + A(i, j) * B(j)
        Next j
    Next i
    MDot_Vec = c
    Erase c
End Function


Private Function MAdd(A() As Double, B() As Double, Optional minusB As Boolean = False) As Double()
Dim i As Long, j As Long, n As Long, m As Long
Dim c() As Double
    n = UBound(A, 1)
    m = UBound(A, 2)
    ReDim c(1 To n, 1 To m)
    If minusB = False Then
        For i = 1 To n
            For j = 1 To m
                c(i, j) = A(i, j) + B(i, j)
            Next j
        Next i
    Else
        For i = 1 To n
            For j = 1 To m
                c(i, j) = A(i, j) - B(i, j)
            Next j
        Next i
    End If
    MAdd = c
    Erase c
End Function


Private Function Add_Vec(A() As Double, B() As Double) As Double()
Dim i As Long, n As Long
Dim c() As Double
    n = UBound(A, 1)
    ReDim c(1 To n)
    For i = 1 To n
        c(i) = A(i) + B(i)
    Next i
    Add_Vec = c
    Erase c
End Function
