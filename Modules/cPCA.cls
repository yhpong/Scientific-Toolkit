VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===========================================================================
'Principal Component Analysis
'===========================================================================
'Input data is assumed to be N observations of D-dimensional vector in the form
'of x(1 to N, 1 to D). It is also assumed that data is normalized beforehand.
'===========================================================================

Private pVec() As Double    'unit vector of each principal component
Private pVal() As Double    'eigen value of each principal component
Private px_PCA() As Double  'Projection of x() onto each PC

Public Property Get Vec(Optional n_vec As Variant) As Double()
Dim y() As Double
    If IsMissing(n_vec) Then
        Vec = pVec
        Exit Property
    End If
    y = pVec
    ReDim Preserve y(1 To UBound(y, 1), 1 To n_vec)
    Vec = y
    Erase y
End Property

Public Property Get Val(Optional n_vec As Variant) As Double()
Dim y() As Double
    If IsMissing(n_vec) Then
        Val = pVal
    Else
        y = pVal
        ReDim Preserve y(1 To n_vec)
        Val = y
        Erase y
    End If
End Property

Public Property Get x_PCA(Optional n_vec As Variant) As Double()
Dim i As Long, j As Long, n_raw As Long
Dim y() As Double
    If IsMissing(n_vec) Then
        x_PCA = px_PCA
        Exit Property
    End If
    y = px_PCA
    ReDim Preserve y(1 To UBound(y, 1), 1 To n_vec)
    x_PCA = y
    Erase y
End Property

Sub Reset()
    Erase pVec, pVal, px_PCA
End Sub


Sub BiPlot_Print(vRng As Range, Optional PC1 As Long = 1, Optional PC2 As Long = 2, _
        Optional magnify As Double = 1)
Dim i As Long, j As Long, k As Long, n_dimension As Long
Dim tmp_x As Double, tmp_y As Double
    n_dimension = UBound(pVec, 1)
    With vRng
        For j = 1 To n_dimension
            tmp_x = pVec(j, PC1) * Sqr(pVal(PC1)) * magnify
            tmp_y = pVec(j, PC2) * Sqr(pVal(PC2)) * magnify
            .Offset(j - 1, 0).Value = tmp_x
            .Offset(j - 1, 1).Value = tmp_y
            .Offset((j - 1) * 3, 2).Value = 0
            .Offset((j - 1) * 3, 3).Value = 0
            .Offset((j - 1) * 3 + 1, 2).Value = tmp_x
            .Offset((j - 1) * 3 + 1, 3).Value = tmp_y
        Next j
    End With
End Sub


'Input: x(1 to n_raw,1 to dimension)
'Input: first_n, if specified, use power iteration fo find the first
'       few components only, suitable for large matrix
'Output: .Vec(i,k) is the i-th element of the k-th PC
'Output: .Val(k) is the eigenvalue of the k-th PC
'Output: .x_PCA(1 to n_raw, 1 to n_PC) gives the projection of x() onto each PC
Sub PCA(x() As Double, Optional first_n As Long = 0, Optional use_SVD As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, m As Long
Dim n_raw As Long, n_dimension As Long
Dim x_covar() As Double
Dim tmp_x As Double

    n_raw = UBound(x, 1)
    n_dimension = UBound(x, 2)
    
    If first_n = 0 Then
        If use_SVD = True Then
            Call modMath.Matrix_SVD(x, px_PCA, pVal, pVec)
            n = UBound(pVec, 2)
            For i = 1 To n_raw
                For k = 1 To n
                    px_PCA(i, k) = px_PCA(i, k) * pVal(k)
                Next k
            Next i
            For k = 1 To n
                pVal(k) = (pVal(k) ^ 2) / (n_raw - 1)
            Next k
        Else
            x_covar = modMath.Covariance_Matrix(x)
            Call modMath.Eigen_Jacobi(x_covar, pVec, pVal)
            ReDim px_PCA(1 To n_raw, 1 To n_dimension)
            For k = 1 To n_dimension
                For n = 1 To n_dimension
                    tmp_x = pVec(n, k)
                    For i = 1 To n_raw
                        px_PCA(i, k) = px_PCA(i, k) + x(i, n) * tmp_x
                    Next i
                Next n
            Next k
        End If
    Else
        x_covar = modMath.Covariance_Matrix(x)
        Call modMath.Eigen_Power(x_covar, pVec, pVal, first_n)
        ReDim px_PCA(1 To n_raw, 1 To first_n)
        For k = 1 To first_n
            For n = 1 To n_dimension
                tmp_x = pVec(n, k)
                For i = 1 To n_raw
                    px_PCA(i, k) = px_PCA(i, k) + x(i, n) * tmp_x
                Next i
            Next n
        Next k
    End If
End Sub
