VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pL_x As Long, pL_y As Long, pdimension As Long
Private pgrid_xy() As Double, pgrid_lines() As Double
Private pquant_err() As Double
Private pnode_w() As Double
Private pnode_wnorm() As Double
Private pnode_UMatrix() As Double
Private pnode_ED() As Double
Private pnode_size() As Long
Private px_BMU() As Long
Private px_D2BMU() As Double


Sub Reset()
    Erase pquant_err, pnode_w, pnode_wnorm, pnode_UMatrix, pnode_ED, pnode_size, px_BMU, px_D2BMU
End Sub

Public Property Get quant_err() As Double()
    quant_err = pquant_err
End Property

Public Property Get x_BMU() As Long()
    x_BMU = px_BMU
End Property

Public Property Get x_D2BMU() As Double()
    x_D2BMU = px_D2BMU
End Property

Public Property Get wgts() As Double()
    wgts = pnode_w
End Property

Public Property Get wgt(d As Long) As Double()
Dim i As Long, j As Long
Dim y() As Double
ReDim y(1 To pL_x, 1 To pL_y)
For i = 1 To pL_x
    For j = 1 To pL_y
        y(i, j) = pnode_w(i, j, d)
    Next j
Next i
wgt = y
End Property

Public Property Get wgts_norm() As Double()
    wgts_norm = pnode_wnorm
End Property

Public Property Get UMatrix() As Double()
    UMatrix = pnode_UMatrix
End Property

Sub Init(L_x As Long, L_y As Long, dimension As Long)
Dim i As Long, j As Long, d As Long
pL_x = L_x
pL_y = L_y
pdimension = dimension
'Grid lines for visualization in Excel
Call Me.HexGen(pL_x, pL_y, 1, pgrid_xy, pgrid_lines)
ReDim pnode_w(1 To pL_x, 1 To pL_y, 1 To pdimension) 'feature vector of each node
ReDim pnode_wnorm(1 To pL_x, 1 To pL_y)         'norm of feature vector of each node
ReDim pnode_size(1 To pL_x, 1 To pL_y)          'size of each node
ReDim pnode_UMatrix(1 To pL_x, 1 To pL_y)       'avg distance of a node to its neighbours
ReDim pnode_ED(1 To pL_x, 1 To pL_y, 1 To 6)    'distance of a node to each of its 6 neighbours
'Radomly initialize node weights
For i = 1 To pL_x
    For j = 1 To pL_y
        Randomize
        For d = 1 To dimension
            pnode_w(i, j, d) = Rnd() / 10 '(-1 + 2 * Rnd())/10
        Next d
    Next j
Next i
End Sub


'=== Output network results
'1st column: index
'2nd column: row index of node
'3rd column: column index of node
'4th column: norm of node's vector
'5th column: average distance to neighbor nodes
'6th column: node vector
Sub Print_Network_results(vRng As Range)
Dim i As Long, j As Long, d As Long, n As Long
With vRng
    n = 0
    For i = 1 To pL_x
    For j = 1 To pL_y
        .Offset(n, 0).Value = n + 1
        .Offset(n, 1).Value = i
        .Offset(n, 2).Value = j
        .Offset(n, 3).Value = pnode_wnorm(i, j)
        .Offset(n, 4).Value = pnode_UMatrix(i, j)
        For d = 1 To pdimension
            vRng.Offset(n, 4 + d).Value = pnode_w(i, j, d)
        Next d
        n = n + 1
    Next j
    Next i
End With
End Sub

'=== Print grid lines for charts
Sub Print_Grid(vRng As Range)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
k = 0
m = 0
For i = 1 To pL_x
    For j = 1 To pL_y
        m = m + 1
        vRng.Offset(m - 1, 0).Value = i
        vRng.Offset(m - 1, 1).Value = j
        vRng.Offset(m - 1, 2).Value = pgrid_xy(i, j, 1)
        vRng.Offset(m - 1, 3).Value = pgrid_xy(i, j, 2)
        For n = 1 To 7
            vRng.Offset(k + n - 1, 4).Value = pgrid_lines(i, j, n, 1)
            vRng.Offset(k + n - 1, 5).Value = pgrid_lines(i, j, n, 2)
        Next n
        k = k + 8
    Next j
Next i
End Sub

'== Format a chart of U-Matrix in gray scale
Sub Format_UMatrix(cht As Chart, Optional write_labels As Boolean = False, Optional node_label As Variant)
Dim i As Long, j As Long, n As Long, vR As Long
Dim temp_max As Double, temp_min As Double, temp As Double
temp_max = -999999999
temp_min = 999999999
For i = 1 To pL_x
    For j = 1 To pL_y
        temp = pnode_UMatrix(i, j)
        If temp > temp_max Then temp_max = temp
        If temp < temp_min Then temp_min = temp
    Next j
Next i
With cht.SeriesCollection(1)
    If .HasDataLabels = True Then .DataLabels.Delete
    If write_labels = True Then .ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False
    n = 0
    For i = 1 To pL_x
        For j = 1 To pL_y
        
            n = n + 1
            temp = (pnode_UMatrix(i, j) - temp_min) / (temp_max - temp_min)
            Call Gray_Scale(temp, vR)
            .Points(n).Format.Fill.ForeColor.RGB = RGB(vR, vR, vR)
            
            If write_labels = True Then
                With .Points(n).DataLabel
                    If pnode_size(i, j) > 0 Then
                        .Position = xlLabelPositionCenter
                        .Text = node_label(i, j)
                        .Format.TextFrame2.TextRange.Font.Size = 8
                        If vR < 75 Then .Font.Color = RGB(255, 255, 255)
                    Else
                        .Delete
                    End If
                End With
            End If
            
        Next j
    Next i
End With
End Sub


Sub Format_Grid(cht As Chart, node_attr() As Double, Optional write_labels As Boolean = False, Optional node_label As Variant)
Dim i As Long, j As Long, n As Long, vR As Long, vG As Long, vB As Long
Dim temp_max As Double, temp_min As Double, temp As Double
temp_max = -999999999
temp_min = 999999999
For i = 1 To pL_x
    For j = 1 To pL_y
        temp = node_attr(i, j)
        If temp > temp_max Then temp_max = temp
        If temp < temp_min Then temp_min = temp
    Next j
Next i
With cht.SeriesCollection(2)
    If .HasDataLabels = True Then .DataLabels.Delete
    If write_labels = True Then .ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False
    n = 0
    For i = 1 To pL_x
        For j = 1 To pL_y
        
            n = n + 1
            temp = (node_attr(i, j) - temp_min) / (temp_max - temp_min)
            Call Color_Scale(temp, vR, vG, vB)
            .Points(n).Format.Fill.ForeColor.RGB = RGB(vR, vG, vB)
            
            If write_labels = True Then
                With .Points(n).DataLabel
                    If pnode_size(i, j) > 0 Then
                        .Position = xlLabelPositionCenter
                        .Text = node_label(i, j)
                        .Font.Size = 8
                        If vB > 160 Then .Font.Color = RGB(255, 255, 255)
                    Else
                        .Delete
                    End If
                End With
            End If
            
        Next j
    Next i
End With
End Sub


'Input: x is a real number between 0 and 1
'Output: vR,vG,vB are integers from 0 to 255
Private Sub Color_Scale(x As Double, vR As Long, vG As Long, vB As Long)
If x <= 0.5 Then
    vR = 255
    vG = Int(510 * x)
    vB = 0
Else
    vR = Int(-510 * (x - 1))
    vG = vR
    vB = Int(510 * x - 255)
End If
End Sub

'Input: x is a real number between 0 and 1
'Output: vR are integers from 0 to 255
Private Sub Gray_Scale(x As Double, vR As Long)
vR = Int(-255 * (1 - Exp(-(x - 1) * 5)) / (1 + Exp(-(x - 1) * 5)))
End Sub

'=== Print weight vectors of each node
Sub Print_Wgts(vRng As Range)
Dim i As Long, j As Long, k As Long, d As Long
k = 0
For i = 1 To pL_x
    For j = 1 To pL_y
        For d = 1 To pdimension
            vRng.Offset(k, d - 1) = pnode_w(i, j, d)
        Next d
        k = k + 1
    Next j
Next i
End Sub


'=== Print grids of all dimensions
Sub Print_All_Dimensions(mysht As Worksheet, Optional cht_width As Long = 280, Optional cht_height As Long = 280, Optional markersize As Long = 17, Optional write_labels As Boolean = False, Optional node_labels As Variant, Optional factor_names As Variant)
Dim i As Long, j As Long, d As Long, m As Long, n As Long, n_pt As Long
Dim source_chart As ChartObject, new_chart As Object
Dim myshp As Shape, mychart As Chart
Dim attr_name() As String
    With mysht
        .Range("P2:U1000000").Clear
        Call Me.Print_Grid(.Range("P1"))
        n_pt = pL_x * pL_y
        n = .Range("T100000").End(xlUp).Row
    End With
    
    ReDim attr_name(1 To pdimension)
    If IsMissing(factor_names) = True Then
        For i = 1 To pdimension
            attr_name(i) = "ATTR_" & Format(i, "000")
        Next i
    Else
        For i = 1 To pdimension
            attr_name(i) = factor_names(i)
        Next i
    End If
    
    m = mysht.ChartObjects.count
    If m = 0 Then
        With mysht
            Set myshp = .Shapes.AddChart2(240, xlXYScatterLinesNoMarkers, .Cells(3, 2).Left, .Cells(3, 2).Top, cht_width, cht_height)
        End With
        Set mychart = myshp.Chart
        With mychart
            .SetSourceData Source:=Range(mysht.Name & "!$T$1:$U$" & n)
            .SeriesCollection(1).Format.Line.Weight = 1
            .SeriesCollection.NewSeries
            .SeriesCollection(2).XValues = "=" & mysht.Name & "!$R$1:$R$" & n_pt
            .SeriesCollection(2).Values = "=" & mysht.Name & "!$S$1:$S$" & n_pt
            .SeriesCollection(2).MarkerStyle = 8
            .SeriesCollection(2).markersize = markersize
            .SeriesCollection(2).Format.Line.Visible = msoFalse
            If .Axes(xlCategory).HasMajorGridlines Then .Axes(xlCategory).MajorGridlines.Delete
            If .Axes(xlValue).HasMajorGridlines Then .Axes(xlValue).MajorGridlines.Delete
            If .HasTitle Then .ChartTitle.Delete
        End With
        myshp.Name = "ATTR_001"
    End If
    
    With mysht
        m = .ChartObjects.count
        If m > pdimension Then
            For d = pdimension + 1 To m
                .ChartObjects("ATTR_" & Format(d, "00")).Delete
                i = 2 + Int((d - 1) * 1# / 2) * 23
                j = 2 + (1 - (d Mod 2)) * 6
                .Cells(i, j).ClearContents
            Next d
        ElseIf m <= pdimension Then
            For i = 1 To m
                Set mychart = .ChartObjects("ATTR_" & Format(i, "000")).Chart
                mychart.SeriesCollection(1).XValues = "=" & mysht.Name & "!$T$1:$T$" & n
                mychart.SeriesCollection(1).Values = "=" & mysht.Name & "!$U$1:$U$" & n
                mychart.SeriesCollection(2).XValues = "=" & mysht.Name & "!$R$1:$R$" & n_pt
                mychart.SeriesCollection(2).Values = "=" & mysht.Name & "!$S$1:$S$" & n_pt
            Next i
        
            Set source_chart = .ChartObjects("ATTR_001")
            For d = m + 1 To pdimension
                i = 2 + Int((d - 1) * 1# / 2) * 23
                j = 2 + (1 - (d Mod 2)) * 6
                Set new_chart = source_chart.Duplicate
                new_chart.Name = "ATTR_" & Format(d, "000")
                new_chart.Left = .Cells(i + 1, j).Left
                new_chart.Top = .Cells(i + 1, j).Top
            Next d
            Set source_chart = Nothing
            Set new_chart = Nothing
        End If
    
        For d = 1 To pdimension
            i = 2 + Int((d - 1) * 1# / 2) * 23
            j = 2 + (1 - (d Mod 2)) * 6
            .Cells(i, j).Value = attr_name(d)
            If write_labels = True Then
                Call Me.Format_Grid(.ChartObjects("ATTR_" & Format(d, "000")).Chart, Me.wgt(d), True, node_labels)
            Else
                Call Me.Format_Grid(.ChartObjects("ATTR_" & Format(d, "000")).Chart, Me.wgt(d), False)
            End If
        Next d
    End With
End Sub


'==== Train the SOM
'Input: x(1 to N, 1 to D) is a matrix with N members, and D-dimension feature vector.
Sub SOM_Hex_Train(x() As Double, _
        Optional iterate_max As Long = 5000, Optional learn_rate As Double = 0.1, _
        Optional batch_training As Boolean = False, Optional use_PCA As Boolean = True, _
        Optional random_sampling As Boolean = True)
        
Dim i As Long, j As Long, m As Long, n As Long, k As Long, d As Long
Dim n_raw As Long, dimension As Long
Dim iterate As Long
Dim Samples() As Long
Dim node_dist() As Double
Dim BMU_ED As Double, BMU_index As Long
Dim neighbour_r As Double, map_radius As Double, Lt As Double
Dim neighbour_const As Double
Dim temp As Double, temp_max As Double, temp_min As Double, tmp_x As Double, qerr As Double
Dim temp_i As Long, temp_j As Long
Dim eigen_vec() As Double, eigen_val() As Double
Dim xk() As Double

n_raw = UBound(x, 1)
dimension = UBound(x, 2)
If dimension <> pdimension Then
    msgbox "Training data has different dimensions from grid"
    Exit Sub
End If

ReDim pquant_err(1 To 2, 1 To iterate_max)
ReDim node_dist(1 To pL_x, 1 To pL_y, 1 To pL_x, 1 To pL_y)

map_radius = 0.5 * Sqr(((pL_x - 1) * 1.5) ^ 2 + ((pL_y - 0.5) * Sqr(3)) ^ 2)
neighbour_const = Log(map_radius)

'==== Initialize node weights to PC1 and PC2
If use_PCA = True Then
    Application.StatusBar = "Initializing grid with PCA..."
    Dim PC1 As New cPCA
    Call PC1.PCA(x)
    eigen_vec = PC1.Vec
    eigen_val = PC1.Val
    n = 0
    For i = 1 To pL_x
        For j = 1 To pL_y
            For d = 1 To dimension
                pnode_w(i, j, d) = eigen_vec(d, 1) * 0.8 * Sqr(eigen_val(1)) * (i * 1# / pL_x) + _
                    eigen_vec(d, 2) * 0.8 * Sqr(eigen_val(2)) * (j * 1# / pL_y)
            Next d
        Next j
    Next i
    Set PC1 = Nothing
    Erase eigen_vec, eigen_val
    Application.StatusBar = False
End If
'========================================

'=== Save node to node distance
'node_dist(i,j,m,n) is the square distance from node(i,j) to node(m,n)
For i = 1 To pL_x
Application.StatusBar = "Computing node to node distances..."
For j = 1 To pL_y
    For m = 1 To pL_x
    For n = 1 To pL_y
        node_dist(i, j, m, n) = (pgrid_xy(i, j, 1) - pgrid_xy(m, n, 1)) ^ 2 + (pgrid_xy(i, j, 2) - pgrid_xy(m, n, 2)) ^ 2
    Next n
    Next m
Next j
Application.StatusBar = False
Next i
'========================================

temp_j = 0 'index to print q-error of each iteration

'=== Create learning sequece base on sampling method chose
If batch_training = False Then
    Application.StatusBar = "Caching training sequences..."
    ReDim Samples(1 To iterate_max)
    If random_sampling = True Then
        'Pick a random input from raw data
        For iterate = 1 To iterate_max
            Randomize
            Samples(iterate) = Int(Rnd() * n_raw) + 1
        Next iterate
    Else
        k = 0
        'Feed input in same order
        For iterate = 1 To iterate_max
            If k + 1 <= n_raw Then
                k = k + 1
            Else
                k = 1
            End If
            Samples(iterate) = k
        Next iterate
    End If
    Application.StatusBar = False
End If
'========================================

'***********************************************************
'Start SOM iteration
'***********************************************************
For iterate = 1 To iterate_max

    DoEvents
    If iterate Mod 20 = 0 Then Application.StatusBar = "Training SOM..." & iterate & "/" & iterate_max
    
    'Shrink neighbourhood and Decay the learning rate
    neighbour_r = (map_radius * Exp(-(iterate - 1) * neighbour_const / iterate_max)) ^ 2
    Lt = learn_rate * Exp(-(iterate - 1) * 1# / iterate_max)
    
    If batch_training = False Then
        
        '=== Stochastic Descent
        k = Samples(iterate)
        Call get_vector(x, k, 1, xk)
        
        'Find BMU of current data
        Call Me.Find_BMU(xk, m, n, BMU_ED)
        
        'Adjust the weights of each node
        For i = 1 To pL_x
            For j = 1 To pL_y
                temp = Lt * Exp(-0.5 * node_dist(i, j, m, n) / neighbour_r)
                For d = 1 To dimension
                    pnode_w(i, j, d) = pnode_w(i, j, d) + (xk(d) - pnode_w(i, j, d)) * temp
                Next d
            Next j
        Next i
        '========================================
        
    ElseIf batch_training = True Then
        
        '=== Batch Descent
        'Find BMU of every raw data point
        Call Me.Find_BMU_Batch(x, px_BMU, px_D2BMU)
        
        'Adjust Weights of each node
        ReDim xk(1 To pL_x, 1 To pL_y, 1 To n_raw)
        n = 0
        For i = 1 To pL_x
        For j = 1 To pL_y
            temp = 0
            For m = 1 To n_raw
                xk(i, j, m) = Exp(-0.5 * node_dist(i, j, px_BMU(m, 1), px_BMU(m, 2)) / neighbour_r)
                temp = temp + xk(i, j, m)
            Next m
            
            For d = 1 To dimension
                tmp_x = 0
                For m = 1 To n_raw
                    tmp_x = tmp_x + x(m, d) * xk(i, j, m)
                Next m
                pnode_w(i, j, d) = tmp_x / temp
            Next d
        Next j
        Next i
        '========================================
    
    End If
        
    '=== Calculate quantization error at fixed intervals
    If (iterate - 1) Mod 10 = 0 Then
        temp_j = temp_j + 1
        qerr = 0
        For k = 1 To n_raw
            Call get_vector(x, k, 1, xk)
            Call Me.Find_BMU(xk, i, j, BMU_ED)
            qerr = qerr + BMU_ED
        Next k
        pquant_err(1, temp_j) = iterate
        pquant_err(2, temp_j) = qerr / (n_raw * dimension)
    End If
    '========================================

Next iterate

ReDim Preserve pquant_err(1 To 2, 1 To temp_j)

'***********************************************************
'Finished SOM iteration, process additional information
'***********************************************************

'=== Find the BMU of each training data
Call Me.Find_BMU_Batch(x, px_BMU, px_D2BMU)
'========================================

'==== Number of data assigned to each node
For m = 1 To n_raw
    i = px_BMU(m, 1)
    j = px_BMU(m, 2)
    pnode_size(i, j) = pnode_size(i, j) + 1
Next m
'========================================

'=== Norm of weight vectors
For i = 1 To pL_x
    For j = 1 To pL_y
        tmp_x = 0
        For d = 1 To dimension
            tmp_x = tmp_x + (pnode_w(i, j, d)) ^ 2
        Next d
        pnode_wnorm(i, j) = Sqr(tmp_x / dimension)
    Next j
Next i
'========================================

'=== U-Matrix Calculation
For i = 1 To pL_x
For j = 1 To pL_y
    '=== U-Matrix Calculation
    tmp_x = 0
    k = 0
    'loop through the 6 nearest neighbours
    For n = -1 To 1 Step 2
        If (j + n) >= 1 And (j + n) <= pL_y Then
        k = k + 1
        temp = 0
        For d = 1 To dimension
            temp = temp + (pnode_w(i, j, d) - pnode_w(i, j + n, d)) ^ 2
        Next d
        tmp_x = tmp_x + Sqr(temp)
        If n = 1 Then pnode_ED(i, j, 1) = Sqr(temp)
        If n = -1 Then pnode_ED(i, j, 2) = Sqr(temp)
        End If
    Next n
    
    For m = -1 To 1 Step 2
    For n = 0 To 1
        temp_i = i Mod 2
        If (i + m) >= 1 And (i + m) <= pL_x And _
            (j + n - temp_i) >= 1 And (j + 1 - temp_i) <= pL_y Then
        k = k + 1
        temp = 0
        For d = 1 To dimension
            temp = temp + (pnode_w(i, j, d) - pnode_w(i + m, j + n - temp_i, d)) ^ 2
        Next d
        tmp_x = tmp_x + Sqr(temp)
        If m = -1 And n = 0 Then pnode_ED(i, j, 3) = Sqr(temp)
        If m = -1 And n = 1 Then pnode_ED(i, j, 4) = Sqr(temp)
        If m = 1 And n = 0 Then pnode_ED(i, j, 5) = Sqr(temp)
        If m = 1 And n = 1 Then pnode_ED(i, j, 6) = Sqr(temp)
        End If
    Next n
    Next m
    pnode_UMatrix(i, j) = tmp_x / k
    '========================================
    
Next j
Next i
'========================================

Application.StatusBar = False

End Sub


Sub Find_BMU(x() As Double, m As Long, n As Long, BMU_ED As Double)
Dim i As Long, j As Long, d As Long
Dim temp As Double
BMU_ED = 1073741824 '2^30
For i = 1 To pL_x
    For j = 1 To pL_y
        temp = 0
        For d = 1 To pdimension
            temp = temp + (x(d) - pnode_w(i, j, d)) ^ 2
        Next d
        If temp < BMU_ED Then
            BMU_ED = temp
            m = i
            n = j
        End If
    Next j
Next i
End Sub


Sub Find_BMU_Batch(x() As Double, x_BMU() As Long, x_D2BMU() As Double)
Dim i As Long, j As Long, k As Long, d As Long, m As Long, n As Long
Dim tmp_min As Double, temp As Double
ReDim x_BMU(1 To UBound(x, 1), 1 To 2)
ReDim x_D2BMU(1 To UBound(x, 1))
For k = 1 To UBound(x, 1)
    tmp_min = 1073741824 '2^30
    For i = 1 To pL_x
        For j = 1 To pL_y
            temp = 0
            For d = 1 To pdimension
                temp = temp + (x(k, d) - pnode_w(i, j, d)) ^ 2
            Next d
            If temp < tmp_min Then
                tmp_min = temp
                m = i
                n = j
            End If
        Next j
    Next i
    x_D2BMU(k) = tmp_min
    x_BMU(k, 1) = m
    x_BMU(k, 2) = n
Next k
End Sub


'=== Generate labels to show members of each node, starting with the closest member
Sub Get_Node_Labels(x_name As Variant, node_label() As String)
'Sort the data by how close they are to the respective BMU
Dim i As Long, j As Long, m As Long, n As Long, n_raw As Long
Dim x_ED() As Double
Dim sort_index() As Long
x_ED = px_D2BMU
n_raw = UBound(x_ED, 1)
Call modMath.Sort_Quick_A(x_ED, 1, n_raw, sort_index, 1)
ReDim node_label(1 To pL_x, 1 To pL_y)
For n = 1 To n_raw
    m = sort_index(n)
    i = px_BMU(m, 1)
    j = px_BMU(m, 2)
    If node_label(i, j) = "" Then
        node_label(i, j) = "" & x_name(m)
    Else
        node_label(i, j) = node_label(i, j) & ", " & x_name(m)
    End If
Next n
End Sub


Function Grid2Idx(i As Long, j As Long) As Long
    Grid2Idx = (i - 1) * pL_y + j
End Function

Sub Idx2Grid(idx As Long, i As Long, j As Long)
    i = Int(idx * 1# / pL_y)
    j = idx Mod pL_y
End Sub


'Input: x(1 to M, 1 to N), 2D Matrix
'Output: y(), vector from the k-th row/column of x()
Sub get_vector(x As Variant, k As Long, idim As Long, y As Variant)
Dim i As Long, n As Long
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
End Sub


Sub HexGen(lattice_x As Long, lattice_y As Long, L As Double, xy() As Double, xy_Edges() As Double)
Dim i As Long, j As Long, m As Long, n As Long
Dim temp As Double, tmp_x As Double, tmp_y As Double

ReDim xy(1 To lattice_x, 1 To lattice_y, 1 To 2)
ReDim xy_Edges(1 To lattice_x, 1 To lattice_y, 1 To 7, 1 To 2)

For i = 1 To lattice_x
    tmp_x = i Mod 2
    For j = 1 To lattice_y
        xy(i, j, 1) = 1 + (i - 1) * 1.5 * L
        xy(i, j, 2) = 1 + (j - 0.5 - 0.5 * tmp_x) * Sqr(3) * L
    Next j
Next i

For i = 1 To lattice_x
    Application.StatusBar = "generating grid lines..." & i & "/" & lattice_x
    temp = i Mod 2
    For j = 1 To lattice_y
        tmp_x = xy(i, j, 1)
        tmp_y = xy(i, j, 2)
        xy_Edges(i, j, 1, 1) = tmp_x - L
        xy_Edges(i, j, 2, 1) = tmp_x - L * 0.5
        xy_Edges(i, j, 3, 1) = tmp_x + L * 0.5
        xy_Edges(i, j, 4, 1) = tmp_x + L
        xy_Edges(i, j, 5, 1) = tmp_x + L * 0.5
        xy_Edges(i, j, 6, 1) = tmp_x - L * 0.5
        xy_Edges(i, j, 7, 1) = tmp_x - L
        
        xy_Edges(i, j, 1, 2) = tmp_y
        xy_Edges(i, j, 2, 2) = tmp_y - L * Sqr(3) / 2
        xy_Edges(i, j, 3, 2) = tmp_y - L * Sqr(3) / 2
        xy_Edges(i, j, 4, 2) = tmp_y
        xy_Edges(i, j, 5, 2) = tmp_y + L * Sqr(3) / 2
        xy_Edges(i, j, 6, 2) = tmp_y + L * Sqr(3) / 2
        xy_Edges(i, j, 7, 2) = tmp_y
    Next j
Next i
Application.StatusBar = False

End Sub


'=== Print a detail U-Matrix
Sub Print_UMatrix_Det(mysht As Worksheet, Optional write_labels As Boolean = False, Optional node_label As Variant)
Dim i As Long, j As Long, m As Long, n As Long, k As Long
Dim vR As Long, vG As Long, vB As Long
Dim temp_i As Long, temp_j As Long
Dim temp As Double, temp_max As Double, temp_min As Double
Dim UMatrix_xy() As Double, UMatrix_xy_Edges() As Double

Call HexGen(2 * pL_x - 1, 2 * pL_y, 1, UMatrix_xy, UMatrix_xy_Edges)
With mysht
    .Range("P2:V100000").Clear
    m = 0
    k = 0
    For i = 1 To pL_x
    For j = 1 To pL_y
    
        temp_i = 2 * i - 1
        temp_j = 2 * j - (i Mod 2)

        m = m + 1
        .Range("R" & 1 + m).Value = UMatrix_xy(temp_i, temp_j, 1)
        .Range("S" & 1 + m).Value = UMatrix_xy(temp_i, temp_j, 2)
        For n = 1 To 7
        .Range("U" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i, temp_j, n, 1)
        .Range("V" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i, temp_j, n, 2)
        Next n
        k = k + 8
        
        If j < pL_y Then
        m = m + 1
        .Range("R" & 1 + m).Value = UMatrix_xy(temp_i, temp_j + 1, 1)
        .Range("S" & 1 + m).Value = UMatrix_xy(temp_i, temp_j + 1, 2)
        For n = 1 To 7
        .Range("U" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i, temp_j + 1, n, 1)
        .Range("V" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i, temp_j + 1, n, 2)
        Next n
        k = k + 8
        End If
        
        If i < pL_x Then
        If Not (i Mod 2 = 0 And j = pL_y) Then
        m = m + 1
        .Range("R" & 1 + m).Value = UMatrix_xy(temp_i + 1, temp_j + 1 - (temp_i Mod 2), 1)
        .Range("S" & 1 + m).Value = UMatrix_xy(temp_i + 1, temp_j + 1 - (temp_i Mod 2), 2)
        For n = 1 To 7
        .Range("U" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i + 1, temp_j + 1 - (temp_i Mod 2), n, 1)
        .Range("V" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i + 1, temp_j + 1 - (temp_i Mod 2), n, 2)
        Next n
        k = k + 8
        End If
        
        If Not (j = 1 And (i Mod 2) = 1) Then
        m = m + 1
        .Range("R" & 1 + m).Value = UMatrix_xy(temp_i + 1, temp_j - (temp_i Mod 2), 1)
        .Range("S" & 1 + m).Value = UMatrix_xy(temp_i + 1, temp_j - (temp_i Mod 2), 2)
        For n = 1 To 7
        .Range("U" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i + 1, temp_j - (temp_i Mod 2), n, 1)
        .Range("V" & 1 + k + n).Value = UMatrix_xy_Edges(temp_i + 1, temp_j - (temp_i Mod 2), n, 2)
        Next n
        k = k + 8
        End If
        End If
        
    Next j
    Next i
    
    With .ChartObjects("UMATRIX_Det").Chart.SeriesCollection(1)
    
    temp_max = -999999
    temp_min = 999999
    For i = 1 To pL_x
    For j = 1 To pL_y
    For n = 1 To 6
    If pnode_ED(i, j, n) > temp_max Then temp_max = pnode_ED(i, j, n)
    If pnode_ED(i, j, n) < temp_min Then temp_min = pnode_ED(i, j, n)
    Next n
    Next j
    Next i
    
    'Call Node_Max(node_UMatrix, temp_max, temp_min)
    If .HasDataLabels = True Then .DataLabels.Delete
    If write_labels = True Then .ApplyDataLabels Type:=xlDataLabelsShowValue, AutoText:=True, LegendKey:=False
    
    m = 0
    For i = 1 To pL_x
    For j = 1 To pL_y
    
        temp_i = 2 * i - 1
        temp_j = 2 * j - (i Mod 2)

        m = m + 1
        temp = (pnode_UMatrix(i, j) - temp_min) / (temp_max - temp_min)
        Call Gray_Scale(temp, vR)
        .Points(m).Format.Fill.ForeColor.RGB = RGB(vR, vR, vR)
        If write_labels = True Then
        With .Points(m).DataLabel
            If pnode_size(i, j) > 0 Then
                .Position = xlLabelPositionCenter
                .Text = node_label(i, j)
                .Format.TextFrame2.TextRange.Font.Size = 8
                If vR < 75 Then .Font.Color = RGB(255, 255, 255)
            Else
                .Delete
            End If
        End With
        End If

        
        If j < pL_y Then
        m = m + 1
        temp = (pnode_ED(i, j, 1) - temp_min) / (temp_max - temp_min)
        Call Gray_Scale(temp, vR)
        .Points(m).Format.Fill.ForeColor.RGB = RGB(vR, vR, vR)
        If write_labels = True Then .Points(m).DataLabel.Delete
        End If
        
        If i < pL_x Then
        If Not (i Mod 2 = 0 And j = pL_y) Then
        m = m + 1
        temp = (pnode_ED(i, j, 6) - temp_min) / (temp_max - temp_min)
        Call Gray_Scale(temp, vR)
        .Points(m).Format.Fill.ForeColor.RGB = RGB(vR, vR, vR)
        If write_labels = True Then .Points(m).DataLabel.Delete
        End If
        
        If Not (j = 1 And (i Mod 2) = 1) Then
        m = m + 1
        temp = (pnode_ED(i, j, 5) - temp_min) / (temp_max - temp_min)
        Call Gray_Scale(temp, vR)
        .Points(m).Format.Fill.ForeColor.RGB = RGB(vR, vR, vR)
        If write_labels = True Then .Points(m).DataLabel.Delete
        End If
        End If
        
    Next j
    Next i
    End With
    
End With
'========================================
End Sub
