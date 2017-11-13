VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTreeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*********************************************************
'Construct a Treemap using Squarified algorithm
'Main reference is "Squarified Treemaps, Mark Bruls (2000)"
'https://www.win.tue.nl/~vanwijk/stm.pdf
'Implementation detail refers to: https://github.com/imranghory/treemap-squared/blob/master/treemap-squarify.js
'*********************************************************

Private pheight As Double
Private pwidth As Double
Private px_pos As Double
Private py_pos As Double
Private pRectList() As Double
Private pRectList_Group As Variant


Public Property Get RectList() As Double()
    RectList = pRectList
End Property

Sub Init(x_pos As Double, y_pos As Double, width As Double, height As Double)
    pheight = height
    pwidth = width
    px_pos = x_pos
    py_pos = y_pos
End Sub

Sub Reset()
    Erase pRectList
    If IsArray(pRectList_Group) Then Erase pRectList_Group
End Sub

Private Function short_edge() As Double
    short_edge = pwidth
    If pheight < pwidth Then short_edge = pheight
End Function

Function getCoordinates(rowA() As Double) As Double()
Dim i As Long, n As Long
Dim areawidth As Double, areaheight As Double, sA As Double
Dim xArr() As Double
Dim tmp_x As Double, tmp_y As Double
    n = UBound(rowA)
    sA = 0
    For i = 1 To n
        sA = sA + rowA(i)
    Next i
    areawidth = sA / pheight
    areaheight = sA / pwidth
    tmp_x = px_pos
    tmp_y = py_pos
    ReDim xArr(1 To 4, 1 To n)
    If pwidth >= pheight Then
        For i = 1 To n
            xArr(1, i) = tmp_x
            xArr(2, i) = tmp_y
            xArr(3, i) = tmp_x + areawidth
            xArr(4, i) = tmp_y + rowA(i) / areawidth
            tmp_y = tmp_y + rowA(i) / areawidth
        Next i
    Else
        For i = 1 To n
            xArr(1, i) = tmp_x
            xArr(2, i) = tmp_y
            xArr(3, i) = tmp_x + rowA(i) / areaheight
            xArr(4, i) = tmp_y + areaheight
            tmp_x = tmp_x + rowA(i) / areaheight
        Next i
    End If
    getCoordinates = xArr
    Erase xArr
End Function

Function CutArea(area As Double) As cTreeMap
Dim areawidth As Double, areaheight As Double
Dim newwidth As Double, newheight As Double
Dim cRect As cTreeMap
    Set cRect = New cTreeMap
    If pwidth >= pheight Then
        areawidth = area / pheight
        newwidth = pwidth - areawidth
        Call cRect.Init(px_pos + areawidth, py_pos, newwidth, pheight)
    Else
        areaheight = area / pwidth
        newheight = pheight - areaheight
        Call cRect.Init(px_pos, py_pos + areaheight, pwidth, newheight)
    End If
    Set CutArea = cRect
End Function

Function Draw_Pts() As Variant
Dim i As Long, j As Long, n As Long
Dim vArr As Variant
    If ArrayIsEmpty(pRectList) = True Then
        Debug.Print "cTreeMap:Draw: Tree not build yet. Use Create."
        Exit Function
    End If
    n = UBound(pRectList, 2)
    ReDim vArr(1 To n, 1 To 2)
    j = 1
    For i = 1 To n
        vArr(i, 1) = (pRectList(1, i) + pRectList(3, i)) / 2
        vArr(i, 2) = (pRectList(2, i) + pRectList(4, i)) / 2
    Next i
    Draw_Pts = vArr
    Erase vArr
End Function

Function Draw_Lines() As Variant
Dim i As Long, j As Long, n As Long
Dim vArr As Variant
    If ArrayIsEmpty(pRectList) = True Then
        Debug.Print "cTreeMap:Draw: Tree not built yet. Use Create."
        Exit Function
    End If
    n = UBound(pRectList, 2)
    ReDim vArr(1 To 6 * n - 1, 1 To 2)
    j = 1
    For i = 1 To n
        vArr(j, 1) = pRectList(1, i)
        vArr(j, 2) = pRectList(2, i)
        vArr(j + 1, 1) = pRectList(3, i)
        vArr(j + 1, 2) = pRectList(2, i)
        vArr(j + 2, 1) = pRectList(3, i)
        vArr(j + 2, 2) = pRectList(4, i)
        vArr(j + 3, 1) = pRectList(1, i)
        vArr(j + 3, 2) = pRectList(4, i)
        vArr(j + 4, 1) = pRectList(1, i)
        vArr(j + 4, 2) = pRectList(2, i)
        j = j + 6
    Next i
    Draw_Lines = vArr
    Erase vArr
End Function


Function Draw_Lines_Group(k As Long) As Variant
Dim i As Long, j As Long, n As Long
Dim vArr As Variant, xArr() As Double
    'If map has no hierarchical structure,
    'simply print the overall border
    If IsArray(pRectList_Group) = False Then
        ReDim vArr(1 To 5, 1 To 2)
        vArr(1, 1) = px_pos
        vArr(1, 2) = py_pos
        vArr(2, 1) = px_pos + pwidth
        vArr(2, 2) = py_pos
        vArr(3, 1) = px_pos + pwidth
        vArr(3, 2) = py_pos + pheight
        vArr(4, 1) = px_pos
        vArr(4, 2) = py_pos + pheight
        vArr(5, 1) = px_pos
        vArr(5, 2) = py_pos
        Draw_Lines_Group = vArr
        Erase vArr
        Exit Function
    End If
    'Print borders for the k-th level groupings
    xArr = pRectList_Group(k)
    n = UBound(xArr, 2)
    ReDim vArr(1 To 6 * n - 1, 1 To 2)
    j = 1
    For i = 1 To n
        vArr(j, 1) = xArr(1, i)
        vArr(j, 2) = xArr(2, i)
        vArr(j + 1, 1) = xArr(3, i)
        vArr(j + 1, 2) = xArr(2, i)
        vArr(j + 2, 1) = xArr(3, i)
        vArr(j + 2, 2) = xArr(4, i)
        vArr(j + 3, 1) = xArr(1, i)
        vArr(j + 3, 2) = xArr(4, i)
        vArr(j + 4, 1) = xArr(1, i)
        vArr(j + 4, 2) = xArr(2, i)
        j = j + 6
    Next i
    Draw_Lines_Group = vArr
    Erase vArr, xArr
End Function


Sub Create(x() As Double, _
    Optional width As Double = 1, Optional height As Double = 1, _
    Optional x_pos As Double = 0, Optional y_pos As Double = 0, Optional normalize_x As Boolean = True)
Dim x2() As Double, currRow() As Double
    x2 = x
    If normalize_x = True Then Call normalize(x2, width * height)
    Call Init(x_pos, y_pos, width, height)
    ReDim currRow(0 To 0)
    ReDim pRectList(1 To 4, 0 To 0)
    Call Squarify(x2, currRow, pRectList)
    Erase x2, currRow
End Sub

Sub Create_Multilevel(x() As Double, x_group As Variant, _
    Optional width As Double = 1, Optional height As Double = 1, _
    Optional x_pos As Double = 0, Optional y_pos As Double = 0, Optional normalize_x As Boolean = True)
Dim i As Long, j As Long, k As Long, n As Long
Dim x2() As Double, x_id() As Long, xArr() As Double
Dim x_Tree As Variant
    n = UBound(x)
    x2 = x
    If normalize_x = True Then Call normalize(x2, width * height)
    Call Init(x_pos, y_pos, width, height)
    ReDim pRectList(1 To 4, 0 To 0)
    ReDim pRectList_Group(1 To UBound(x_group, 2))
    For i = 1 To UBound(x_group, 2)
        ReDim xArr(1 To 4, 0 To 0)
        pRectList_Group(i) = xArr
    Next i
    Call Build_Hierarchical_Array(x_Tree, x_group, x2)
    ReDim x_id(0 To 0)
    Call Create_Multilevel_Step(x2, x_Tree, width, height, x_pos, y_pos, x_id, 1)
    'Convert leafs back to original order
    x2 = pRectList
    For i = 1 To n
        j = x_id(i)
        For k = 1 To 4
            pRectList(k, j) = x2(k, i)
        Next k
    Next i
    Erase x2, x_id, x_Tree
End Sub


Private Sub Create_Multilevel_Step(x() As Double, x_Tree As Variant, _
    width As Double, height As Double, x_pos As Double, y_pos As Double, x_id() As Long, level As Long)
Dim i As Long, k As Long, n As Long
Dim x_ptr() As Long
Dim y() As Double, xArr() As Double
Dim xT1 As cTreeMap
    n = UBound(x_Tree)
    Set xT1 = New cTreeMap
    ReDim y(1 To n)
    If IsArray(x_Tree(1)) = True Then
        For k = 1 To n
            Call All_Leafs(x_Tree(k), x_ptr, 1)
            For i = 1 To UBound(x_ptr)
                y(k) = y(k) + x(x_ptr(i))
            Next i
        Next k
        Call xT1.Create(y, width, height, x_pos, y_pos, False)
        y = xT1.RectList
        
        xArr = pRectList_Group(level)
        Call Stack_Push(xArr, y)
        pRectList_Group(level) = xArr
        
        For k = 1 To n
            Call Create_Multilevel_Step(x, x_Tree(k), _
                y(3, k) - y(1, k), y(4, k) - y(2, k), y(1, k), y(2, k), x_id, level + 1)
        Next k
    Else
        For i = 1 To n
            y(i) = x(x_Tree(i))
            Call Append_Vec(x_id, x_Tree(i))    'save order of creation of leafs
        Next i
        Call xT1.Create(y, width, height, x_pos, y_pos, False)
        Call Stack_Push(pRectList, xT1.RectList)
    End If
    Call xT1.Reset
    Set xT1 = Nothing
    Erase x_ptr, y
End Sub


Sub Squarify(x() As Double, currRow() As Double, rStack() As Double)
Dim i As Long, j As Long, m As Long, n As Long, k As Long, n_raw As Long
Dim length As Double, s As Double
Dim nextData As Double, cRectNew As cTreeMap
Dim xArr() As Double, vArr As Variant
    n_raw = UBound(x, 1)
    n = UBound(rStack, 2)
    m = UBound(currRow)
    length = short_edge
    nextData = x(1)
    
    If ImprovesRatio(currRow, nextData, length) = True Then

        'Append data to current row
        ReDim Preserve currRow(0 To m + 1)
        currRow(m + 1) = nextData
        If n_raw > 1 Then
            'Process next node
            ReDim xArr(1 To n_raw - 1)
            For i = 1 To n_raw - 1
                xArr(i) = x(i + 1)
            Next i
            Call Squarify(xArr, currRow, rStack)

        Else
            'This is the last data, append to output stack
            Call Stack_Push(rStack, getCoordinates(currRow))
            
        End If
        
    Else
        
        'Fix current row
        Call Stack_Push(rStack, getCoordinates(currRow))
        
        'move to next row
        s = 0
        For i = 1 To m
            s = s + currRow(i)
        Next i
        Set cRectNew = CutArea(s)
        ReDim xArr(0 To 0)
        Call cRectNew.Squarify(x, xArr, rStack)
        
    End If
End Sub

Private Function ImprovesRatio(currRow() As Double, nextnode As Double, w As Double) As Boolean
Dim i As Long, n As Long
Dim tmp_x As Double, tmp_y As Double
Dim newRow() As Double
    ImprovesRatio = False
    n = UBound(currRow)
    If n = 0 Then
        ImprovesRatio = True
        Exit Function
    End If
    newRow = currRow
    ReDim Preserve newRow(0 To n + 1)
    newRow(n + 1) = nextnode
    tmp_x = Calc_Ratio(currRow, w)
    tmp_y = Calc_Ratio(newRow, w)
    If tmp_x >= tmp_y Then ImprovesRatio = True
    Erase newRow
End Function

Private Function Calc_Ratio(rowA() As Double, w As Double) As Double
Dim i As Long
Dim s As Double, Amax As Double, Amin As Double
Dim tmp_x As Double
    s = 0
    Amax = -Exp(70)
    Amin = Exp(70)
    For i = 1 To UBound(rowA)
        s = s + rowA(i)
        If rowA(i) > Amax Then Amax = rowA(i)
        If rowA(i) < Amin Then Amin = rowA(i)
    Next i
    Calc_Ratio = (w ^ 2) * Amax / (s ^ 2)
    tmp_x = (s ^ 2) / ((w ^ 2) * Amin)
    If tmp_x > Calc_Ratio Then Calc_Ratio = tmp_x
End Function

'Normalize data so their values sum to designated area
Private Sub normalize(x() As Double, area As Double) 'As Double()
Dim i As Long, n As Long
Dim tmp_x As Double
    n = UBound(x)
    tmp_x = 0
    For i = 1 To n
        tmp_x = tmp_x + x(i)
    Next i
    tmp_x = area / tmp_x
    For i = 1 To n
        x(i) = x(i) * tmp_x
    Next i
End Sub



'===================================================
'General operations and data structure
'===================================================

'Input: x_group(1:N,1:D), N samples of D-levels hierarchical data
'Output: x_Tree(), nested array of depth D
Private Sub Build_Hierarchical_Array(x_Tree As Variant, x_group As Variant, x() As Double)
Dim i As Long, n As Long
Dim x_ptr As Variant
    n = UBound(x_group, 1)
    ReDim x_ptr(1 To n)
    For i = 1 To n
        x_ptr(i) = i
    Next i
    Call Group_Vector(x_Tree, x_group, x_ptr, 1, x)
    Erase x_ptr
End Sub

Private Sub Group_Vector(x_Tree As Variant, x_group As Variant, x_ptr As Variant, k As Long, x() As Double)
Dim i As Long, j As Long, m As Long, n As Long, n_group As Long
Dim n_raw As Long, n_dimension As Long
Dim vArr As Variant, xArr As Variant
Dim iArr() As Long, jArr() As Long, sort_idx() As Long
Dim x_i() As Long, x_list_size() As Long

    n_dimension = UBound(x_group, 2)    'number of levels
    n = UBound(x_ptr)                   'size of current subset
    
    ReDim vArr(1 To n)
    For i = 1 To n
        vArr(i) = x_group(x_ptr(i), k)
    Next i
    
    Call modMath.Unique_Items(vArr, x_i, xArr, n_group, x_list_size, True)
    
    'Sort group based on their total value of x()
    ReDim vArr(1 To n_group)
    For i = 1 To n
        j = x_i(i)
        vArr(j) = vArr(j) + x(x_ptr(i))
    Next i
    Call modMath.Sort_Quick_A(vArr, 1, n_group, iArr, 1)
    ReDim jArr(1 To n_group)
    ReDim sort_idx(1 To n_group)
    For j = 1 To n_group
        jArr(j) = iArr(n_group - j + 1)
        sort_idx(iArr(n_group - j + 1)) = j
    Next j
    iArr = x_list_size
    For j = 1 To n_group
        x_list_size(j) = iArr(jArr(j))
    Next j
    For i = 1 To n
        x_i(i) = sort_idx(x_i(i))
    Next i

    'elements to be passed onto each child
    ReDim vArr(1 To n_group)
    For j = 1 To n_group
        ReDim xArr(1 To x_list_size(j))
        vArr(j) = xArr
    Next j
    ReDim iArr(1 To n_group)
    For i = 1 To n
        j = x_i(i)
        iArr(j) = iArr(j) + 1
        vArr(j)(iArr(j)) = x_ptr(i)
    Next i

    If k = n_dimension Then
        x_Tree = vArr
    ElseIf k < n_dimension Then
        ReDim x_Tree(1 To n_group)
        For j = 1 To n_group
            Call Group_Vector(x_Tree(j), x_group, vArr(j), k + 1, x)
        Next j
    End If
    Erase vArr, xArr, iArr, jArr, sort_idx, x_i, x_list_size
End Sub

'Find all leafs under a tree and return a vector of their pointers
Private Sub All_Leafs(x_Tree As Variant, x_ptr As Variant, Optional first_run As Long = 1)
Dim i As Long, j As Long, k As Long, m As Long, n As Long
    If first_run = 1 Then ReDim x_ptr(0 To 0)
    n = UBound(x_Tree, 1)
    If VBA.IsArray(x_Tree(1)) = False Then
        m = UBound(x_ptr)
        n = UBound(x_Tree)
        If m = 0 Then
            ReDim x_ptr(1 To n)
        Else
            ReDim Preserve x_ptr(1 To m + n)
        End If
        For i = 1 To n
            x_ptr(m + i) = x_Tree(i)
        Next i
    Else
        For i = 1 To n
            Call All_Leafs(x_Tree(i), x_ptr, 0)
        Next i
    End If
End Sub

'Append vector x() to array rStack()
Private Sub Stack_Push(rStack() As Double, x() As Double)
Dim i As Long, j As Long, m As Long, n As Long, k As Long
    n = UBound(rStack, 2)
    m = UBound(x, 2)
    ReDim Preserve rStack(LBound(rStack, 1) To UBound(rStack, 1), LBound(rStack, 2) To n + m)
    For i = 1 To m
        For j = LBound(rStack, 1) To UBound(rStack, 1)
            rStack(j, n + i) = x(j, i)
        Next j
    Next i
End Sub

'Append a value v to vector x()
Private Sub Append_Vec(x As Variant, v As Variant)
Dim n As Long
    n = UBound(x, 1)
    ReDim Preserve x(LBound(x, 1) To n + 1)
    x(n + 1) = v
End Sub

'Check if array is allocated
Private Function ArrayIsEmpty(A As Variant) As Boolean
Dim i As Long
    ArrayIsEmpty = False
    i = -1
    On Error Resume Next
    i = UBound(A)
    If i = -1 Then ArrayIsEmpty = True
    On Error GoTo 0
End Function
