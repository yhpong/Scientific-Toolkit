VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHeap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'======================================
'Binary Heap data structure
'only implemented for DOUBLE data type
'can be set as a MIN- or MAX-heap
'======================================

Private pQueue() As Double
Private pQueueID() As Long
Private pSize As Long
Private pHType As String

Public Property Get Size() As Long
    Size = pSize
End Property

Public Property Get HType() As String
    HType = pHType
End Property

Sub Init(Optional strType As String = "MIN")
    pSize = 0
    ReDim pQueue(0 To 0)
    ReDim pQueueID(0 To 0)
    pHType = VBA.UCase(strType)
End Sub

Sub Reset()
    pSize = 0
    Erase pQueue, pQueueID
End Sub

'Release memory from unused address in Queue
Sub Resize_Queue()
    ReDim Preserve pQueue(0 To pSize)
    ReDim Preserve pQueueID(0 To pSize)
End Sub


'================================================================================
'Build heap from input array x(1:N), data id is assumed to be same as input order
'================================================================================
Sub Build(x() As Double, Optional strType As String = "MIN")
Dim i As Long
    pHType = VBA.UCase(strType)
    pSize = UBound(x, 1)
    ReDim pQueue(0 To pSize)
    ReDim pQueueID(0 To pSize)
    For i = 1 To pSize
        pQueue(i) = x(i)
        pQueueID(i) = i
    Next i
    If pHType = "MIN" Then
        For i = pSize \ 2 To 1 Step -1
            Call Heapify_Min(i)
        Next i
    ElseIf pHType = "MAX" Then
        For i = pSize \ 2 To 1 Step -1
            Call Heapify_Max(i)
        Next i
    End If
End Sub

'Move i down the tree until heap condition is satisfied
Private Sub Heapify_Min(i As Long)
Dim vL As Long, vR As Long, v As Long, k As Long
Dim tmp_x As Double
    vL = 2 * i
    vR = 2 * i + 1
    v = i
    tmp_x = pQueue(i)
    If vL <= pSize And vR <= pSize Then
        If pQueue(vL) < tmp_x And pQueue(vL) <= pQueue(vR) Then
            v = vL
        ElseIf pQueue(vR) < tmp_x And pQueue(vR) <= pQueue(vL) Then
            v = vR
        End If
    ElseIf vL <= pSize Then
        If pQueue(vL) < tmp_x Then v = vL
    ElseIf vR <= pSize Then
        If pQueue(vR) < tmp_x Then v = vR
    End If
    If v <> i Then
        pQueue(i) = pQueue(v)
        pQueue(v) = tmp_x
        k = pQueueID(i)
        pQueueID(i) = pQueueID(v)
        pQueueID(v) = k
        Call Heapify_Min(v)
    End If
End Sub

Private Sub Heapify_Max(i As Long)
Dim vL As Long, vR As Long, v As Long, k As Long
Dim tmp_x As Double
    vL = 2 * i
    vR = 2 * i + 1
    v = i
    tmp_x = pQueue(i)
    If vL <= pSize And vR <= pSize Then
        If pQueue(vL) > tmp_x And pQueue(vL) >= pQueue(vR) Then
            v = vL
        ElseIf pQueue(vR) > tmp_x And pQueue(vR) >= pQueue(vL) Then
            v = vR
        End If
    ElseIf vL <= pSize Then
        If pQueue(vL) > tmp_x Then v = vL
    ElseIf vR <= pSize Then
        If pQueue(vR) > tmp_x Then v = vR
    End If
    If v <> i Then
        pQueue(i) = pQueue(v)
        pQueue(v) = tmp_x
        k = pQueueID(i)
        pQueueID(i) = pQueueID(v)
        pQueueID(v) = k
        Call Heapify_Max(v)
    End If
End Sub

'===========================================================================================
'Add data point x to end of heap, then move it up the tree until heap property is satisfied
'===========================================================================================
Sub Add(x As Double, id As Long)
Dim i As Long, n As Long, parent As Long
    pSize = pSize + 1
    ReDim Preserve pQueue(0 To pSize)
    ReDim Preserve pQueueID(0 To pSize)
    pQueue(pSize) = x
    pQueueID(pSize) = id
    If pSize = 1 Then Exit Sub
    
    n = pSize
    For i = 1 To pSize
        parent = n \ 2
        If pQueue(parent) <= x And pHType = "MIN" Then
            Exit For
        ElseIf pQueue(parent) >= x And pHType = "MAX" Then
            Exit For
        Else
            pQueue(n) = pQueue(parent)
            pQueueID(n) = pQueueID(parent)
            pQueue(parent) = x
            pQueueID(parent) = id
            n = parent
            If n = 1 Then Exit For
        End If
    Next i
End Sub

'======================================================================
'Pop_Min/Pop_Max should be called according to Heap type, code does not
'check it for you here. Non-matching type will return incorrect results.
'Queue() array is not resized, elements after pSize are no longer valid.
'======================================================================
Sub Pop_Min(x_out As Double, id_out As Long)
    x_out = pQueue(1)
    id_out = pQueueID(1)
    pQueue(1) = pQueue(pSize)
    pQueueID(1) = pQueueID(pSize)
    pSize = pSize - 1
    'Not necessary to resize if memory is not a concern
'    ReDim Preserve pQueue(0 To pSize)
'    ReDim Preserve pQueueID(0 To pSize)
    If pSize <= 1 Then Exit Sub
    Call Heapify_Min(1)
End Sub

Sub Pop_Max(x_out As Double, id_out As Long)
    x_out = pQueue(1)
    id_out = pQueueID(1)
    pQueue(1) = pQueue(pSize)
    pQueueID(1) = pQueueID(pSize)
    pSize = pSize - 1
    'Not necessary to resize if memory is not a concern
'    ReDim Preserve pQueue(0 To pSize)
'    ReDim Preserve pQueueID(0 To pSize)
    If pSize <= 1 Then Exit Sub
    Call Heapify_Max(1)
End Sub

'See top of queue without popping it
Sub Top(x As Double, i As Long)
    x = pQueue(1)
    i = pQueueID(1)
End Sub

'Sorting function incoporate here so it doesn't need to be re-written everytime
Sub Sort(x() As Double, sort_idx() As Long, Optional sort_dir As String = "ASCEND", _
    Optional init_sort_idx As Boolean = True)
Dim i As Long, n As Long, j As Long, iArr() As Long
    n = UBound(x)
    If init_sort_idx = True Then
        ReDim sort_idx(1 To n)
        For i = 1 To n
            sort_idx(i) = i
        Next i
    End If
    If sort_dir = "ASCEND" Then
        Call Build(x, "MIN")
        iArr = sort_idx
        For i = 1 To n
            Call Pop_Min(x(i), j)
            sort_idx(i) = iArr(j)
        Next i
    ElseIf sort_dir = "DESCEND" Then
        Call Build(x, "MAX")
        iArr = sort_idx
        For i = 1 To n
            Call Pop_Max(x(i), j)
            sort_idx(i) = iArr(j)
        Next i
    End If
End Sub
