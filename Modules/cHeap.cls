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
'Heap (min) data structure
'only implemented for DOUBLE data type
'======================================

Private pQueue() As Double
Private pQueueID() As Long
Private pSize As Long

Public Property Get Size() As Long
    Size = pSize
End Property

Sub Init()
    pSize = 0
    ReDim pQueue(0 To 0)
    ReDim pQueueID(0 To 0)
End Sub

Sub Reset()
    pSize = 0
    Erase pQueue
    Erase pQueueID
End Sub

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
        If pQueue(parent) <= x Then
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


Sub Pop_Min(x_out As Double, id_out As Long)
Dim i As Long, id As Long, n As Long, parent As Long, vL As Long, vR As Long
Dim x As Double, xL As Double, xr As Double, INFINITY As Double
    x_out = pQueue(1)
    id_out = pQueueID(1)
    pQueue(1) = pQueue(pSize)
    pQueueID(1) = pQueueID(pSize)
    pSize = pSize - 1
    ReDim Preserve pQueue(0 To pSize)
    ReDim Preserve pQueueID(0 To pSize)
    If pSize <= 1 Then Exit Sub
    INFINITY = Exp(70)
    n = 1
    x = pQueue(1)
    id = pQueueID(1)
    For i = 1 To pSize
        If (2 * n) <= pSize Then
            vL = 2 * n
            xL = pQueue(vL)
        Else
            vL = -1
            xL = INFINITY
        End If
        
        If (2 * n + 1) <= pSize Then
            vR = 2 * n + 1
            xr = pQueue(vR)
        Else
            vR = -1
            xr = INFINITY
        End If
        
        If x <= xL And x <= xr Then
            Exit For
        ElseIf (x > xL And x <= xr) Or (x > xL And xL <= xr) Then
            pQueue(n) = pQueue(vL)
            pQueue(vL) = x
            pQueueID(n) = pQueueID(vL)
            pQueueID(vL) = id
            n = vL
        ElseIf (x > xr And x <= xL) Or (x > xr And xr <= xL) Then
            pQueue(n) = pQueue(vR)
            pQueue(vR) = x
            pQueueID(n) = pQueueID(vR)
            pQueueID(vR) = id
            n = vR
        End If
    
    Next i
End Sub
