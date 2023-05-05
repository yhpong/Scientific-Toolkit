VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cQueue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pSize As Long
Private pHead As Long, pTail As Long
Private pList() As Variant

Sub Init(Optional maxSize As Long = 0)
    pSize = 0
    pHead = 0
    pTail = 0
    ReDim pList(0 To maxSize)
End Sub

Property Get Head()
    Head = pList(pHead)
End Property

Property Get Tail()
    Tail = pList(pTail)
End Property

Property Get size() As Long
    size = pSize
End Property

Sub Add(x)
    pSize = pSize + 1
    If pSize = 1 Then
        pHead = 1
        pTail = 1
        ReDim Preserve pList(0 To 1)
        pList(1) = x
    Else
        pTail = pTail + 1
        If pTail > UBound(pList, 1) Then
            ReDim Preserve pList(0 To pTail)
        End If
        pList(pTail) = x
    End If
End Sub

Function Pop()
    If pSize = 0 Then
        Debug.Print "cQueue: Pop: Failed, queue is empty."
        Exit Function
    End If
    pSize = pSize - 1
    Pop = pList(pHead)
    pList(pHead) = 0
    pHead = pHead + 1
End Function

Sub Purge()
Dim i As Long
    If pHead = 1 Then Exit Sub
    For i = 1 To pSize
        pList(i) = pList(pHead + i - 1)
    Next i
    pHead = 1
    pTail = pSize
End Sub



