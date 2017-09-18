VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPMFG_Stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pTop As Long
Private pSize As Long
Private pStackList() As Long

Public Property Get Top() As Long
Top = pTop
End Property

Public Property Let Top(i As Long)
pTop = i
End Property

Public Property Get Size() As Long
Size = pSize
End Property

Public Property Let Size(i As Long)
pSize = i
End Property

Sub NewStack(Size As Long)
ReDim pStackList(0 To Size - 1)
Me.ClearStack
End Sub

Sub ClearStack()
pTop = 0
End Sub

Sub Pop(A As Long)
If pTop < 1 Then
    Debug.Print "Pop Fail, Top=" & pTop
    Exit Sub
End If
A = pStackList(pTop)
pTop = pTop - 1
End Sub

Sub Pop2(A As Long, B As Long)
If pTop <= 1 Then
    Debug.Print "Pop2 Fail, Top=" & pTop
    Exit Sub
End If
A = pStackList(pTop - 1)
B = pStackList(pTop)
pTop = pTop - 2
End Sub

Sub Push(A As Long)
pStackList(pTop + 1) = A
pTop = pTop + 1
End Sub


Sub Push2(A As Long, B As Long)
pStackList(pTop + 1) = A
pStackList(pTop + 2) = B
pTop = pTop + 2
End Sub
