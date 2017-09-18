VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPMFG_ListColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pList() As cPMFG_lcnode
Private pSize As Long

Sub LCNew(n As Long)
Dim i As Long
Dim v As cPMFG_lcnode
ReDim pList(0 To n - 1)
For i = 0 To n - 1
    Set v = New cPMFG_lcnode
    v.fwd = -1
    v.prev = -1
    Set pList(i) = v
Next i
pSize = n
End Sub


Sub LCReset()
Dim i As Long
For i = 0 To pSize - 1
    pList(i).fwd = -1
    pList(i).prev = -1
Next i
End Sub

Public Property Get List(i As Long) As cPMFG_lcnode
Set List = pList(i)
End Property

Public Property Let List(i As Long, lList As cPMFG_lcnode)
Set pList(i) = lList
End Property

Function LCGetNext(theList As Long, theNode As Long) As Long
Dim inext As Long

If theList = -1 Or theNode = -1 Then
    LCGetNext = -1
    Exit Function
End If

inext = pList(theNode).fwd
If theList = inext Then
    LCGetNext = -1
Else
    LCGetNext = inext
End If

End Function

Function LCGetPrev(theList As Long, theNode As Long) As Long
If theList = -1 Then
    LCGetPrev = -1
ElseIf theNode = -1 Then
    LCGetPrev = pList(theList).prev
ElseIf theNode = theList Then
    LCGetPrev = -1
Else
    LCGetPrev = pList(theNode).prev
End If
End Function

Function LCPrepend(theList As Long, theNode As Long) As Long
Dim newList As Long
newList = Me.LCAppend(theList, theNode)
If newList <> -1 Then newList = pList(newList).prev
LCPrepend = newList
End Function

Function LCAppend(theList As Long, theNode As Long) As Long
Dim Pred As Long
If theList = -1 Then
    pList(theNode).fwd = theNode
    pList(theNode).prev = theNode
    theList = theNode
Else
    Pred = pList(theList).prev
    
    pList(theList).prev = theNode
    pList(theNode).fwd = theList
    pList(theNode).prev = Pred
    pList(Pred).fwd = theNode
End If
LCAppend = theList
End Function


Function LCDelete(theList As Long, theNode As Long) As Long
Dim Pred As Long, succ As Long
If pList(theList).fwd = theList Then
    pList(theList).prev = -1
    pList(theList).fwd = -1
    theList = -1
Else
    Pred = pList(theNode).prev
    succ = pList(theNode).fwd
    
    pList(Pred).fwd = succ
    pList(succ).prev = Pred
    
    pList(theNode).prev = -1
    pList(theNode).fwd = -1
    
    If theList = theNode Then theList = succ
End If
LCDelete = theList
End Function
