VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdjList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pHead As cAdjList_Node
Private pSize As Long

Public Property Get Size() As Long
    Size = pSize
End Property

Sub Init()
    Set pHead = Nothing
    pSize = 0
End Sub

Sub Reset()
Dim u As cAdjList_Node, v As cAdjList_Node
    If pHead Is Nothing Then Exit Sub
    Set u = pHead
    Set pHead = Nothing
    Do While Not u Is Nothing
        Set v = u.Nxt
        Set u = Nothing
        Set u = v
    Loop
    pSize = 0
End Sub

Function LastNode() As cAdjList_Node
Dim v As cAdjList_Node
    Set v = pHead
    Do While Not v Is Nothing
        Set LastNode = v
        Set v = v.Nxt
    Loop
End Function

Function Append(v As Long, Optional Wgt As Double = 1, Optional replace_wgt As Long = 0, _
            Optional check_duplicate As Boolean = True) As Long
Dim i As Long
Dim v_node As cAdjList_Node, v_last As cAdjList_Node, u_node As cAdjList_Node
    Append = 0
    Set v_node = New cAdjList_Node
    Call v_node.Init(v, Wgt)
    If pHead Is Nothing Then
        Set pHead = v_node
        Append = 1
    Else
        Set u_node = pHead
        If check_duplicate = True Then
            For i = 1 To pSize
                If u_node.Ptr = v Then 'Edge already exists.
                    If replace_wgt = 1 Then
                        u_node.Wgt = Wgt                'Replace wgt by new value
                    ElseIf replace_wgt = 2 Then
                        u_node.Wgt = u_node.Wgt + Wgt   'Accumulate wgt
                    End If
                    Exit Function
                End If
                If i < pSize Then Set u_node = u_node.Nxt
            Next i
            Set u_node.Nxt = v_node
            Append = 1
        Else
            For i = 2 To pSize
                Set u_node = u_node.Nxt
            Next i
            Set u_node.Nxt = v_node
            Append = 1
        End If
    End If
    pSize = pSize + Append
End Function


Function Prepend(v As Long, Optional Wgt As Double = 1, Optional replace_wgt As Long = 0, _
            Optional check_duplicate As Boolean = True) As Long
Dim v_node As cAdjList_Node, u_node As cAdjList_Node
    Prepend = 0
    Set v_node = New cAdjList_Node
    Call v_node.Init(v, Wgt)
    If pHead Is Nothing Then
        Set pHead = v_node
        Prepend = 1
    Else
        If check_duplicate = True Then
            Set u_node = pHead
            Do While Not u_node Is Nothing
                If u_node.Ptr = v Then 'Edge already exists.
                    If replace_wgt = 1 Then
                        u_node.Wgt = Wgt                'Replace wgt by new value
                    ElseIf replace_wgt = 2 Then
                        u_node.Wgt = u_node.Wgt + Wgt   'Accumulate wgt
                    End If
                    Exit Function
                End If
                Set u_node = u_node.Nxt
            Loop
        End If
        Set u_node = pHead
        Set pHead = v_node
        Set v_node.Nxt = u_node
        Prepend = 1
    End If
    pSize = pSize + Prepend
End Function


'Delete node with key v, if there are multiple occurences then the first occurence is deleted
Function Remove(v As Long) As Long
Dim i As Long
Dim u_node As cAdjList_Node, v_node As cAdjList_Node
    Remove = 0
    If pSize = 0 Then Exit Function
    If pHead.Ptr = v Then
        Set pHead = pHead.Nxt
        Remove = 1
        pSize = pSize - Remove
        Exit Function
    End If
    Set u_node = pHead
    For i = 2 To pSize
        Set v_node = u_node.Nxt
        If v_node.Ptr = v Then
            Set u_node.Nxt = v_node.Nxt
            Set v_node.Nxt = Nothing
            Remove = 1
            pSize = pSize - Remove
            Set v_node = Nothing
            Exit Function
        End If
        Set u_node = v_node
    Next i
End Function

'Delete node at the i-th position
Function Remove_by_pos(i As Long) As Long
Dim j As Long
Dim u_node As cAdjList_Node, u_nxt As cAdjList_Node, u_prev As cAdjList_Node
    Remove_by_pos = 0
    If pSize = 0 Or i > pSize Then Exit Function
    If i = 1 Then
        If pSize > 1 Then
            Set u_node = pHead.Nxt
            Set pHead = u_node
        ElseIf pSize = 1 Then
            Set pHead = Nothing
        End If
        Remove_by_pos = 1
        pSize = pSize - Remove_by_pos
        Exit Function
    End If
    
    j = 0
    Set u_node = pHead
    Do While Not u_node Is Nothing
        j = j + 1
        If j = i Then
            Set u_nxt = u_node.Nxt
            Set u_node.Nxt = Nothing
            Set u_prev.Nxt = u_nxt
            Remove_by_pos = 1
            pSize = pSize - Remove_by_pos
            Exit Function
        End If
        Set u_prev = u_node
        Set u_node = u_node.Nxt
    Loop
End Function

'Return a vector of keys in the list
Function GetList() As Long()
Dim i As Long, v As cAdjList_Node
Dim iArr() As Long
    If Not pHead Is Nothing Then
        ReDim iArr(1 To pSize)
        Set v = pHead
        For i = 1 To pSize
            iArr(i) = v.Ptr
            Set v = v.Nxt
        Next i
    Else
        ReDim iArr(0 To 0)
    End If
    GetList = iArr
End Function

'Return a vector of weights in the list
Function GetListWgt() As Double()
Dim i As Long, v As cAdjList_Node
Dim xArr() As Double
    If Not pHead Is Nothing Then
        ReDim xArr(1 To pSize)
        Set v = pHead
        For i = 1 To pSize
            xArr(i) = v.Wgt
            Set v = v.Nxt
        Next i
    Else
        ReDim xArr(0 To 0)
    End If
    GetListWgt = xArr
End Function

'Return both keys and weights in the list
Sub GetListnWgt(iArr() As Long, xArr() As Double)
Dim i As Long, v As cAdjList_Node
    If Not pHead Is Nothing Then
        ReDim iArr(1 To pSize)
        ReDim xArr(1 To pSize)
        Set v = pHead
        For i = 1 To pSize
            iArr(i) = v.Ptr
            xArr(i) = v.Wgt
            Set v = v.Nxt
        Next i
    Else
        ReDim iArr(0 To 0)
        ReDim xArr(0 To 0)
    End If
End Sub

'Return key (and optional weight) of node at the i-th position
Function GetNode(i As Long, Optional Wgt As Variant, Optional getWgt As Boolean = False) As Long
Dim j As Long
Dim u As cAdjList_Node
    j = 0
    Set u = pHead
    Do While Not u Is Nothing
        j = j + 1
        If j = i Then
            GetNode = u.Ptr
            If getWgt = True Then Wgt = u.Wgt
            Exit Do
        End If
        Set u = u.Nxt
    Loop
End Function

'Returns positive integer if v is in adjlist and zero otherwise, wgt is also return
Function isConnect(v As Long, Wgt As Double) As Long
Dim k As Long
Dim u_node As cAdjList_Node
    isConnect = 0: Wgt = 0
    If pHead Is Nothing Then
        Exit Function
    Else
        Set u_node = pHead
        For k = 1 To pSize
            If u_node.Ptr = v Then 'v is in the list at address k
                isConnect = k
                Wgt = u_node.Wgt
                Exit Function
            End If
            Set u_node = u_node.Nxt
        Next k
    End If
End Function


'Replace the whole list by v() and optional wgts()
Function ReplaceList(v() As Long, Optional wgts As Variant, Optional isWgt As Boolean = False) As Long
Dim i As Long
Dim v_node As cAdjList_Node, u_node As cAdjList_Node
    Call Init
    pSize = UBound(v)
    ReplaceList = pSize
    If pSize = 0 Then Exit Function
    If isWgt = True Then
        Set pHead = New cAdjList_Node
        Call pHead.Init(v(1), VBA.CDbl(wgts(1)))
        
        Set v_node = pHead
        For i = 2 To pSize
            Set u_node = New cAdjList_Node
            Call u_node.Init(v(i), VBA.CDbl(wgts(i)))
            Set v_node.Nxt = u_node
            Set v_node = u_node
        Next i
    Else
        Set pHead = New cAdjList_Node
        Call pHead.Init(v(1))
        
        Set v_node = pHead
        For i = 2 To pSize
            Set u_node = New cAdjList_Node
            Call u_node.Init(v(i))
            Set v_node.Nxt = u_node
            Set v_node = u_node
        Next i
    End If
End Function


'Reorder node position
'Input: integer vector iorder(i)=v means than node with key v is moved to the i-th position
'Output: optional output swap_idx(i)=j means that the new position i is mapped from the original j-th poistion
Sub Reorder(iorder() As Long, Optional swap_idx As Variant, Optional output_swap_idx As Boolean = False)
Dim i As Long, j As Long, k As Long, n As Long, v As cAdjList_Node
Dim wgt_new() As Double, idx_new() As Long
    If UBound(iorder) <> pSize Then
        Debug.Print "cAdjList: Reorder: Failed. Size is not consistent."
        Exit Sub
    End If
    ReDim idx_new(1 To pSize)
    ReDim wgt_new(1 To pSize)
    For i = 1 To pSize
        idx_new(i) = isConnect(iorder(i), wgt_new(i))
        If idx_new(i) <= 0 Then
            Debug.Print "cAdjList: Reorder: Failed. Node " & iorder(i) & " does not exist."
            Exit Sub
        End If
    Next i
    Set v = pHead
    For i = 1 To pSize
        v.Ptr = iorder(i)
        v.Wgt = wgt_new(i)
        Set v = v.Nxt
    Next i
    If output_swap_idx = True Then swap_idx = idx_new
End Sub
