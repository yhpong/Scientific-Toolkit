Attribute VB_Name = "gp_Embed"
Option Explicit

Sub PlanarityTest()
Dim i As Long, j As Long, k As Long
Dim theGraph As cPMFG_Graph
Dim mywkbk As Workbook
Dim isPlanar As String
Dim strtemp As String


msgbox "ABC"

Set mywkbk = ActiveWorkbook
Call gp.ReadAdjMatrix(theGraph, mywkbk)
isPlanar = Embed(theGraph)

Debug.Print isPlanar

With mywkbk.Sheets("Chart")
.Range("R2:Y1000").Clear
For i = 0 To (2 * theGraph.N_Nodes - 1 + 2 * theGraph.M_Edges)
.Range("R" & 2 + i).Value = i
.Range("S" & 2 + i).Value = theGraph.g(i).index
.Range("T" & 2 + i).Value = theGraph.g(i).Link(0)
.Range("U" & 2 + i).Value = theGraph.g(i).Link(1)
.Range("V" & 2 + i).Value = theGraph.g(i).etype
If i < theGraph.N_Nodes Then
.Range("W" & 2 + i).Value = theGraph.Vertices(i).DFSParent
.Range("X" & 2 + i).Value = theGraph.Vertices(i).Lowpoint
.Range("Y" & 2 + i).Value = theGraph.Vertices(i).leastAncestor
End If
Next i
End With

'For i = 0 To theGraph.N_Nodes - 1
'strtemp = i & " :"
'j = theGraph.Vertices(i).separatedDFSChildList
'
'k = theGraph.DFSChildLists.LCGetPrev(j, -1)
'Do While k <> -1
'    strtemp = strtemp & " " & k
'    k = theGraph.DFSChildLists.LCGetPrev(j, k)
'Loop
'Debug.Print strtemp
'Next i

Set mywkbk = Nothing

End Sub


Function Embed(theGraph As cPMFG_Graph) As String
Dim n As Long, i As Long, j As Long, w As Long, child As Long
Dim RetVal As String

If theGraph.N_Nodes <= 4 Then
    Embed = "OK"
    Exit Function
ElseIf theGraph.N_Nodes > 4 Then
    If theGraph.M_Edges > (3 * theGraph.N_Nodes - 6) Then
        Embed = "NONPLANAR"
        Exit Function
    End If
End If

n = theGraph.N_Nodes
Call CreateDFSTree(theGraph)
Call SortVertices(theGraph)
Call LowpointAndLeastAncestor(theGraph)
Call CreateSortedSeparatedDFSChildLists(theGraph)
Call CreateFwdArcLists(theGraph)
Call CreateDFSTreeEmbedding(theGraph)
Call gp.FillVisitedFlags(theGraph, n)

For i = n - 1 To 0 Step -1
    
    RetVal = "OK"

    j = theGraph.Vertices(i).fwdArcList
    Do While j <> -1
        w = theGraph.g(j).index
        theGraph.Vertices(w).adjacentTo = j
        Call Walkup(theGraph, i, w)

        j = theGraph.g(j).Link(0)
        If j = theGraph.Vertices(i).fwdArcList Then j = -1
    Loop

    child = theGraph.Vertices(i).separatedDFSChildList
    Do While child <> -1
        If theGraph.Vertices(child).pertinentBicompList <> -1 Then
        Call WalkDown(theGraph, i, child + n)
        End If
        child = theGraph.DFSChildLists.LCGetNext(theGraph.Vertices(i).separatedDFSChildList, child)
    Loop

    If theGraph.Vertices(i).fwdArcList <> -1 Then
        RetVal = "NONPLANAR"
        Exit For
    End If

Next i

Embed = RetVal

End Function


Sub CreateDFSTree(graph As cPMFG_Graph)
Dim DFI As Long, u As Long, i As Long, uparent As Long, e As Long, j As Long
Dim n As Long
Dim theStack As cPMFG_Stack

Set theStack = graph.Stack
Call theStack.ClearStack

With graph
    n = .N_Nodes
    For i = 0 To n - 1
        .g(i).visited = 0
    Next i
End With

With graph
i = 0
DFI = 0

'This outer loop causes the connected subgraphs of a disconnected graph to be numbered
Do While i < n And DFI < n
    
    If .Vertices(i).DFSParent <> -1 Then GoTo has_parent_already
    
    Call theStack.Push2(-1, -1)
    
    Do While theStack.Top > 0
        Call theStack.Pop2(uparent, e)
        If uparent = -1 Then
            u = i
        Else
            u = .g(e).index
        End If
        
        
        If .g(u).visited = 0 Then
            .g(u).visited = 1
            .g(u).index = DFI
            .Vertices(u).DFSParent = uparent
            DFI = DFI + 1
            If e <> -1 Then
                .g(e).etype = "EDGE_DFSCHILD"
                
                'Delete the edge from the list
                .g(.g(e).Link(0)).Link(1) = .g(e).Link(1)
                .g((.g(e).Link(1))).Link(0) = .g(e).Link(0)
                
                'Tell the edge where it belongs now
                .g(e).Link(0) = .g(uparent).Link(0)
                .g(e).Link(1) = uparent
                
                'Tell the rest of the list where the edge belongs
                .g(uparent).Link(0) = e
                .g(.g(e).Link(0)).Link(1) = e
            End If

            'Push all neighbors
            j = .g(u).Link(0)
            Do While j >= n
                Call theStack.Push2(u, j)
                j = .g(j).Link(0)
            Loop
            
        Else
        
            If .g(uparent).index < .g(u).index Then
                .g(e).etype = "EDGE_FORWARD"
                .g(.g(e).Link(0)).Link(1) = .g(e).Link(1)
                .g(.g(e).Link(1)).Link(0) = .g(e).Link(0)
                .g(e).Link(0) = uparent
                .g(e).Link(1) = .g(uparent).Link(1)
                .g(uparent).Link(1) = e
                .g(.g(e).Link(1)).Link(0) = e
            ElseIf .g(gp.GetTwinArc(graph, e)).etype = "EDGE_DFSCHILD" Then
                .g(e).etype = "EDGE_DFSPARENT"
            Else
                .g(e).etype = "EDGE_BACK"
            End If
        
        End If
    
    Loop
    
has_parent_already:
    i = i + 1
Loop
End With

End Sub


Sub SortVertices(theGraph As cPMFG_Graph)
Dim i As Long, n As Long, m As Long, e As Long, j As Long, srcPos As Long, dstPos As Long
Dim tempV As cPMFG_VertexRec
Dim tempG As cPMFG_Node
With theGraph

n = .N_Nodes
m = .M_Edges

e = 0
j = 2 * n
Do While e < m
    .g(j).index = .g(.g(j).index).index
    If .g(j).Link(0) < n Then .g(j).Link(0) = .g(.g(j).Link(0)).index
    If .g(j).Link(1) < n Then .g(j).Link(1) = .g(.g(j).Link(1)).index
    
    .g(j + 1).index = .g(.g(j + 1).index).index
    If .g(j + 1).Link(0) < n Then .g(j + 1).Link(0) = .g(.g(j + 1).Link(0)).index
    If .g(j + 1).Link(1) < n Then .g(j + 1).Link(1) = .g(.g(j + 1).Link(1)).index
    
    e = e + 1
    j = j + 2
Loop

For i = 0 To n - 1
    If .Vertices(i).DFSParent <> -1 Then .Vertices(i).DFSParent = .g(.Vertices(i).DFSParent).index
Next i

For i = 0 To n - 1
    .g(i).visited = 0
Next i

For i = 0 To n - 1
    srcPos = i
    Do While .g(i).visited = 0
        dstPos = .g(i).index
        Set tempG = .g(dstPos)
        Set tempV = .Vertices(dstPos)
        Set .g(dstPos) = .g(i)
        Set .Vertices(dstPos) = .Vertices(i)
        Set .g(i) = tempG
        Set .Vertices(i) = tempV
        .g(dstPos).visited = 1
        .g(dstPos).index = srcPos
        srcPos = dstPos
    Loop
Next i

End With
End Sub


Sub LowpointAndLeastAncestor(theGraph As cPMFG_Graph)
Dim theStack As cPMFG_Stack
Dim i As Long, u As Long, uneighbor As Long, j As Long, L As Long, leastAncestor As Long
Dim n As Long

Set theStack = New cPMFG_Stack
Set theStack = theGraph.Stack

Call theStack.ClearStack

With theGraph

n = .N_Nodes
For i = 0 To n - 1
    .g(i).visited = 0
Next i

For i = 0 To n - 1
If .g(i).visited = 0 Then
    
    Call theStack.Push(i)
    Do While theStack.Top > 0
        Call theStack.Pop(u)
        If .g(u).visited = 0 Then
        
            .g(u).visited = 1
            Call theStack.Push(u)
            
            j = .g(u).Link(0)
            Do While j >= n
                If .g(j).etype = "EDGE_DFSCHILD" Then
                    Call theStack.Push(.g(j).index)
                Else
                    Exit Do
                End If
                j = .g(j).Link(0)
            Loop
        
        Else
            
            leastAncestor = u
            L = u
        
            j = .g(u).Link(0)
            Do While j >= n
                uneighbor = .g(j).index
                If .g(j).etype = "EDGE_DFSCHILD" Then
                    If L > .Vertices(uneighbor).Lowpoint Then L = .Vertices(uneighbor).Lowpoint
                ElseIf .g(j).etype = "EDGE_BACK" Then
                    If leastAncestor > uneighbor Then leastAncestor = uneighbor
                ElseIf .g(j).etype = "EDGE_FORWARD" Then
                    Exit Do
                End If
                j = .g(j).Link(0)
            Loop
            
            .Vertices(u).leastAncestor = leastAncestor
            .Vertices(u).Lowpoint = fMIN2(L, leastAncestor)

        End If
    Loop
    
End If
Next i

End With

End Sub


Sub CreateSortedSeparatedDFSChildLists(theEmbedding As cPMFG_Graph)
Dim i As Long, j As Long, n As Long, DFSParent As Long, theList As Long
Dim buckets() As Long
Dim Bins As cPMFG_ListColl
With theEmbedding

n = .N_Nodes

Set Bins = .bin
Bins.LCReset
ReDim buckets(0 To n - 1)
For i = 0 To n - 1
    buckets(i) = -1
Next i

For i = 0 To n - 1
    j = .Vertices(i).Lowpoint
    buckets(j) = Bins.LCAppend(buckets(j), i)
Next i

For i = 0 To n - 1
    j = buckets(i)
    If j <> -1 Then
        Do While j <> -1
            DFSParent = .Vertices(j).DFSParent
            If DFSParent <> -1 And DFSParent <> j Then
                theList = .Vertices(DFSParent).separatedDFSChildList
                theList = .DFSChildLists.LCAppend(theList, j)
                .Vertices(DFSParent).separatedDFSChildList = theList
            End If
            j = Bins.LCGetNext(buckets(i), j)
        Loop
    End If
Next i

End With
End Sub


Sub CreateFwdArcLists(theGraph As cPMFG_Graph)
Dim i As Long, jfirst As Long, jnext As Long, jlast As Long
Dim n  As Long
With theGraph

n = .N_Nodes

For i = 0 To n - 1
    jfirst = .g(i).Link(1)
    If .g(jfirst).etype = "EDGE_FORWARD" Then
        jnext = jfirst
        Do While .g(jnext).etype = "EDGE_FORWARD"
            jnext = .g(jnext).Link(1)
        Loop
        jlast = .g(jnext).Link(0)
        
        .g(jnext).Link(0) = i
        .g(i).Link(1) = jnext
        
        .Vertices(i).fwdArcList = jfirst
        .g(jfirst).Link(0) = jlast
        .g(jlast).Link(1) = jfirst
    End If
Next i

End With
End Sub


Sub CreateDFSTreeEmbedding(theGraph As cPMFG_Graph)
Dim i As Long, j As Long, jtwin As Long, n As Long, r As Long
With theGraph

n = .N_Nodes
i = 0
r = n
Do While i < n
    If .Vertices(i).DFSParent = -1 Then
        .g(i).Link(0) = i
        .g(i).Link(1) = i
    Else
        j = .g(i).Link(0)
        Do While .g(j).etype <> "EDGE_DFSPARENT"
            j = .g(j).Link(0)
        Loop
        
        .g(i).Link(0) = j
        .g(i).Link(1) = j
        .g(j).Link(0) = i
        .g(j).Link(1) = i
        .g(j).index = r
        
        jtwin = gp.GetTwinArc(theGraph, j)
        
        .g(r).Link(0) = jtwin
        .g(r).Link(1) = jtwin
        .g(jtwin).Link(0) = r
        .g(jtwin).Link(1) = r
        
        .extFace(r).Link(0) = i
        .extFace(r).Link(1) = i
        .extFace(i).Link(0) = r
        .extFace(i).Link(1) = r
    End If

    i = i + 1
    r = r + 1
Loop

End With
End Sub


Sub EmbedBackEdgeToDescendant(theGraph As cPMFG_Graph, RootSide As Long, RootVertex As Long, w As Long, WPrevLink As Long)
Dim fwdArc As Long, backArc As Long, parentCopy As Long
With theGraph

fwdArc = .Vertices(w).adjacentTo
backArc = gp.GetTwinArc(theGraph, fwdArc)

parentCopy = .Vertices(RootVertex - .N_Nodes).DFSParent

If .Vertices(parentCopy).fwdArcList = fwdArc Then
    If .g(fwdArc).Link(0) = fwdArc Then
        .Vertices(parentCopy).fwdArcList = -1
    Else
        .Vertices(parentCopy).fwdArcList = .g(fwdArc).Link(0)
    End If
End If

.g(.g(fwdArc).Link(0)).Link(1) = .g(fwdArc).Link(1)
.g(.g(fwdArc).Link(1)).Link(0) = .g(fwdArc).Link(0)

.g(fwdArc).Link(fXOR(1, RootSide)) = RootVertex
.g(fwdArc).Link(RootSide) = .g(RootVertex).Link(RootSide)
.g(.g(RootVertex).Link(RootSide)).Link(fXOR(1, RootSide)) = fwdArc
.g(RootVertex).Link(RootSide) = fwdArc

.g(backArc).index = RootVertex

.g(backArc).Link(fXOR(1, WPrevLink)) = w
.g(backArc).Link(WPrevLink) = .g(w).Link(WPrevLink)
.g(.g(w).Link(WPrevLink)).Link(fXOR(1, WPrevLink)) = backArc
.g(w).Link(WPrevLink) = backArc

.extFace(RootVertex).Link(RootSide) = w
.extFace(w).Link(WPrevLink) = RootVertex

End With
End Sub

Function VertexActiveStatus(theEmbedding As cPMFG_Graph, theVertex As Long, i As Long) As String
Dim leastLowpoint As Long, DFSChild As Long
With theEmbedding

DFSChild = .Vertices(theVertex).separatedDFSChildList
If DFSChild = -1 Then
    leastLowpoint = theVertex
Else
    leastLowpoint = .Vertices(DFSChild).Lowpoint
End If

If leastLowpoint > .Vertices(theVertex).leastAncestor Then
    leastLowpoint = .Vertices(theVertex).leastAncestor
End If

If leastLowpoint < i Then
    VertexActiveStatus = "VAS_EXTERNAL"
    Exit Function
End If

If .Vertices(theVertex).adjacentTo <> -1 Or _
    .Vertices(theVertex).pertinentBicompList <> -1 Then
    VertexActiveStatus = "VAS_INTERNAL"
    Exit Function
End If

VertexActiveStatus = "VAS_INACTIVE"

End With
End Function


Sub InvertVertex(theEmbedding As cPMFG_Graph, v As Long)
Dim j As Long, jtemp As Long
With theEmbedding

j = v
Do
    jtemp = .g(j).Link(0)
    .g(j).Link(0) = .g(j).Link(1)
    .g(j).Link(1) = jtemp
    j = .g(j).Link(0)
Loop While j >= (2 * .N_Nodes)

jtemp = .extFace(v).Link(0)
.extFace(v).Link(0) = .extFace(v).Link(1)
.extFace(v).Link(1) = jtemp

End With
End Sub


Sub SetSignOfChildEdge(theEmbedding As cPMFG_Graph, v As Long, sign As Long)
Dim j As Long
With theEmbedding

j = .g(v).Link(0)
Do While (j >= 2 * .N_Nodes)
    If .g(j).etype = "EDGE_DFSCHILD" Then
        .g(j).sign = sign
        Exit Do
    End If
    j = .g(j).Link(0)
Loop

End With
End Sub


Sub MergeVertex(theEmbedding As cPMFG_Graph, w As Long, WPrevLink As Long, r As Long)
Dim j As Long, jtwin As Long, n As Long
Dim e_w As Long, e_r As Long, e_ext As Long
With theEmbedding

n = .N_Nodes

j = .g(r).Link(0)
Do While j >= 2 * n
    jtwin = gp.GetTwinArc(theEmbedding, j)
    .g(jtwin).index = w
    j = .g(j).Link(0)
Loop

e_w = .g(w).Link(WPrevLink)
e_r = .g(r).Link(fXOR(1, WPrevLink))
e_ext = .g(r).Link(WPrevLink)

.g(e_w).Link(fXOR(1, WPrevLink)) = e_r
.g(e_r).Link(WPrevLink) = e_w

.g(w).Link(WPrevLink) = e_ext
.g(e_ext).Link(fXOR(1, WPrevLink)) = w

.g(r).InitGraphNode


End With
End Sub


Sub MergeBicomps(theEmbedding As cPMFG_Graph)
Dim i As Long
Dim r As Long, Rout As Long, z As Long, ZPrevLink As Long
Dim theList As Long, DFSChild As Long, RootId As Long
Dim extFaceVertex As Long
With theEmbedding

Do While .Stack.Top > 0

    Call .Stack.Pop2(r, Rout)
    Call .Stack.Pop2(z, ZPrevLink)

    extFaceVertex = .extFace(r).Link(fXOR(1, Rout))
    .extFace(z).Link(ZPrevLink) = extFaceVertex
    
    If .extFace(extFaceVertex).Link(0) = .extFace(extFaceVertex).Link(1) Then
        .extFace(extFaceVertex).Link(fXOR(Rout, .extFace(extFaceVertex).inversionFlag)) = z
    Else
        If .extFace(extFaceVertex).Link(0) = r Then
        i = 0
        Else
        i = 1
        End If
        .extFace(extFaceVertex).Link(i) = z
    End If
    
    If ZPrevLink = Rout Then
        If .g(r).Link(0) <> .g(r).Link(1) Then
            Call InvertVertex(theEmbedding, r)
        End If
        Call SetSignOfChildEdge(theEmbedding, r, -1)
        Rout = fXOR(1, ZPrevLink)
    End If
    
    RootId = r - .N_Nodes
    theList = .Vertices(z).pertinentBicompList
    theList = .BicompLists.LCDelete(theList, RootId)
    .Vertices(z).pertinentBicompList = theList
    
    DFSChild = r - .N_Nodes
    theList = .Vertices(z).separatedDFSChildList
    theList = .DFSChildLists.LCDelete(theList, DFSChild)
    .Vertices(z).separatedDFSChildList = theList
    
    Call MergeVertex(theEmbedding, z, ZPrevLink, r)
    
Loop

End With
End Sub


Sub RecordPertinentChildBicomp(theEmbedding As cPMFG_Graph, i As Long, RootVertex As Long)
Dim parentCopy As Long, DFSChild As Long, RootId As Long, BicompList As Long
With theEmbedding

RootId = RootVertex - .N_Nodes
DFSChild = RootId
parentCopy = .Vertices(DFSChild).DFSParent

BicompList = .Vertices(parentCopy).pertinentBicompList

If .Vertices(DFSChild).Lowpoint < i Then
    BicompList = .BicompLists.LCAppend(BicompList, RootId)
Else
    BicompList = .BicompLists.LCPrepend(BicompList, RootId)
End If

.Vertices(parentCopy).pertinentBicompList = BicompList

End With
End Sub


Function GetPertinentChildBicomp(theEmbedding As cPMFG_Graph, w As Long) As Long
Dim RootId As Long
With theEmbedding

RootId = .Vertices(w).pertinentBicompList
If RootId = -1 Then
    GetPertinentChildBicomp = -1
    Exit Function
End If
GetPertinentChildBicomp = RootId + .N_Nodes

End With
End Function

Sub Walkup(theEmbedding As cPMFG_Graph, i As Long, w As Long)
Dim Zig As Long, Zag As Long, ZigPrevLink As Long, ZagPrevLink As Long
Dim n As Long, r As Long, parentCopy As Long
Dim nextVertex As Long
With theEmbedding

n = .N_Nodes
Zig = w
Zag = w
ZigPrevLink = 1
ZagPrevLink = 0

Do While Zig <> i

    If .g(Zig).visited = i Then Exit Do
    If .g(Zag).visited = i Then Exit Do

    .g(Zig).visited = i
    .g(Zag).visited = i
    
    If Zig >= n Then
        r = Zig
    ElseIf Zag >= n Then
        r = Zag
    Else
        r = -1
    End If

    If r <> -1 Then
        parentCopy = .Vertices(r - n).DFSParent
        If parentCopy <> i Then Call RecordPertinentChildBicomp(theEmbedding, i, r)
        Zig = parentCopy
        Zag = parentCopy
        ZigPrevLink = 1
        ZagPrevLink = 0
    Else
        nextVertex = .extFace(Zig).Link(fXOR(1, ZigPrevLink))
        If .extFace(nextVertex).Link(0) = Zig Then
            ZigPrevLink = 0
        Else
            ZigPrevLink = 1
        End If
        Zig = nextVertex
        
        nextVertex = .extFace(Zag).Link(fXOR(1, ZagPrevLink))
        If .extFace(nextVertex).Link(0) = Zag Then
            ZagPrevLink = 0
        Else
            ZagPrevLink = 1
        End If
        Zag = nextVertex
    End If

Loop

End With
End Sub

Sub WalkDown(theEmbedding As cPMFG_Graph, i As Long, RootVertex As Long)
Dim w As Long, WPrevLink As Long, r As Long, Rout As Long
Dim x As Long, XPrevLink As Long, y As Long, YPrevLink As Long
Dim RootSide As Long, RootEdgeChild As Long
With theEmbedding

RootEdgeChild = RootVertex - .N_Nodes

Call .Stack.ClearStack

For RootSide = 0 To 1

    WPrevLink = fXOR(1, RootSide)
        
    w = .extFace(RootVertex).Link(RootSide)
    
    Do While w <> RootVertex
        
        If .Vertices(w).adjacentTo <> -1 Then
            Call MergeBicomps(theEmbedding)
            Call EmbedBackEdgeToDescendant(theEmbedding, RootSide, RootVertex, w, WPrevLink)
            .Vertices(w).adjacentTo = -1
        End If
        
        If .Vertices(w).pertinentBicompList <> -1 Then
            Call .Stack.Push2(w, WPrevLink)
            r = GetPertinentChildBicomp(theEmbedding, w)
            
            x = .extFace(r).Link(0)
            If .extFace(x).Link(1) = r Then
                XPrevLink = 1
            Else
                XPrevLink = 0
            End If
            y = .extFace(r).Link(1)
            If .extFace(y).Link(0) = r Then
                YPrevLink = 0
            Else
                YPrevLink = 1
            End If
            
            If x = y And .extFace(x).inversionFlag > 0 Then
                XPrevLink = 0
                YPrevLink = 1
            End If
            
            If VertexActiveStatus(theEmbedding, x, i) = "VAS_INTERNAL" Then
                w = x
            ElseIf VertexActiveStatus(theEmbedding, y, i) = "VAS_INTERNAL" Then
                w = y
            ElseIf PERTINENT(theEmbedding, x, i) = 1 Then
                w = x
            Else
                w = y
            End If
            
            If w = x Then
                WPrevLink = XPrevLink
            Else
                WPrevLink = YPrevLink
            End If
            
            If w = x Then
                Rout = 0
            Else
                Rout = 1
            End If
            
            Call .Stack.Push2(r, Rout)
        
        ElseIf VertexActiveStatus(theEmbedding, w, i) = "VAS_INACTIVE" Then
        
            x = .extFace(w).Link(fXOR(1, WPrevLink))
            If .extFace(x).Link(0) = w Then
                WPrevLink = 0
            Else
                WPrevLink = 1
            End If
            w = x
        
        Else
        
            Exit Do
            
        End If
            
    Loop
    
    If .Stack.Top = 0 Then
        
        .extFace(RootVertex).Link(RootSide) = w
        .extFace(w).Link(WPrevLink) = RootVertex
        
        If .extFace(w).Link(0) = .extFace(w).Link(1) And WPrevLink = RootSide Then
            .extFace(w).inversionFlag = 1
        Else
            .extFace(w).inversionFlag = 0
        End If
    
    End If
    
    If .Stack.Top > 0 Or w = RootVertex Then Exit For

Next RootSide

End With
End Sub


Function PERTINENT(theEmbedding As cPMFG_Graph, theVertex As Long, i As Long)
With theEmbedding
If .Vertices(theVertex).adjacentTo <> -1 Or _
    .Vertices(theVertex).pertinentBicompList <> -1 Then
    PERTINENT = 1
Else
    PERTINENT = 0
End If
End With
End Function

Function fMIN2(i As Long, j As Long) As Long
If i < j Then
    fMIN2 = i
Else
    fMIN2 = j
End If
End Function

Function fMAX2(i As Long, j As Long) As Long
If i > j Then
    fMIN2 = i
Else
    fMIN2 = j
End If
End Function

Function fMIN3(i As Long, j As Long, k As Long) As Long
fMIN3 = fMIN2(fMIN2(i, j), fMIN2(j, k))
End Function

Function fMAX3(i As Long, j As Long, k As Long) As Long
fMAX3 = fMAX2(fMAX2(i, j), fMAX2(j, k))
End Function

Function fXOR(i As Long, j As Long) As Long
    If i = j Then
        fXOR = 0
    ElseIf i > 0 Or j > 0 Then
        fXOR = 1
    End If
End Function
