Attribute VB_Name = "gp"
Option Explicit


Sub ReadAdjMatrix(theGraph As cPMFG_Graph, AdjM() As Long, n As Long)
Dim i As Long, j As Long, n_raw As Long
Dim u As Long, v As Long
Set theGraph = New cPMFG_Graph
theGraph.InitGraph (n)
For i = 1 To UBound(AdjM, 2)
    u = AdjM(1, i)
    v = AdjM(2, i)
    Call gp.AddEdge(theGraph, u, 0, v, 0)
Next i
End Sub

Function WriteAdjList(theGraph As cPMFG_Graph) As String
Dim i As Long, j As Long, n As Long
Dim strtemp As String
WriteAdjList = ""
n = theGraph.N_Nodes
For i = 0 To n - 1
    strtemp = i & ":"
    j = theGraph.g(i).Link(1)
    Do While j >= n
        strtemp = strtemp & " " & theGraph.g(j).index
        j = theGraph.g(j).Link(1)
    Loop
    WriteAdjList = WriteAdjList & vbCrLf & strtemp
Next i
End Function

Function IsNeighbor(theGraph As cPMFG_Graph, u As Long, v As Long)
Dim j As Long, n As Long
IsNeighbor = 0
With theGraph
    n = .N_Nodes
    j = .g(u).Link(0)
    Do While j >= 2 * n
        If .g(j).index = v Then
            IsNeighbor = 1
            Exit Do
        End If
        j = .g(j).Link(0)
    Loop
End With
End Function

Function GetVertexDegree(theGraph As cPMFG_Graph, v As Long) As Long
Dim j As Long, degree As Long, n As Long
n = theGraph.N_Nodes
degree = 0
With theGraph
    j = .g(v).Link(0)
    Do While j >= 2 * n
        degree = degree + 1
        j = theGraph.g(j).Link(0)
    Loop
End With
GetVertexDegree = degree
End Function

Sub AddArc(theGraph As cPMFG_Graph, u As Long, v As Long, arcPos As Long, iLink As Long)
Dim u0 As Long
With theGraph
    .g(arcPos).index = v
    If .g(u).Link(0) = -1 Then
        .g(u).Link(1) = arcPos
        .g(u).Link(0) = arcPos
        .g(arcPos).Link(1) = u
        .g(arcPos).Link(0) = u
    Else
        u0 = .g(u).Link(iLink)
        .g(arcPos).Link(iLink) = u0
        .g(arcPos).Link(fXOR(1, iLink)) = u
        .g(u).Link(iLink) = arcPos
        .g(u0).Link(fXOR(1, iLink)) = arcPos
    End If
End With
End Sub


Sub AddEdge(theGraph As cPMFG_Graph, u As Long, ulink As Long, v As Long, vlink As Long)
Dim upos As Long, vpos As Long
vpos = 2 * theGraph.N_Nodes + 2 * theGraph.M_Edges
upos = gp.GetTwinArc(theGraph, vpos)
Call gp.AddArc(theGraph, u, v, vpos, ulink)
Call gp.AddArc(theGraph, v, u, upos, vlink)
theGraph.M_Edges = theGraph.M_Edges + 1
End Sub

Function GetTwinArc(graph As cPMFG_Graph, arc As Long) As Long
If arc Mod 2 = 1 Then
    GetTwinArc = arc - 1
ElseIf arc Mod 2 = 0 Then
    GetTwinArc = arc + 1
End If
End Function

Sub FillVisitedFlags(theGraph As cPMFG_Graph, FillValue As Long)
Dim i As Long, limit As Long
limit = 2 * (theGraph.N_Nodes + theGraph.M_Edges)
For i = 0 To limit - 1
    theGraph.g(i).visited = FillValue
Next i
End Sub
