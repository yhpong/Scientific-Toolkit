VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cQuadEdge_Quad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pEdges() As cQuadEdge

Sub Init()
    ReDim pEdges(0 To 3)
End Sub

Sub Quad_Add(i As Long, A_edge As cQuadEdge)
    Set pEdges(i) = A_edge
End Sub

Public Property Get Edge(k As Long) As cQuadEdge
    Set Edge = pEdges(k)
End Property
