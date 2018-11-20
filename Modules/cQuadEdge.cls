VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cQuadEdge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pid As Long
Private pidx As Long
Private pOrg As Long, pDest As Long, pLeft As Long, pRight As Long
Private pNext As cQuadEdge
Private pQuad As cQuadEdge_Quad
Private pisVisited As Boolean

Public Sub Init()
Dim i As Long
Dim tmp_edge As cQuadEdge
    
    pidx = 0
    Set pQuad = New cQuadEdge_Quad
    Call pQuad.Init
    Call pQuad.Quad_Add(0, Me)
    
    'Add the duals the quad list
    For i = 1 To 3
        Set tmp_edge = New cQuadEdge
        tmp_edge.idx = i
        Set tmp_edge.Quad = pQuad
        Call pQuad.Quad_Add(i, tmp_edge)
    Next i

    Set pQuad.Edge(0).Nxt = pQuad.Edge(0)
    Set pQuad.Edge(1).Nxt = pQuad.Edge(3)
    Set pQuad.Edge(2).Nxt = pQuad.Edge(2)
    Set pQuad.Edge(3).Nxt = pQuad.Edge(1)

End Sub

Public Property Let isVisited2(x As Boolean)
    pisVisited = x
End Property

Public Property Let isVisited(x As Boolean)
    pisVisited = x
    Me.Sym.isVisited2 = x
    Me.Rot.isVisited2 = x
    Me.InvRot.isVisited2 = x
End Property

Public Property Get isVisited() As Boolean
    isVisited = pisVisited
End Property

Public Property Set Quad(A As cQuadEdge_Quad)
    Set pQuad = A
End Property

Public Property Set Nxt(A As cQuadEdge)
    Set pNext = A
End Property

Public Property Let id(id As Long)
    pid = id
End Property

Public Property Get id() As Long
    id = pid
End Property

Public Property Let idx(i As Long)
    pidx = i
End Property

Public Property Get idx() As Long
    idx = pidx
End Property

Public Property Let Org2(v As Long)
    pOrg = v
End Property

Public Property Let Org(v As Long)
    pOrg = v
    Me.Sym.Dest2 = v
End Property

Public Property Get Org() As Long
    Org = pOrg
End Property

Public Property Let Dest2(v As Long)
    pDest = v
End Property

Public Property Let Dest(v As Long)
    pDest = v
    Me.Sym.Org2 = v
End Property

Public Property Get Dest() As Long
    Dest = pDest
End Property

Public Property Let setLeft(v As Long)
    pLeft = v
End Property

Public Property Get Left() As Long
    Left = pLeft
End Property

Public Property Let setRight(v As Long)
    pRight = v
End Property

Public Property Get Right() As Long
    Right = pRight
End Property

Public Function Rot() As cQuadEdge
Dim i As Long
    If pidx < 3 Then
        i = pidx + 1
    Else
        i = pidx - 3
    End If
    Set Rot = pQuad.Edge(i)
End Function

Public Function InvRot() As cQuadEdge
Dim i As Long
    If pidx > 0 Then
        i = pidx - 1
    Else
        i = pidx + 3
    End If
    Set InvRot = pQuad.Edge(i)
End Function

Public Function Sym() As cQuadEdge
Dim i As Long
    If pidx < 2 Then
        i = pidx + 2
    Else
        i = pidx - 2
    End If
    Set Sym = pQuad.Edge(i)
End Function

Public Function Onext() As cQuadEdge
    Set Onext = pNext
End Function

Public Function Oprev() As cQuadEdge
    Set Oprev = Me.Rot.Onext.Rot
End Function

Public Function Dnext() As cQuadEdge
    Set Dnext = Me.Sym.Onext.Sym
End Function

Public Function Dprev() As cQuadEdge
    Set Dprev = Me.InvRot.Onext.InvRot
End Function

Public Function Lnext() As cQuadEdge
    Set Lnext = Me.InvRot.Onext.Rot
End Function

Public Function Lprev() As cQuadEdge
    Set Lprev = Me.Onext.Sym
End Function

Public Function Rnext() As cQuadEdge
    Set Rnext = Me.Rot.Onext.InvRot
End Function

Public Function Rprev() As cQuadEdge
    Set Rprev = Me.Sym.Onext
End Function
