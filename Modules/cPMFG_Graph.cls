VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPMFG_Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pN_Nodes As Long
Private pM_Edges As Long
Private pG() As cPMFG_Node
Private pVertices() As cPMFG_VertexRec
Private pBicompLists As cPMFG_ListColl
Private pDFSChildLists As cPMFG_ListColl
Private ptheStack As cPMFG_Stack
Private pbuckets() As Long
Private pbin As cPMFG_ListColl
Private pextFace() As cPMFG_extFaceLinkRec
Private pinternalFlags As Long

'FLAGS_DFSNUMBERED 1
'FLAGS_SORTEDBYDFI 2

Sub InitGraph(n As Long)
Dim i As Long, v As cPMFG_Node, vRec As cPMFG_VertexRec, vextFace As cPMFG_extFaceLinkRec
    pN_Nodes = n
    
    ReDim pG(0 To (2 * n - 1 + n * (n - 1)))
    For i = 0 To (2 * n - 1 + n * (n - 1))
        Set v = New cPMFG_Node
        v.InitGraphNode
        Set pG(i) = v
    Next i
    
    ReDim pVertices(0 To n - 1)
    For i = 0 To (n - 1)
        Set vRec = New cPMFG_VertexRec
        vRec.InitVertexRec (i)
        Set pVertices(i) = vRec
    Next i
    
    ReDim pextFace(0 To 2 * n - 1)
    For i = 0 To 2 * n - 1
        Set vextFace = New cPMFG_extFaceLinkRec
        vextFace.Link(0) = -1
        vextFace.Link(1) = -1
        vextFace.inversionFlag = 0
        Set pextFace(i) = vextFace
    Next i
    
    Set pbin = New cPMFG_ListColl
    Set pDFSChildLists = New cPMFG_ListColl
    Set pBicompLists = New cPMFG_ListColl
    
    pbin.LCNew (n)
    pDFSChildLists.LCNew (n)
    pBicompLists.LCNew (n)
    
    ReDim pbuckets(0 To n - 1)
    
    Set ptheStack = New cPMFG_Stack
    ptheStack.NewStack (n * (n - 1))
    ptheStack.NewStack (n * (n - 1))
    pinternalFlags = 0
    
End Sub

Public Property Get N_Nodes() As Long
N_Nodes = pN_Nodes
End Property

Public Property Let N_Nodes(lN_Nodes As Long)
pN_Nodes = lN_Nodes
End Property

Public Property Get M_Edges() As Long
M_Edges = pM_Edges
End Property

Public Property Let M_Edges(lM_Edges As Long)
pM_Edges = lM_Edges
End Property

Public Property Get g(i As Long) As cPMFG_Node
Set g = pG(i)
End Property

Public Property Set g(i As Long, j As cPMFG_Node)
Set pG(i) = j
End Property

Public Property Get Vertices(i As Long) As cPMFG_VertexRec
Set Vertices = pVertices(i)
End Property

Public Property Set Vertices(i As Long, j As cPMFG_VertexRec)
Set pVertices(i) = j
End Property

Public Property Get bin() As cPMFG_ListColl
Set bin = pbin
End Property

Public Property Get DFSChildLists() As cPMFG_ListColl
Set DFSChildLists = pDFSChildLists
End Property


Public Property Get BicompLists() As cPMFG_ListColl
Set BicompLists = pBicompLists
End Property


Public Property Get extFace(i As Long) As cPMFG_extFaceLinkRec
Set extFace = pextFace(i)
End Property

Public Property Set extFace(i As Long, j As cPMFG_extFaceLinkRec)
Set pextFace(i) = j
End Property


Public Property Get internalFlags() As Long
internalFlags = pinternalFlags
End Property

Public Property Let internalFlags(linternalFlags As Long)
pinternalFlags = linternalFlags
End Property


Public Property Get Stack() As cPMFG_Stack
Set Stack = New cPMFG_Stack
Set Stack = ptheStack
End Property
