VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPMFG_VertexRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pleastAncestor As Long
Private pLowpoint As Long
Private pDFSParent As Long
Private padjacentTo As Long
Private ppertinentBicompList As Long
Private pseparatedDFSChildList As Long
Private pfwdArcList As Long


Sub InitVertexRec(i As Long)
    pleastAncestor = 0
    pLowpoint = i
    pDFSParent = -1
    padjacentTo = -1
    ppertinentBicompList = -1
    pseparatedDFSChildList = -1
    pfwdArcList = -1
End Sub

Public Property Get DFSParent() As Long
DFSParent = pDFSParent
End Property

Public Property Let DFSParent(v As Long)
pDFSParent = v
End Property

Public Property Get leastAncestor() As Long
leastAncestor = pleastAncestor
End Property

Public Property Let leastAncestor(v As Long)
pleastAncestor = v
End Property


Public Property Get Lowpoint() As Long
Lowpoint = pLowpoint
End Property

Public Property Let Lowpoint(v As Long)
pLowpoint = v
End Property


Public Property Get separatedDFSChildList() As Long
separatedDFSChildList = pseparatedDFSChildList
End Property

Public Property Let separatedDFSChildList(i As Long)
pseparatedDFSChildList = i
End Property


Public Property Get fwdArcList() As Long
fwdArcList = pfwdArcList
End Property

Public Property Let fwdArcList(i As Long)
pfwdArcList = i
End Property

Public Property Get adjacentTo() As Long
adjacentTo = padjacentTo
End Property

Public Property Let adjacentTo(v As Long)
padjacentTo = v
End Property


Public Property Get pertinentBicompList() As Long
pertinentBicompList = ppertinentBicompList
End Property

Public Property Let pertinentBicompList(v As Long)
ppertinentBicompList = v
End Property

