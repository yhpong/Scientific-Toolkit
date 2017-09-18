VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPMFG_Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pindex As Long
Private pvisited As Long
Private plink() As Long
Private ptype As String
Private psign As Long

Sub InitGraphNode()
    ReDim plink(0 To 1)
    pindex = -1
    plink(0) = -1
    plink(1) = -1
    pvisited = 0
    ptype = "TYPE_UNKNOWN"
    psign = 1
End Sub


Public Property Get index() As Long
index = pindex
End Property

Public Property Let index(v As Long)
pindex = v
End Property


Public Property Get visited() As Long
visited = pvisited
End Property

Public Property Let visited(v As Long)
pvisited = v
End Property


Public Property Get Link(i As Long) As Long
Link = plink(i)
End Property

Public Property Let Link(i As Long, llink As Long)
plink(i) = llink
End Property


Public Property Get etype() As String
etype = ptype
End Property

Public Property Let etype(ltype As String)
ptype = ltype
End Property

Public Property Get sign() As Long
sign = psign
End Property

Public Property Let sign(lsign As Long)
psign = lsign
End Property
