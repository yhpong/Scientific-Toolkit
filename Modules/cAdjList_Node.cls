VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdjList_Node"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pPtr As Long
Private pWgt As Double
Private pNxt As cAdjList_Node

Sub Init(v As Long, Optional Wgt As Double = 1)
    pPtr = v
    pWgt = Wgt
    Set pNxt = Nothing
End Sub

Public Property Set Nxt(v As cAdjList_Node)
    Set pNxt = v
End Property

Public Property Get Nxt() As cAdjList_Node
    Set Nxt = pNxt
End Property

Public Property Get Ptr() As Long
    Ptr = pPtr
End Property

Public Property Let Ptr(v As Long)
    pPtr = v
End Property

Public Property Get Wgt() As Double
    Wgt = pWgt
End Property

Public Property Let Wgt(x As Double)
    pWgt = x
End Property

