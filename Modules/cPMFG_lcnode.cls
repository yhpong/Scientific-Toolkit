VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPMFG_lcnode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pprev As Long
Private pfwd As Long

Public Property Get prev() As Long
prev = pprev
End Property

Public Property Let prev(v As Long)
pprev = v
End Property


Public Property Get fwd() As Long
fwd = pfwd
End Property

Public Property Let fwd(v As Long)
pfwd = v
End Property

