VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cqtree_point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pindex As Long
Private pxy(1 To 2) As Double

Sub Init(x As Double, y As Double, index As Long)
pxy(1) = x
pxy(2) = y
pindex = index
End Sub

Public Property Get xy(i As Long) As Double
xy = pxy(i)
End Property

Public Property Get index() As Long
index = pindex
End Property
