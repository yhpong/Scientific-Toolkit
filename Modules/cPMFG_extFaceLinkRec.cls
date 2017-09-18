VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPMFG_extFaceLinkRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private plink() As Long
Private pinversionFlag As Long

Sub Class_Initialize()
ReDim plink(0 To 1)
End Sub

Public Property Get Link(i As Long) As Long
Link = plink(i)
End Property

Public Property Let Link(i As Long, llink As Long)
plink(i) = llink
End Property

Public Property Get inversionFlag() As Long
inversionFlag = pinversionFlag
End Property

Public Property Let inversionFlag(i As Long)
pinversionFlag = i
End Property

