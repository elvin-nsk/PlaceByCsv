VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEither"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Property Get IsLeft() As Boolean
End Property

Public Property Get IsRight() As Boolean
End Property

Public Property Get Left() As Variant
End Property

'@DefaultMember
Public Property Get Right() As Variant
Attribute Right.VB_UserMemId = 0
End Property

'===============================================================================

Public Property Get Self() As IEither
End Property
