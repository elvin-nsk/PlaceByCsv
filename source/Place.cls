VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Place"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Struct
Option Explicit

Implements IComposerElement

'===============================================================================

Public CropBox As Shape
Public IsFront As Boolean
Public Name As String
Public Shape As Shape

Public Property Get AsComposerElement() As IComposerElement
    Set AsComposerElement = Me
End Property

'===============================================================================

Private Property Get IComposerElement_Height() As Double
    IComposerElement_Height = Shape.SizeHeight
End Property

Private Property Let IComposerElement_PivotX(ByVal RHS As Double)
    Shape.LeftX = RHS + (CropBox.LeftX - Shape.LeftX)
End Property
Private Property Get IComposerElement_PivotX() As Double
    IComposerElement_PivotX = CropBox.LeftX
End Property

Private Property Let IComposerElement_PivotY(ByVal RHS As Double)
    Shape.TopY = RHS + (Shape.TopY - CropBox.TopY)
End Property
Private Property Get IComposerElement_PivotY() As Double
    IComposerElement_PivotY = CropBox.TopY
End Property

Private Property Get IComposerElement_Self() As IComposerElement
    Set IComposerElement_Self = Me
End Property

Private Property Get IComposerElement_Shapes() As ShapeRange
    Set IComposerElement_Shapes = CreateShapeRange
    IComposerElement_Shapes.Add Shape
End Property

Private Property Get IComposerElement_Width() As Double
    IComposerElement_Width = Shape.SizeWidth
End Property
