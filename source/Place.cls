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

Public Content As ShapeRange
Public CropBox As ShapeRange
Public IsEmpty As Boolean
Public IsFront As Boolean
Public Name As String
Public Shape As Shape
Public ShapesByTags As Dictionary
Public ToLayer1 As ShapeRange
Public ToLayer2 As ShapeRange
Public ToLayer3 As ShapeRange
Public ToCropBoxLayer As ShapeRange

Public Property Get AsComposerElement() As IComposerElement
    Set AsComposerElement = Me
End Property

Public Sub Parse(ByVal Cfg As Config, ByVal Table As Dictionary)
    If Not Shape.Type = cdrGroupShape Then Throw "тут должна быть группа"
    Set ShapesByTags = New Dictionary
    Dim Shapes As ShapeRange
    Dim Tag As Variant
    
    For Each Tag In Table.Keys
        Set Shapes = FindShapesWithText(ActivePage.Shapes.All, Tag)
        If Shapes.Count > 0 Then ShapesByTags.Add Tag, Shapes
    Next Tag
    
    Set ToCropBoxLayer = _
            FindByOutlineColor( _
                Shape.Shapes.All, _
                Cfg.CropBoxOutlineColor _
            )
    If ToCropBoxLayer.Count = 0 Then
        Set CropBox = Shape.Shapes.All
    Else
        Set CropBox = ToCropBoxLayer
    End If
    
    Set Content = _
        FilterShapesWithTags(ShapesByTags, Cfg.ContentTagsCommaSeparated)
    Set ToLayer1 = _
        FilterShapesWithTags(ShapesByTags, Cfg.Layer1TagsCommaSeparated)
    Set ToLayer2 = _
        FilterShapesWithTags(ShapesByTags, Cfg.Layer2TagsCommaSeparated)
    Set ToLayer3 = _
        FilterShapesWithTags(ShapesByTags, Cfg.Layer3TagsCommaSeparated)
    
End Sub

Private Function FindByOutlineColor( _
                     ByVal Source As ShapeRange, _
                     ByVal ColorString As String _
                 ) As ShapeRange
    Dim Color As Color
    Set Color = CreateColor(ColorString)
    Set FindByOutlineColor = CreateShapeRange
    Dim Shape As Shape
    For Each Shape In Source.Shapes.FindShapes
        If ShapeHasOutline(Shape) Then
            If Shape.Outline.Color.IsSame(Color) Then _
                FindByOutlineColor.Add Shape
        End If
    Next Shape
End Function

Private Function FilterShapesWithTags( _
                     ByVal ShapesByTags As Dictionary, _
                     ByVal TagsCommaSeparated As String _
                 ) As ShapeRange
    Set FilterShapesWithTags = CreateShapeRange
    Dim Tag As Variant
    For Each Tag In VBA.Split(TagsCommaSeparated, ",")
        If ShapesByTags.Exists(Tag) Then _
            FilterShapesWithTags.AddRange ShapesByTags(Tag)
    Next Tag
End Function

Public Property Get Self() As Place
    Set Self = Me
End Property

'===============================================================================

Private Property Get IComposerElement_Height() As Double
    IComposerElement_Height = CropBox.SizeHeight
End Property

Private Property Let IComposerElement_PivotX(ByVal RHS As Double)
    Shape.LeftX = RHS + (Shape.LeftX - CropBox.LeftX)
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
    Set IComposerElement_Shapes = PackShapes(Shape)
End Property

Private Property Get IComposerElement_Width() As Double
    IComposerElement_Width = CropBox.SizeWidth
End Property
