VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Either"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements IEither

'===============================================================================
' # Declarations

Private Type typeThis
    Left As Variant
    Right As Variant
End Type
Private This As typeThis

'===============================================================================
' # Constructors

Public Function SetLeft( _
                    ByVal Left As Variant _
                ) As IEither
    With New Either
        Set SetLeft = .Init(Left:=Left)
    End With
End Function

Public Function SetRight( _
                    ByVal Right As Variant _
                ) As IEither
    With New Either
        Set SetRight = .Init(Right:=Right)
    End With
End Function

Friend Function Init( _
                    Optional ByVal Left As Variant, _
                    Optional ByVal Right As Variant _
                ) As IEither
    With This
        Assign .Left, Left
        Assign .Right, Right
    End With
    Set Init = Me
End Function

'===============================================================================
' # Public

Private Property Get IEither_IsLeft() As Boolean
    IEither_IsLeft = Not VBA.IsMissing(This.Left)
End Property

Private Property Get IEither_IsRight() As Boolean
    IEither_IsRight = Not VBA.IsMissing(This.Right)
End Property

Private Property Get IEither_Left() As Variant
    Assign IEither_Left, This.Left
End Property

Private Property Get IEither_Right() As Variant
    Assign IEither_Right, This.Right
End Property

Private Property Get IEither_Self() As IEither
    Set IEither_Self = Me
End Property

'===============================================================================
' # Helpers

Public Sub Assign(ByRef Variable As Variant, ByRef Value As Variant)
    If VBA.IsObject(Value) Then
        Set Variable = Value
    Else
        Variable = Value
    End If
End Sub
