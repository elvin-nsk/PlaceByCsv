Attribute VB_Name = "PlaceByTable"
'===============================================================================
'   Макрос          : PlaceByTable
'   Версия          : 2023.01.05
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "PlaceByTable"

'===============================================================================

Private Type tTargetLayers
    Layer1 As Layer
    Layer2 As Layer
    Layer3 As Layer
    CropBoxLayer As Layer
End Type

Private Const ImportExt As String = "cdr"

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    Dim Cfg As Config
    Set Cfg = Config.Bind
    
    Dim Table As Dictionary
    Set Table = ParseTable(GetTable(Cfg.CsvFile, Cfg.CsvSeparator))
    
    Dim Imposition As Document
    Set Imposition = CreateDocument
    Imposition.Unit = cdrMillimeter
    
    BoostStart APP_NAME, RELEASE
    
    Dim Groups As Collection
    Set Groups = ProcessTableAsGroups(Table, Cfg)
    ComposeGroupsAndReturnGroupsAsElements Groups, Cfg
    ComposeGroups GroupsToElements(Groups), Cfg
    Dim TargetLayers As tTargetLayers
    With ActivePage
        .Shapes.All.SetPositionEx _
            cdrCenter, .CenterX, .CenterY
        .SetSize .Shapes.All.SizeWidth + (Cfg.GroupsMinDistance * 2), _
                 .Shapes.All.SizeHeight + (Cfg.GroupsMinDistance * 2)
        ActiveLayer.Name = Cfg.DefaultLayerName
        Set TargetLayers.Layer1 = .CreateLayer(Cfg.Layer1Name)
        Set TargetLayers.Layer2 = .CreateLayer(Cfg.Layer2Name)
        Set TargetLayers.Layer3 = .CreateLayer(Cfg.Layer3Name)
        Set TargetLayers.CropBoxLayer = .CreateLayer(Cfg.LayerCropBoxName)
        TargetLayers.CropBoxLayer.Color.CopyAssign _
            CreateColor(Cfg.CropBoxOutlineColor)
    End With
    SpreadGroupsToLayers Groups, TargetLayers
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================

Private Function ParseTable(ByVal Table As Variant) As Dictionary
    Dim Dic As New Dictionary
    Dim Row As Long
    Dim Column As Long
    Dim Key As String
    For Column = 1 To UBound(Table, 2)
        Key = Table(1, Column)
        Dic.Add Key, New Collection
        For Row = 2 To UBound(Table, 1)
            Dic(Key).Add Table(Row, Column)
        Next Row
    Next Column
    Set ParseTable = Dic
End Function

Private Function GetTable( _
                     ByVal CsvFile As String, _
                     ByVal CsvSeparator As String _
                 ) As Variant
    Dim Str As String
    Str = ReadFileAD(CsvFile)
    GetTable = _
        CsvUtils.Create(CsvSeparator) _
            .ParseCsvToArray(Str)
End Function

'-------------------------------------------------------------------------------

Private Function ProcessTableAsGroups( _
                     ByVal Table As Dictionary, _
                     ByVal Cfg As Config _
                 ) As Collection
    Dim TotalRows As Long
    TotalRows = Table(Cfg.FrontTag).Count
    If TotalRows < 1 Then Throw "Таблица пустая"
    Set ProcessTableAsGroups = New Collection
    Dim Side As Collection
    Dim Group As GroupTwoSides
    Dim CurrentPairIndex As Long
    Dim MaxPlacesPerSide As Long
    MaxPlacesPerSide = Cfg.MaxPlacesPerSideX * Cfg.MaxPlacesPerSideY
    Dim Row As Long
    Row = 0
    Dim LastCropBoxFace As Rect
    Dim LastCropBoxBack As Rect
    
    Do Until Row > TotalRows
        
        CurrentPairIndex = 1
        Set Group = New GroupTwoSides
        Set Group.Fronts = New Collection
        Set Group.Backs = New Collection
        
        Do Until CurrentPairIndex > MaxPlacesPerSide
            Row = Row + 1
            If Row <= TotalRows Then
                Group.Fronts.Add ProcessPlace(True, Row, Table, Cfg)
                Group.Backs.Add ProcessPlace(False, Row, Table, Cfg)
                Set LastCropBoxFace = _
                    Group.Fronts(Group.Fronts.Count). _
                    ToCropBoxLayer.BoundingBox.GetCopy
                Set LastCropBoxBack = _
                    Group.Backs(Group.Backs.Count). _
                    ToCropBoxLayer.BoundingBox.GetCopy
                CurrentPairIndex = CurrentPairIndex + 1
            Else
                If CurrentPairIndex <= MaxPlacesPerSide Then
                    Group.Fronts.Add CreateEmptyPlace(True, LastCropBoxFace)
                    Group.Backs.Add CreateEmptyPlace(False, LastCropBoxBack)
                    CurrentPairIndex = CurrentPairIndex + 1
                Else
                    Exit Do
                End If
            End If
        Loop
                        
        ProcessTableAsGroups.Add Group
        
    Loop
    
End Function

Private Function ProcessPlace( _
                     ByVal Front As Boolean, _
                     ByVal Row As Long, _
                     ByVal Table As Dictionary, _
                     ByVal Cfg As Config _
                 ) As Place
    Dim Tag As String
    If Front Then Tag = Cfg.FrontTag Else Tag = Cfg.BackTag
    Dim File As IFileSpec
    Set File = FileSpec.Create
    File.Path = Cfg.SourceFolder
    File.NameWithoutExt = Table(Tag)(Row)
    File.Ext = ImportExt
    Dim Shape As Shape
    With TryImportShape(File.ToString)
        If .IsRight Then
            Set Shape = .Right
            ' // для отладки
            If Not RELEASE Then
                With ActiveLayer.CreateRectangleRect(Shape.BoundingBox)
                    .OrderBackOf Shape.Shapes.Last
                    .Fill.ApplyUniformFill CreateCMYKColor(0, 100, 0, 0)
                    .Outline.SetNoOutline
                End With
            End If
            ' //
        Else
            Throw "Ошибка импорта файла из таблицы, строка " _
                & VBA.CStr(Row + 1) & ", столбец " & Tag
        End If
    End With
    
    With New Place
        Set .Shape = Shape
        .IsFront = Front
        .Name = Table(Tag)(Row)
        .Parse Cfg, Table
        
        ReplaceTextInTaggedShapes .ShapesByTags, Row, Table
        TuneContentSize .Content, .CropBox, Cfg.ContentMaxSizeMultiplierToCropBox
        
        Set ProcessPlace = .Self
    End With
    
End Function

Private Function CreateEmptyPlace( _
                     ByVal Front As Boolean, _
                     ByVal Size As Rect _
                 ) As Place
    With New Place
        Set .Shape = ActiveLayer.CreateRectangleRect(Size)
        .IsFront = Front
        .IsEmpty = True
        Set .CropBox = PackShapes(.Shape)
        Set CreateEmptyPlace = .Self
        ' // для отладки
        If RELEASE Then .Shape.Outline.SetNoOutline
        ' //
    End With
End Function

Private Function TryImportShape(ByVal File As String) As IEither
On Error GoTo Fail
    ActiveLayer.Import File, cdrCDR
    Set TryImportShape = Either.SetRight(ActiveShape)
    Exit Function
Fail:
    Set TryImportShape = Either.SetLeft("Ошибка импорта файла " & File)
End Function

Private Sub ReplaceTextInTaggedShapes( _
                ByVal ShapesByTags As Dictionary, _
                ByVal Row As Long, _
                ByVal Table As Dictionary _
            )
    Dim Tag As Variant
    Dim Shape As Shape
    For Each Tag In ShapesByTags.Keys
        For Each Shape In ShapesByTags(Tag)
            Shape.Text.Story.Text = _
                VBA.Replace( _
                    Shape.Text.Story.Text, Tag, _
                    Table(Tag)(Row), 1, , vbTextCompare _
                )
        Next Shape
    Next Tag
End Sub

Private Sub TuneContentSize( _
                ByVal Content As ShapeRange, _
                ByVal CropBox As ShapeRange, _
                ByVal Mult As Double _
            )
    Dim MaxWidth As Double
    MaxWidth = CropBox.SizeWidth * Mult
    Dim MaxHeight As Double
    MaxHeight = CropBox.SizeHeight * Mult
    If Content.SizeWidth > MaxWidth Then _
        Content.SetSizeEx Content.CenterX, Content.CenterY, Width:=MaxWidth
    If Content.SizeHeight > MaxHeight Then _
        Content.SetSizeEx Content.CenterX, Content.CenterY, Height:=MaxHeight
End Sub

'-------------------------------------------------------------------------------

Private Function ComposeGroupsAndReturnGroupsAsElements( _
                     ByVal Groups As Collection, _
                     ByVal Cfg As Config _
                 ) As Collection
    Set ComposeGroupsAndReturnGroupsAsElements = New Collection
    Dim FrontShapes As ShapeRange
    Dim BackShapes As ShapeRange
    Dim Group As GroupTwoSides
    For Each Group In Groups
        ComposeSide Group.Fronts, Cfg
        Set FrontShapes = GetShapesFromPlaces(Group.Fronts)
        ComposeGroupsAndReturnGroupsAsElements.Add _
            GetShapesFromPlaces(Group.Fronts)
        
        ComposeSide Group.Backs, Cfg
        Set BackShapes = GetShapesFromPlaces(Group.Backs)
        With BackShapes
            .TopY = FrontShapes.TopY
            .LeftX = FrontShapes.RightX + Cfg.SidesMinDistanceX
        End With
        ComposeGroupsAndReturnGroupsAsElements.Add _
            GetShapesFromPlaces(Group.Backs)
    Next Group
End Function

Private Sub ComposeSide( _
                ByVal Side As Collection, _
                ByVal Cfg As Config _
            )
    With Composer.CreateAndCompose( _
             Elements:=Side, _
             StartingPoint:=FreePoint.Create(0, 297), _
             MaxPlacesInWidth:=Cfg.MaxPlacesPerSideX, _
             MaxPlacesInHeight:=Cfg.MaxPlacesPerSideY, _
             HorizontalSpace:=Cfg.PlacesMinDistanceX, _
             VerticalSpace:=Cfg.PlacesMinDistanceY _
         )
    End With
End Sub

Private Function GetShapesFromPlaces(ByVal Places As Collection) As ShapeRange
    Set GetShapesFromPlaces = CreateShapeRange
    Dim Place As Place
    For Each Place In Places
        GetShapesFromPlaces.Add Place.Shape
    Next Place
End Function

'-------------------------------------------------------------------------------

Private Sub ComposeGroups( _
                ByVal Elements As Collection, _
                ByVal Cfg As Config _
            )
    With Composer.CreateAndCompose( _
             Elements:=Elements, _
             StartingPoint:=FreePoint.Create(0, 297), _
             MaxPlacesInWidth:=VBA.Int(VBA.Sqr(Elements.Count)), _
             HorizontalSpace:=Cfg.GroupsMinDistance, _
             VerticalSpace:=Cfg.GroupsMinDistance _
         )
    End With
End Sub

Private Function GroupsToElements(ByVal Groups As Collection) As Collection
    Set GroupsToElements = New Collection
    Dim Shapes As ShapeRange
    Dim Group As GroupTwoSides
    For Each Group In Groups
        Set Shapes = CreateShapeRange
        Shapes.AddRange GetShapesFromPlaces(Group.Fronts)
        Shapes.AddRange GetShapesFromPlaces(Group.Backs)
        GroupsToElements.Add _
            ComposerElement.Create(Shapes)
    Next Group
End Function

'-------------------------------------------------------------------------------

Private Sub SpreadGroupsToLayers( _
                ByVal Groups As Collection, _
                ByRef TargetLayers As tTargetLayers _
            )
    Dim Group As GroupTwoSides
    For Each Group In Groups
        SpreadPlacesToLayers Group.Fronts, TargetLayers
        SpreadPlacesToLayers Group.Backs, TargetLayers
    Next Group
End Sub

Private Sub SpreadPlacesToLayers( _
                ByVal Places As Collection, _
                ByRef TargetLayers As tTargetLayers _
            )
    Dim Place As Place
    For Each Place In Places
        If Place.IsEmpty Then
            MoveToLayer Place.Shape, TargetLayers.CropBoxLayer
        Else
            Place.Shape.Ungroup
            MoveToLayer Place.ToLayer1, TargetLayers.Layer1
            MoveToLayer Place.ToLayer2, TargetLayers.Layer2
            MoveToLayer Place.ToLayer3, TargetLayers.Layer3
            MoveToLayer Place.ToCropBoxLayer, TargetLayers.CropBoxLayer
        End If
    Next Place
End Sub

'===============================================================================
' # тесты

Private Sub testSomething()
    Debug.Print ActiveShape.SizeWidth
End Sub
