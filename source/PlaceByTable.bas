Attribute VB_Name = "PlaceByTable"
'===============================================================================
'   Макрос          : PlaceByTable
'   Версия          : 2022.12.22
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = False

Public Const APP_NAME As String = "PlaceByTable"

'===============================================================================

Private Const SomeConst As String = ""

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    Dim Cfg As Config
    Set Cfg = Config.Bind
    
    Dim Table As Dictionary
    Set Table = ParseTable(GetTable(Cfg.CsvFile, Cfg.CsvSeparator))
    
    Dim Groups As Collection
    Set Groups = ProcessTableAsGroups(Table, Cfg)
    Debug.Print Groups.Count
    
    Dim Imposition As Document
    
    
    'BoostStart APP_NAME, RELEASE
    
    
    
Finally:
    'BoostFinish
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
    
    Do
        
        CurrentPairIndex = 1
        Set Group = New GroupTwoSides
        Set Group.Fronts = New Collection
        Set Group.Backs = New Collection
        
        Do
            CurrentPairIndex = CurrentPairIndex + 1
            Row = Row + 1
            Group.Fronts.Add ProcessPlace(True, Row, Table, Cfg)
            Group.Backs.Add ProcessPlace(False, Row, Table, Cfg)
        Loop Until CurrentPairIndex >= MaxPlacesPerSide
        
        ProcessTableAsGroups.Add Group
        
    Loop Until Row >= TotalRows
    
End Function

Private Function ProcessPlace( _
                     ByVal Face As Boolean, _
                     ByVal Row As Long, _
                     ByVal Table As Dictionary, _
                     ByVal Cfg As Config _
                 ) As Place

End Function

'===============================================================================
' # тесты

Private Sub testSomething()
'
End Sub
