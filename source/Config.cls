VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' Config v2022.12.22
' хранение, загрузка и сохранение настроек
'===============================================================================

'@PredeclaredId
Option Explicit

#Const ReadOnly = False
#Const DefaultContent = False

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public BackTag As String
Public ContentMaxSizeMultiplierToCropBox As Double
Public ContentTagsCommaSeparated As String
Public CropBoxOutlineColor As String
Public CsvFile As String
Public CsvSeparator As String
Public DefaultLayerName As String
Public FrontTag As String
Public GroupsMinDistance As Double
Public Layer1Name As String
Public Layer1TagsCommaSeparated As String
Public Layer2Name As String
Public Layer2TagsCommaSeparated As String
Public Layer3Name As String
Public Layer3TagsCommaSeparated As String
Public LayerCropBoxName As String
Public MaxPlacesPerSideX As Long
Public MaxPlacesPerSideY As Long
Public PlacesMinDistanceX As Double
Public PlacesMinDistanceY As Double
Public SidesMinDistanceX As Double
Public SourceFolder As String

'-------------------------------------------------------------------------------

Private Sub Assign(ByVal Action As iniAction)

    With iniFile.Create
    
        iniFilePrepare .Self, Action
        
        .KeyVar "BackTag", BackTag, _
            "design_back"
        .KeyVar "ContentMaxSizeMultiplierToCropBox", _
            ContentMaxSizeMultiplierToCropBox, 0.75
        .KeyVar "ContentTagsCommaSeparated", _
            ContentTagsCommaSeparated, "Name,tel1,tel2"
        .KeyVar "CropBoxOutlineColor", CropBoxOutlineColor, _
            "CMYK,USER,100,0,100,0"
        .KeyVar "CsvFile", CsvFile, _
            "e:\WORK\макросы Corel\на заказ\Евгений Галанов\PlaceByCsv\материалы\test\import.csv"
        .KeyVar "CsvSeparator", CsvSeparator, _
            ","
        .KeyVar "DefaultLayerName", DefaultLayerName, _
            "Элементы дизайна"
        .KeyVar "FrontTag", FrontTag, _
            "design_front"
        .KeyVar "GroupsMinDistance", GroupsMinDistance, _
            60
        .KeyVar "Layer1Name", Layer1Name, _
            "Имена"
        .KeyVar "Layer1TagsCommaSeparated", Layer1TagsCommaSeparated, _
            "Name"
        .KeyVar "Layer2Name", Layer2Name, _
            "Телефоны"
        .KeyVar "Layer2TagsCommaSeparated", Layer2TagsCommaSeparated, _
            "tel1,tel2"
        .KeyVar "Layer3Name", Layer3Name, _
            "Комментарии"
        .KeyVar "Layer3TagsCommaSeparated", Layer3TagsCommaSeparated, _
            "Color,Number,Comment"
        .KeyVar "LayerCropBoxName", LayerCropBoxName, _
            "Рамка"
        .KeyVar "MaxPlacesPerSideX", MaxPlacesPerSideX, _
            2
        .KeyVar "MaxPlacesPerSideY", MaxPlacesPerSideY, _
            4
        .KeyVar "PlacesMinDistanceX", PlacesMinDistanceX, _
            4
        .KeyVar "PlacesMinDistanceY", PlacesMinDistanceY, _
            1.6
        .KeyVar "SidesMinDistanceX", SidesMinDistanceX, _
            48
        .KeyVar "SourceFolder", SourceFolder, _
            "c:\temp"
    
    End With
    
End Sub

'-------------------------------------------------------------------------------

'Опционально. Создаёт ini-файл заданного содержания (напр. с комментами).
#If DefaultContent Then
Private Sub MakeDefaultFile(ByVal File As String)

    Dim Content As String
    
    Content = Content & "; кодировка Windows-1251" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "[DEFAULT]" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; булка" & vbCrLf
    Content = Content & "Bulka=Рулька" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; палка" & vbCrLf
    Content = Content & "Palka=1" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; крышка" & vbCrLf
    Content = Content & "Kryshka=2" & vbCrLf
    
    WriteFile Content, File
    
End Sub
#End If

'===============================================================================
'           * * * Контент, не предназначенный для изменения * * *
'===============================================================================
' # Constructors

Public Function Bind() As Config
    Set Bind = New Config
    Bind.ForceLoad
End Function

#If Not ReadOnly Then
Private Sub Class_Terminate()
    ForceSave
End Sub
#End If

'===============================================================================
' # Public

Public Sub ForceLoad()
    Assign iniRead
End Sub

#If Not ReadOnly Then
Public Sub ForceSave()
    Assign iniWrite
End Sub
#End If

Public Property Get Self() As Config
    Set Self = Me
End Property

'===============================================================================

Private Sub iniFilePrepare(ByVal iniFile As iniFile, ByVal Action As iniAction)

    Dim iniFilePath As String

    With iniFile
        Select Case True
        Case FileExists(GMSManager.UserGMSPath & iniFileName)
            .File = Application.GMSManager.UserGMSPath & iniFileName
            .Action = Action
        Case FileExists(GMSManager.GMSPath & iniFileName)
            .File = Application.GMSManager.GMSPath & iniFileName
            .Action = Action
        Case Else
            .File = GMSManager.UserGMSPath & iniFileName
            #If DefaultContent Then
            MakeDefaultFile .File
            If Action = iniRead Then
                .Action = iniRead
            Else
                .Action = iniWrite
            End If
            #Else
            If Action = iniRead Then .Action = iniReset Else .Action = iniWrite
            #End If
        End Select
    End With

End Sub

Private Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Private Sub WriteFile( _
               ByVal Content As String, _
               ByVal File As String _
           )
    Dim FileNum As Long
    FileNum = VBA.FreeFile
    Open File For Output Access Write As #FileNum
    Print #FileNum, Content
    Close #FileNum
End Sub
