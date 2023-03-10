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
' ????????, ???????? ? ?????????? ????????
'===============================================================================

'@PredeclaredId
Option Explicit

#Const ReadOnly = True
#Const DefaultContent = False

'===============================================================================

Private Const iniFileName As String = "elvin_" & APP_NAME & ".ini"

'===============================================================================

Public BacksLayerName As String
Public BackTag As String
Public ContentMaxSizeMultiplierFace As Double
Public ContentMaxSizeMultiplierBack As Double
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
        
        .KeyVar "BacksLayerName", BacksLayerName
        .KeyVar "BackTag", BackTag
        .KeyVar "ContentMaxSizeMultiplierFace", _
            ContentMaxSizeMultiplierFace
        .KeyVar "ContentMaxSizeMultiplierBack", _
            ContentMaxSizeMultiplierBack
        .KeyVar "ContentTagsCommaSeparated", _
            ContentTagsCommaSeparated
        .KeyVar "CropBoxOutlineColor", CropBoxOutlineColor
        .KeyVar "CsvFile", CsvFile
        .KeyVar "CsvSeparator", CsvSeparator
        .KeyVar "DefaultLayerName", DefaultLayerName
        .KeyVar "FrontTag", FrontTag
        .KeyVar "GroupsMinDistance", GroupsMinDistance
        .KeyVar "Layer1Name", Layer1Name
        .KeyVar "Layer1TagsCommaSeparated", Layer1TagsCommaSeparated
        .KeyVar "Layer2Name", Layer2Name
        .KeyVar "Layer2TagsCommaSeparated", Layer2TagsCommaSeparated
        .KeyVar "Layer3Name", Layer3Name
        .KeyVar "Layer3TagsCommaSeparated", Layer3TagsCommaSeparated
        .KeyVar "LayerCropBoxName", LayerCropBoxName
        .KeyVar "MaxPlacesPerSideX", MaxPlacesPerSideX
        .KeyVar "MaxPlacesPerSideY", MaxPlacesPerSideY
        .KeyVar "PlacesMinDistanceX", PlacesMinDistanceX
        .KeyVar "PlacesMinDistanceY", PlacesMinDistanceY
        .KeyVar "SidesMinDistanceX", SidesMinDistanceX
        .KeyVar "SourceFolder", SourceFolder
    
    End With
    
End Sub

'-------------------------------------------------------------------------------

'???????????. ??????? ini-???? ????????? ?????????? (????. ? ??????????).
#If DefaultContent Then
Private Sub MakeDefaultFile(ByVal File As String)

    Dim Content As String
    
    Content = Content & "; ????????? Windows-1251" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "[DEFAULT]" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; ?????" & vbCrLf
    Content = Content & "Bulka=??????" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; ?????" & vbCrLf
    Content = Content & "Palka=1" & vbCrLf
    Content = Content & "" & vbCrLf
    Content = Content & "; ??????" & vbCrLf
    Content = Content & "Kryshka=2" & vbCrLf
    
    WriteFile Content, File
    
End Sub
#End If

'===============================================================================
'           * * * ???????, ?? ??????????????? ??? ????????? * * *
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
