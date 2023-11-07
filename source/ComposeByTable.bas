Attribute VB_Name = "ComposeByTable"
'===============================================================================
'   Макрос          : ComposeByTable
'   Версия          : 2023.11.07
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "ComposeByTable"

'===============================================================================

Private Const PAGE_WIDTH As Double = 210 'мм
Private Const PAGE_HEIGHT As Double = 297 'мм

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    Dim Cfg As PresetsConfig
    Set Cfg = PresetsConfig.Create("elvin_ComposeByTable", Defaults)
    
    Dim View As New MainView
    CfgToMainView View, Cfg
    View.Show
    MainViewToCfg View, Cfg
    
    If View.IsCancel Then Exit Sub
        
    CreateDocument
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.MasterPage.SetSize PAGE_WIDTH, PAGE_HEIGHT
    
    BoostStart APP_NAME, RELEASE
    
    With New Main
        .MainRoutine Cfg
    End With
    
    ActiveDocument.ClearSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # настройки

Private Sub CfgToMainView( _
                ByVal View As MainView, _
                ByVal Cfg As PresetsConfig _
            )
    With View
        .PlaceWidth = Cfg("PlaceWidth")
        .PlaceHeight = Cfg("PlaceHeight")
        .SpaceWidth = Cfg("SpaceWidth")
        .SpaceHeight = Cfg("SpaceHeight")
        .TextBoxMotifsFolder = Cfg("MotifsPath")
        .TextBoxTable = Cfg("TableFile")
    End With
End Sub

Private Sub MainViewToCfg( _
                ByVal View As MainView, _
                ByVal Cfg As PresetsConfig _
            )
    With View
        Cfg("PlaceWidth") = .PlaceWidth
        Cfg("PlaceHeight") = .PlaceHeight
        Cfg("SpaceWidth") = .SpaceWidth
        Cfg("SpaceHeight") = .SpaceHeight
        Cfg("MotifsPath") = .TextBoxMotifsFolder
        Cfg("TableFile") = .TextBoxTable
    End With
End Sub

Private Property Get Defaults() As Dictionary
    Set Defaults = New Dictionary
    With Defaults
        .Item("PlaceWidth") = 100#
        .Item("PlaceHeight") = 100#
        .Item("SpaceWidth") = 0
        .Item("SpaceHeight") = 0
    End With
End Property

'===============================================================================
' # тесты

Private Sub Test1()
    Dim File As String
    File = ""
    With CsvUtilsTableFile.CreateReadOnly(File, , 2, 1)
        Show .Cell(1, 1)
    End With
End Sub
