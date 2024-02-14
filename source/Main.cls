VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===============================================================================
' # Declarations

Private Const TEXT_SPACE As Double = 0.1
Private Const BOTTOM_MARGIN As Double = 4

Private Type This
    Cfg As PresetsConfig
    Logger As SimpleLogger
End Type
Private This As This

Private WithEvents Imposer As MotifsImposer
Attribute Imposer.VB_VarHelpID = -1

'===============================================================================
' # Handlers

Private Sub Class_Initialize()
    Set This.Logger = New SimpleLogger
End Sub

Private Sub Class_Terminate()
    This.Logger.Check
End Sub

Private Sub Imposer_OnImportSurface(ByVal Surface As Surface)
    With Surface
        .Content.SetSize This.Cfg("PlaceWidth"), This.Cfg("PlaceHeight")
        Dim RectShape As Shape
        Set RectShape = _
            .Content.Layer.CreateRectangleRect(.Content.BoundingBox)
        Dim Title As Shape
        If This.Cfg("SetTitles") Then
            Set Title = SetTitle(.Content, .File.BaseName)
            Set .Content = PackShapes(RectShape, Title, .Content).Group
        Else
            Set .Content = PackShapes(RectShape, .Content).Group
        End If
    End With
End Sub

Private Sub Imposer_OnPageImposed(ByVal Page As Page)
    SetTextOnPageBottom Page, Page.Index
End Sub

'===============================================================================
' # Public

Public Sub MainRoutine(ByVal Cfg As PresetsConfig)
    Set This.Cfg = Cfg
    
    Dim Motifs As Collection
    Set Motifs = GetValidMotifs(GetRawMotifsFromTable)
    
    Set Imposer = MotifsImposer.NewFromImport(Motifs, False)
    With Imposer
        .EdgeMarginLeft = Cfg("SpaceWidth")
        .EdgeMarginRight = Cfg("SpaceWidth")
        .EdgeMarginTop = 1
        .EdgeMarginBottom = BOTTOM_MARGIN
        .VerticalSpace = VBA.IIf(Cfg("SetTitles"), 0, Cfg("SpaceHeight"))
        .HorizontalSpace = Cfg("SpaceWidth")
        .PlaceCropMarks = False
        
        .ImposeAutoAddPages
                
        Dim Page As Page
        Set Page = ActiveDocument.Pages.Last
        Dim Summary As Shape
        Set Summary = _
            SetTextOnPageBottom( _
                Page, "Видов: " & Motifs.Count & "   Мест: " & .PlacesTotal _
            )
        Summary.LeftX = Page.LeftX + TEXT_SPACE
    End With
End Sub

'===============================================================================
' # Helpers

Private Function GetRawMotifsFromTable() As Collection
    Dim TableFile As String
    TableFile = This.Cfg("TableFile")
    If Not FileExists(TableFile) Then
        Throw "Не найден файл " & vbCrLf & TableFile
    End If
    Dim Table As ITableFile
    Set Table = CsvUtilsTableFile.NewReadOnly(TableFile)
    Set GetRawMotifsFromTable = New Collection
    Dim Row As Long
    For Row = 1 To Table.MaxRow
        TryAddRawMotifFromTableRow GetRawMotifsFromTable, Table, Row
    Next Row
End Function

Private Sub TryAddRawMotifFromTableRow( _
                ByVal ioMotifs As Collection, _
                ByVal Table As ITableFile, _
                ByVal Row As Long _
            )
    With New Motif
        Set .SurfaceA = New Surface
        Set .SurfaceA.File = New FileSpec
        .SurfaceA.File.Path = This.Cfg("MotifsPath")
        
        On Error GoTo ErrFileName
        .SurfaceA.File.Name = Table.Cell(Row, 1)
        On Error GoTo ErrQuantity
        .Quantity = Table.Cell(Row, 2)
        On Error GoTo 0

        ioMotifs.Add .Self
    End With
    Exit Sub

ErrFileName:
    This.Logger.Add "Ошибка имени файла на строке " & Row
    On Error GoTo 0
    Exit Sub
ErrQuantity:
    This.Logger.Add "Ошибка количества на строке " & Row
    On Error GoTo 0
End Sub

Private Function GetValidMotifs(ByVal RawMotifs As Collection) As Collection
    Set GetValidMotifs = New Collection
    Dim Motif As Motif
    For Each Motif In RawMotifs
        With Motif.SurfaceA.File
            .Path = This.Cfg("MotifsPath")
            .Ext = "jpg"
            If Not FileExists(.ToString) Then .Ext = "png"
            If FileExists(.ToString) Then
                GetValidMotifs.Add Motif
            Else
                This.Logger.Add "Не найден файл " & .BaseName
            End If
        End With
    Next Motif
End Function

Private Function SetTitle( _
                     ByVal ParentShape As Shape, _
                     ByVal Text As String _
                 ) As Shape
    With ParentShape
        Dim Background As Shape
        Set Background = _
            .Layer.CreateRectangle( _
                .LeftX, .BottomY, .RightX, .BottomY - This.Cfg("SpaceHeight") _
            )
        Background.Outline.SetNoOutline
        Background.Fill.ApplyNoFill
            
        Dim Title As Shape
        Set Title = .Layer.CreateArtisticText(0, 0, Text)
        FitInside Title, SpaceBox(Background, -TEXT_SPACE)
        
        Set SetTitle = PackShapes(Title, Background).Group
        
    End With
End Function

Private Function SetTextOnPageBottom( _
                     ByVal Page As Page, ByVal Text As String _
                 ) As Shape
    Dim Shape As Shape
    Set Shape = Page.ActiveLayer.CreateArtisticText(0, 0, Text)
    Shape.SetSize , BOTTOM_MARGIN - TEXT_SPACE * 2
    Shape.RightX = Page.RightX - TEXT_SPACE
    Shape.BottomY = Page.BottomY + TEXT_SPACE
    Set SetTextOnPageBottom = Shape
End Function
