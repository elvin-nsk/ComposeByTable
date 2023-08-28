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

Private Type typeThis
    Cfg As PresetsConfig
    Logger As SimpleLogger
End Type
Private This As typeThis

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

Private Sub Imposer_OnImportFile(ByVal ioMotif As Motif)
    ioMotif.SurfaceA.Content.SetSize _
        This.Cfg("PlaceWidth"), This.Cfg("PlaceHeight")
End Sub

'===============================================================================
' # Public

Public Sub MainRoutine(ByVal Cfg As PresetsConfig)
    Set This.Cfg = Cfg
    
    Dim Motifs As Collection
    Set Motifs = GetValidMotifs(GetRawMotifsFromTable)
    
    Set Imposer = MotifsImposer.CreateFromImport(Motifs, False)
    With Imposer
        .SetEdgeMarginsEqual Cfg("Space")
        .VerticalSpace = Cfg("Space")
        .HorizontalSpace = Cfg("Space")
        .PlaceCropMarks = False
        
        .ImposeAutoAddPages
    End With
End Sub

'===============================================================================
' # Helpers

Private Function GetRawMotifsFromTable() As Collection
    Dim TableFile As IFileSpec
    Set TableFile = FileSpec.Create(This.Cfg("TableFile"))
    If Not FileExists(TableFile) Then
        Throw "Не найден файл " & vbCrLf & TableFile
    End If
    Dim Table As ITableFile
    Set Table = ExcelConnection.CreateReadOnly(TableFile)
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
        Set .SurfaceA.File = FileSpec.Create
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
                This.Logger.Add "Не найден файл " & .Name
            End If
        End With
    Next Motif
End Function
