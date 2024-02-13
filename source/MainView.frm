VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "–‡ÒÍÎ‡‰ ÔÓ Ú‡·ÎËˆÂ"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "MainView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'===============================================================================

Private Const MIN_SIZE As Double = 0.1
Private Const MIN_VERTICAL_SPACE As Double = 4

Public IsOk As Boolean
Public IsCancel As Boolean

Public PlaceWidth As TextBoxHandler
Attribute PlaceWidth.VB_VarHelpID = -1
Public PlaceHeight As TextBoxHandler
Attribute PlaceHeight.VB_VarHelpID = -1
Public SpaceWidth As TextBoxHandler
Attribute SpaceWidth.VB_VarHelpID = -1
Public SpaceHeight As TextBoxHandler

'===============================================================================

Private Sub UserForm_Initialize()
    Set PlaceWidth = _
        TextBoxHandler.Create(TextBoxPlaceWidth, TextBoxTypeDouble, MIN_SIZE)
    Set PlaceHeight = _
        TextBoxHandler.Create(TextBoxPlaceHeight, TextBoxTypeDouble, MIN_SIZE)
    Set SpaceWidth = _
        TextBoxHandler.Create(TextBoxSpaceWidth, TextBoxTypeDouble)
    Set SpaceHeight = _
        TextBoxHandler.Create( _
            TextBoxSpaceHeight, TextBoxTypeDouble, MIN_VERTICAL_SPACE _
        )
    LabelMinVerticalSpace = "( ÏËÌËÏÛÏ " & MIN_VERTICAL_SPACE & " )"
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub ButtonBrowseMotifsFolder_Click()
    Dim LastPath As String
    LastPath = TextBoxMotifsFolder
    Dim Folder As New FileSpec
    Folder.Inject TextBoxMotifsFolder
    Folder.Path = CorelScriptTools.GetFolder(Folder.Path)
    If Folder.Path = "\" Then
        TextBoxMotifsFolder = LastPath
    Else
        TextBoxMotifsFolder = Folder.Path
    End If
End Sub

Private Sub ButtonBrowseTable_Click()
    With New FileBrowser
        .Filter = _
            "Excel (*.csv)" & Chr(0) & "*.csv"
        .MultiSelect = False
        Dim File As New FileSpec
        File.Inject TextBoxTable
        .InitialDir = File.Path
        Dim Result As Collection
        Set Result = .ShowFileOpenDialog
        If Result.Count > 0 Then TextBoxTable = Result(1)
    End With
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================

Private Sub FormŒ ()
    Me.Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Me.Hide
    IsCancel = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
