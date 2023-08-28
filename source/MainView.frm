VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainView 
   Caption         =   "–‡ÒÍÎ‡‰ ÔÓ Ú‡·ÎËˆÂ"
   ClientHeight    =   3210
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

Public IsOk As Boolean
Public IsCancel As Boolean

Public PlaceWidth As TextBoxHandler
Attribute PlaceWidth.VB_VarHelpID = -1
Public PlaceHeight As TextBoxHandler
Attribute PlaceHeight.VB_VarHelpID = -1
Public Space As TextBoxHandler
Attribute Space.VB_VarHelpID = -1

'===============================================================================

Private Sub UserForm_Initialize()
    Set PlaceWidth = _
        TextBoxHandler.Create(TextBoxPlaceWidth, TextBoxTypeDouble, MIN_SIZE)
    Set PlaceHeight = _
        TextBoxHandler.Create(TextBoxPlaceHeight, TextBoxTypeDouble, MIN_SIZE)
    Set Space = _
        TextBoxHandler.Create(TextBoxSpace, TextBoxTypeDouble)
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub ButtonBrowseMotifsFolder_Click()
    Dim LastPath As String
    LastPath = TextBoxMotifsFolder
    Dim Folder As IFileSpec
    Set Folder = FileSpec.Create(TextBoxMotifsFolder)
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
            "Excel (*.xlsx)" & Chr(0) & "*.xlsx"
        .MultiSelect = False
        .InitialDir = FileSpec.Create(TextBoxTable).Path
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
