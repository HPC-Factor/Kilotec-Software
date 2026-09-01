Attribute VB_Name = "Module1"
Option Explicit
'Const ID_CLEAR = 101
Const ID_EXIT = 101

Const ID_NEW = 201
Const ID_DELETE = 202
Const ID_DELETEALL = 203

Const ID_HELP = 301
Const ID_INFO = 302

Const SIZE_SUPPORTSBUTTONS = 7500

Public strPath As String


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Main()
    Dim strMemTitle As String

    If App.PrevInstance Then
        ' Keine zweite Instanz des Programms erlauben.
        strMemTitle = App.Title
        App.Title = "%&SuperClipboard - zweite Instanz"
        AppActivate strMemTitle
    Else
        strPath = App.Path
        If Right(strPath, 1) <> "\" Then strPath = strPath & "\"

        frmMain.Show
    End If
End Sub

