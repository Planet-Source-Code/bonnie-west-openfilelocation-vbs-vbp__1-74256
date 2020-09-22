Attribute VB_Name = "modOpenFileLocation"
Option Explicit

Public wsShell As WshShell

Private Sub Main()
    Dim CmdLine As String

    Set wsShell = New WshShell
    CmdLine = Command$

    If LenB(CmdLine) Then OpenFileLocation CmdLine Else InstallUninstallOFL

    Set wsShell = Nothing
End Sub

Private Sub OpenFileLocation(ByRef sFileSpec As String)
    Dim sTarget As String

    On Error Resume Next
    sTarget = wsShell.CreateShortcut(sFileSpec).TargetPath

    If FileExists(sTarget) Then
        Shell "explorer.exe /select,""" & sTarget & """", vbNormalFocus
    ElseIf FolderExists(sTarget) Then
        Shell "explorer.exe /select,""" & sTarget & """", vbNormalFocus
    ElseIf InStrB(" " & LCase$(sFileSpec) & " ", " /src ") Then
        Extract_SourceCode
    Else
        MsgBox "Could not find:" & vbNewLine & vbNewLine & _
               """" & sTarget & """", vbExclamation, OFL
    End If
End Sub

Public Function FileExists(ByRef sFileName As String) As Boolean
    On Error Resume Next
    If LenB(sFileName) Then _
        FileExists = Dir(sFileName, vbArchive Or vbHidden Or _
                         vbReadOnly Or vbSystem) <> vbNullString
End Function

Private Function FolderExists(ByRef sPath As String) As Boolean
    On Error Resume Next
    If LenB(sPath) Then FolderExists = Dir(sPath & "\NUL") <> vbNullString
End Function
