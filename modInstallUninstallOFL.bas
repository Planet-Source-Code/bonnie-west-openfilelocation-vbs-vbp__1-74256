Attribute VB_Name = "modInstallUninstallOFL"
Option Explicit

 Public Const OFL    As String = "OpenFileLocation"

Private Const HKCU   As String = "HKCU\"
Private Const HKLM   As String = "HKLM\"
Private Const sKEY   As String = "SOFTWARE\Classes\lnkfile\shell\OpenFileLocation\"
Private Const sVALUE As String = "Open &file location"

Private App_EXEName  As String

Public Sub InstallUninstallOFL()
    Dim sPrompt As String, iButtons As Integer

    App_EXEName = "\" & App.EXEName & ".exe"

    sPrompt = "Do you want to add the ""Open file location"" context menu " & _
              "option to shortcut files?" & vbNewLine & "(Select NO to remove)"
    iButtons = vbYesNoCancel Or vbQuestion Or vbDefaultButton3

    Select Case MsgBox(sPrompt, iButtons, "Install " & OFL)
        Case vbYes: sPrompt = Environ$("WINDIR"):   InstallOFL sPrompt
        Case vbNo:  sPrompt = Environ$("WINDIR"): UninstallOFL sPrompt
    End Select
End Sub

Private Sub InstallOFL(ByRef WinDir As String)
    Dim EXE() As Byte

    On Error GoTo Catch

    If App.Path & App_EXEName = WinDir & App_EXEName Then
        App_EXEName = WinDir & App_EXEName
        GoTo 1
    End If

    Open App.Path & App_EXEName For Binary As #7
        ReDim EXE(LOF(7))
        Get #7, 1, EXE
    Close

    App_EXEName = WinDir & App_EXEName

    If FileExists(App_EXEName) Then Kill App_EXEName

    Open App_EXEName For Binary As #9
        Put #9, 1, EXE
        Erase EXE
    Close

1   Select Case MsgBox("Install for ALL User Accounts?", _
                vbQuestion Or vbYesNo Or vbDefaultButton2, OFL)
        Case vbYes: WinDir = HKLM
        Case vbNo:  WinDir = HKCU
    End Select

    App_EXEName = App_EXEName & " %1"

    wsShell.RegWrite WinDir & sKEY, sVALUE, "REG_SZ"
    wsShell.RegWrite WinDir & sKEY & "command\", App_EXEName, "REG_SZ"

    MsgBox "Installed successfully!", vbInformation, OFL
    Exit Sub

Catch:
    MsgBox Err.Description, vbCritical, Err.Source
End Sub

Private Sub UninstallOFL(ByRef WinDir As String)
    On Error GoTo 4
    wsShell.RegDelete HKCU & sKEY & "command\"
    wsShell.RegDelete HKCU & sKEY
    GoTo 2

1   On Error GoTo 5
    wsShell.RegDelete HKLM & sKEY & "command\"
    wsShell.RegDelete HKLM & sKEY

2   MsgBox "Uninstalled successfully!", vbInformation, OFL

3   If FileExists(WinDir & App_EXEName) Then NukeOFL WinDir
    Exit Sub

4   On Error GoTo 0
    Resume 1

5   MsgBox "Unable to remove registry keys ""HKCU|HKLM\" _
           & sKEY & """.", vbCritical, Err.Source
    Resume 3
End Sub

Private Sub NukeOFL(ByRef WinDir As String)
    App_EXEName = WinDir & "\" & App.EXEName & "_Nuke.bat"

    If FileExists(App_EXEName) Then Kill App_EXEName

    Open App_EXEName For Output As #7
        Print #7, "TITLE Self-Deleting Uninstaller"
        Print #7, Left$(WinDir, 2)                  'Change Drive to SystemDrive
        Print #7, "CD """ & WinDir & """"           'Change Directory to WinDir
        Print #7, "PING 127.0.0.1"    'Do something else while this program ends
        Print #7, "DEL /F /Q """ & App.EXEName & ".exe"""
        Print #7, "DEL /F /Q """ & App.EXEName & "_Nuke.bat""";
    Close

    Shell """" & App_EXEName & """", vbHide
End Sub
