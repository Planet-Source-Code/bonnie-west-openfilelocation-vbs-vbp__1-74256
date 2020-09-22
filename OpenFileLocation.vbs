'OpenFileLocation.vbs
'(Bonnie West)

'Provides functionality similar to Windows Vista's "Open file location" context
'menu for pre-Vista OSes. In Vista, shortcut files have a handy context menu
'option, that upon choosing, pre-selects that shortcut's target in a new
'Explorer window. Requires an enabled Microsoft Windows Script Host (wscript.exe).
'Open it once to install or uninstall.

Option Explicit

Private Const sKEY = "HKCU\Software\Classes\lnkfile\shell\OpenFileLocation\"
                     'You may also put this under HKLM\SOFTWARE\Classes\lnkfile
                     'if you want all user profiles to have this context menu
Private Const sVALUE = "Open &file location"
                       '&f immediately selects this menu unlike the default
                       '&i in Vista which collides with "P&in to Start menu"
Private Const sCMD = "wscript.exe %WINDIR%\OpenFileLocation.vbs ""%1"""
                     'Save this in a file named "OpenFileLocation.vbs" in your
                     '"\WINDOWS" directory, or if you prefer, edit the
                     'location & filename in this constant
Private wsShell

Set wsShell = WScript.CreateObject("WScript.Shell")

If WScript.Arguments.Count Then    'If arguments were passed to this file, Then
    OpenFileLocation               '    a shortcut file's location was specified
Else                               'Else, no arguments were passed
    InstallUninstallOFL            '    go to Install/Uninstall mode
End If

Set wsShell = Nothing    'Plug leaks

Private Sub OpenFileLocation
    Dim FSO, oShortcut, sFileSpec, sTarget

   'Get the shortcut file's location
    sFileSpec = WScript.Arguments.Item(0)
   'Open the shortcut to expose its properties
    Set oShortcut = wsShell.CreateShortcut(sFileSpec)
   'Retrieve the shortcut's target into sTarget
    sTarget = oShortcut.TargetPath

    Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
   'If the shortcut points to an existing file or folder
    If FSO.FileExists(sTarget) Then
       'Pre-select that target in a new Explorer window
        wsShell.Run "explorer.exe /select,""" & sTarget & """"
    ElseIf FSO.FolderExists(sTarget) Then
       'Short-circuit the preceding expressions instead of using Or
        wsShell.Run "explorer.exe /select,""" & sTarget & """"
    Else
       'Complain, er, inform if it's missing
        MsgBox "Could not find:" & vbNewLine & vbNewLine & _
               """" & sTarget & """", vbExclamation, "OpenFileLocation"
    End If

    Set FSO = Nothing
    Set oShortcut = Nothing    'Plug leaks
End Sub

Private Sub InstallUninstallOFL            'Does this still need explanation ???
    Dim sPrompt, iButtons

    sPrompt = "Do you want to add the ""Open file location"" context menu " & _
              "option to shortcut files?" & vbNewLine & "(Select NO to remove)"
    iButtons = vbYesNoCancel + vbQuestion + vbDefaultButton3

    Select Case MsgBox(sPrompt, iButtons, "Install OpenFileLocation.vbs")
        Case vbYes: InstallOFL
        Case vbNo:  UninstallOFL
    End Select
End Sub

Private Sub InstallOFL                    'On installation, add the context menu
    On Error Resume Next                  '              entries to the Registry
    wsShell.RegWrite sKEY, sVALUE, "REG_SZ"
    wsShell.RegWrite sKEY & "command\", sCMD, "REG_EXPAND_SZ"

    If Err Then
        MsgBox Err.Description, vbCritical, Err.Source
    Else
        MsgBox "Installed successfully!", vbInformation, "OpenFileLocation"
    End If
End Sub

Private Sub UninstallOFL             'On uninstallation, remove the context menu
    On Error Resume Next             '                 entries from the Registry
    wsShell.RegDelete sKEY & "command\"
    wsShell.RegDelete sKEY

    If Err Then
        MsgBox Err.Description, vbCritical, Err.Source
    Else
        MsgBox "Uninstalled successfully!", vbInformation, "OpenFileLocation"
    End If
End Sub
