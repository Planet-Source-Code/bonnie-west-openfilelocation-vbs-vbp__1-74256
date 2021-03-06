        ��  ��                  �      �� ��     0           <?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
<assembly xmlns="urn:schemas-microsoft-com:asm.v1" manifestVersion="1.0">
    <assemblyIdentity
        name="Bonnie West.Open File Location.OpenFileLocation"
        processorArchitecture="X86"
        type="win32"
        version="1.0.0.0" />
    <description>Open File Location</description>
    <dependency>
        <dependentAssembly>
            <assemblyIdentity
                language="*"
                name="Microsoft.Windows.Common-Controls"
                processorArchitecture="X86"
                publicKeyToken="6595b64144ccf1df"
                type="win32"
                version="6.0.0.0" />
        </dependentAssembly>
    </dependency>
</assembly>  �      ��
 ��     0           Type=Exe
Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#WINDOWS\system32\stdole2.tlb#OLE Automation
Reference=*\G{F935DC20-1CF0-11D0-ADB9-00C04FD58A0B}#1.0#0#WINDOWS\system32\wshom.ocx#Windows Script Host Object Model
Module=modOpenFileLocation; modOpenFileLocation.bas
Module=modInstallUninstallOFL; modInstallUninstallOFL.bas
Module=modExtractSource; modExtractSource.bas
ResFile32="prjOpenFileLocation.res"
Startup="Sub Main"
HelpFile=""
Title="Open File Location"
ExeName32="OpenFileLocation.exe"
Command32=""
Name="OpenFileLocation"
HelpContextID="0"
Description="Open File Location"
CompatibleMode="0"
MajorVer=1
MinorVer=0
RevisionVer=0
AutoIncrementVer=1
ServerSupportFiles=0
VersionComments="Open File Location"
VersionCompanyName="Bonnie West�"
VersionFileDescription="Open File Location"
VersionLegalCopyright="� Bonnie West"
VersionLegalTrademarks="Bonnie West�"
VersionProductName="Open File Location"
CompilationType=0
OptimizationType=0
FavorPentiumPro(tm)=0
CodeViewDebugInfo=0
NoAliasing=-1
BoundsCheck=-1
OverflowCheck=-1
FlPointCheck=-1
FDIVCheck=-1
UnroundedFP=-1
StartMode=0
Unattended=0
Retained=0
ThreadPerObject=0
MaxNumberOfThreads=1
DebugStartupOption=0
  {       ��
 ��     0           modOpenFileLocation = 0, 0, 1012, 651, 
modInstallUninstallOFL = 0, 0, 1012, 651, 
modExtractSource = 0, 0, 1012, 651, 
 �      ��
 ��     0                   ��  ��                        ��
 ��     0           1 24 OpenFileLocation.exe.manifest

1 RCDATA prjOpenFileLocation.vbp
2 RCDATA prjOpenFileLocation.vbw
3 RCDATA resOpenFileLocation.res
4 RCDATA modOpenFileLocation.bas
5 RCDATA modInstallUninstallOFL.bas
6 RCDATA modExtractSource.bas
7 RCDATA OpenFileLocation.vbs       ��
 ��     0           @ECHO OFF
TITLE Compiling Resource Script...
SETLOCAL
SET RC=prjOpenFileLocation.rc
ECHO.
ECHO "%PROGRAMFILES%\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /l 0 %RC%
"%PROGRAMFILES%\Microsoft Visual Studio\VB98\Wizards\RC.EXE" /r /l 0 %RC%
ECHO.
ECHO Done!!!
PAUSE > NUL  |      ��
 ��     0           Attribute VB_Name = "modOpenFileLocation"
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
�      ��
 ��     0           Attribute VB_Name = "modInstallUninstallOFL"
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
   4      ��
 ��     0           Attribute VB_Name = "modExtractSource"
Option Explicit

Private Type BROWSEINFO
    hwndOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SendMessageA Lib "user32" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, ByVal lParam$) As Long
Private Declare Function SHBrowseForFolderA Lib "shell32" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pidl&, ByVal pszPath$) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal pv&)

Sub Extract_SourceCode()
    Dim sPath$

    sPath = Browse4Folder(GetDesktopWindow) & "\"
    If LenB(sPath) = 2 Then Exit Sub

    DumpRes2File 1, 10, sPath & "prjOpenFileLocation.vbp"
    DumpRes2File 2, 10, sPath & "prjOpenFileLocation.vbw"
    DumpRes2File 3, 10, sPath & "prjOpenFileLocation.res"
    DumpRes2File 4, 10, sPath & "modOpenFileLocation.bas"
    DumpRes2File 5, 10, sPath & "modInstallUninstallOFL.bas"
    DumpRes2File 6, 10, sPath & "modExtractSource.bas"
    DumpRes2File 7, 10, sPath & "OpenFileLocation.vbs"

    Shell "explorer /e,/select,""" & sPath & "prjOpenFileLocation.vbp""", vbNormalFocus
End Sub

Private Sub DumpRes2File(nResIdx%, nResFrmt%, sFilePathName$)
    Dim bResFile() As Byte

    bResFile = LoadResData(nResIdx, nResFrmt)

    On Error GoTo 1
    If FileExists(sFilePathName) Then Kill sFilePathName
    Open sFilePathName For Binary As #7
        Put #7, 1, bResFile
        GoTo 2

1       MsgBox Err.Description & vbCrLf & vbCrLf & """" & sFilePathName & """", vbExclamation, Err.Source
2   Close
End Sub

Private Function Browse4Folder(hWnd&) As String
    Const BIF_RETURNONLYFSDIRS = &H1
    Const BIF_EDITBOX = &H10
    Const BIF_NEWDIALOGSTYLE = &H40
    Dim udtBI As BROWSEINFO

    udtBI.hwndOwner = hWnd
    udtBI.lpszTitle = "Select a directory to extract the files into:"
    udtBI.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_EDITBOX Or BIF_NEWDIALOGSTYLE
    udtBI.lpfn = GetAddrOfFunc(AddressOf BrowseCallbackProc)

    hWnd = SHBrowseForFolderA(udtBI)
    If hWnd = 0 Then Exit Function
    Browse4Folder = Space$(260)
    SHGetPathFromIDListA hWnd, Browse4Folder
    CoTaskMemFree hWnd
    Browse4Folder = Left$(Browse4Folder, InStr(Browse4Folder, vbNullChar) - 1)
End Function

Private Function BrowseCallbackProc(ByVal hWnd&, ByVal uMsg&, ByVal lParam&, ByVal lpData&) As Long
    Const BFFM_INITIALIZED = 1
    Const BFFM_SETSELECTION = &H466

    If uMsg = BFFM_INITIALIZED Then SendMessageA hWnd, BFFM_SETSELECTION, 1, App.Path
End Function

Private Function GetAddrOfFunc(Address&) As Long
    GetAddrOfFunc = Address
End Function
�      ��
 ��     0           'OpenFileLocation.vbs
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
