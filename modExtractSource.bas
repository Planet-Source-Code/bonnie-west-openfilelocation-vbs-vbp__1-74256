Attribute VB_Name = "modExtractSource"
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
