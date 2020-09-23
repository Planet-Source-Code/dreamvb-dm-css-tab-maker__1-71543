Attribute VB_Name = "Tools"
Option Explicit

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Type cRgb
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type TProject
    ID As String * 3 'ctm
    Items() As String
    Color1 As Long
    Color2 As Long
    Color3 As Long
    Color4 As Long
    Color5 As Long
    TabFloat As Integer
    fStyle As Integer
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Public ButtonPress As VbMsgBoxResult
Public Project As TProject
Public EditOP As Integer    '=0add, 1=edit
Public mUrlName As String
Public mUrlAddress As String
Public mUrlIndex As Integer

Public Const VBQuote = """"
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE As Long = &H40

Public Function FixPath(lPath As String) As String
    If Right$(lPath, 1) = "\" Then
        FixPath = lPath
    Else
        FixPath = lPath & "\"
    End If
End Function

Public Function FindFile(lzFileName As String) As Boolean
On Error Resume Next
    If Len(lzFileName) = 0 Then Exit Function
    FindFile = (GetAttr(lzFileName) And vbNormal) = vbNormal
    Err.Clear
End Function

Public Sub Long2Rgb(lColor As Long, RgbType As cRgb)
Dim Tmp As Long
On Error Resume Next

    Tmp = lColor
    'Convert Long To RGB
    With RgbType
        .Red = (Tmp Mod &H100)
        Tmp = (Tmp \ &H100)
        .Green = (Tmp Mod &H100)
        Tmp = (Tmp \ &H100)
        .Blue = (Tmp Mod &H100)
    End With
End Sub

Public Function RgbToHex(r, g, b) As String
Dim WebColor As OLE_COLOR
    WebColor = b + 256 * (g + 256 * r)
    'Format Hex to 6 places
    RgbToHex = Right$("000000" & Hex$(WebColor), 6)
End Function

Public Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String)
Dim bInf As BROWSEINFO
Dim RetVal As Long
Dim PathID As Long
Dim RetPath As String
Dim Offset As Integer
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    RetVal = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If RetVal Then
        Offset = InStr(RetPath, Chr$(0))
        GetFolder = Left$(RetPath, Offset - 1)
    End If
End Function

