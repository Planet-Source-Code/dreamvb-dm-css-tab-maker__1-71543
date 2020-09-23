VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frmmain 
   Caption         =   "DM CSS Tab Maker"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   10605
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdButton 
      Height          =   405
      Index           =   7
      Left            =   15
      Picture         =   "frmmain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   32
      Tag             =   "EXPORT"
      ToolTipText     =   "Export"
      Top             =   3180
      Width           =   450
   End
   Begin VB.CommandButton cmdButton 
      Enabled         =   0   'False
      Height          =   405
      Index           =   5
      Left            =   15
      Picture         =   "frmmain.frx":0253
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "DELETE"
      ToolTipText     =   "Delete Item"
      Top             =   2790
      Width           =   450
   End
   Begin VB.CommandButton cmdButton 
      Enabled         =   0   'False
      Height          =   405
      Index           =   4
      Left            =   15
      Picture         =   "frmmain.frx":032A
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "EDIT"
      ToolTipText     =   "Edit Item"
      Top             =   2400
      Width           =   450
   End
   Begin VB.CommandButton cmdButton 
      Height          =   405
      Index           =   3
      Left            =   15
      Picture         =   "frmmain.frx":03E3
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "ADD"
      ToolTipText     =   "Add Item"
      Top             =   2010
      Width           =   450
   End
   Begin VB.CommandButton cmdButton 
      Height          =   405
      Index           =   6
      Left            =   15
      Picture         =   "frmmain.frx":054B
      Style           =   1  'Graphical
      TabIndex        =   31
      Tag             =   "RELOAD"
      ToolTipText     =   "Refresh"
      Top             =   1620
      Width           =   450
   End
   Begin VB.ComboBox cboFontStyle 
      Height          =   315
      Left            =   1005
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   6690
      Width           =   1215
   End
   Begin VB.ComboBox cboFloat 
      Height          =   315
      Left            =   1005
      Style           =   2  'Dropdown List
      TabIndex        =   28
      Top             =   6300
      Width           =   1215
   End
   Begin VB.PictureBox pTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H009B6900&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   345
      Index           =   3
      Left            =   15
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   25
      Top             =   5880
      Width           =   3090
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Other Properties ::"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   26
         Top             =   75
         Width           =   1530
      End
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   150
      ScaleHeight     =   150
      ScaleWidth      =   135
      TabIndex        =   23
      Top             =   5580
      Width           =   195
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   150
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   21
      Top             =   5325
      Width           =   195
   End
   Begin VB.PictureBox pTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H009B6900&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   345
      Index           =   2
      Left            =   15
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   19
      Top             =   4905
      Width           =   3090
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font Color Properties ::"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   20
         Top             =   75
         Width           =   1920
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   3270
      Top             =   2445
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00CFCFCF&
      Height          =   195
      Index           =   2
      Left            =   150
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   18
      Top             =   4620
      Width           =   195
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H00D99300&
      Height          =   195
      Index           =   1
      Left            =   150
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   15
      Top             =   4380
      Width           =   195
   End
   Begin VB.PictureBox pColor 
      BackColor       =   &H009B6900&
      Height          =   195
      Index           =   0
      Left            =   150
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   13
      Top             =   4140
      Width           =   195
   End
   Begin VB.PictureBox pTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H009B6900&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   345
      Index           =   1
      Left            =   15
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   11
      Top             =   3690
      Width           =   3090
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tab Color Properties ::"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   12
         Top             =   75
         Width           =   1860
      End
   End
   Begin SHDocVwCtl.WebBrowser WebView 
      Height          =   2250
      Left            =   3150
      TabIndex        =   10
      Top             =   90
      Width           =   7290
      ExtentX         =   12859
      ExtentY         =   3969
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox pBar 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   10605
      TabIndex        =   9
      Top             =   7230
      Width           =   10605
      Begin VB.Line ln3d 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   480
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line ln3d 
         BorderColor     =   &H80000003&
         Index           =   0
         X1              =   0
         X2              =   480
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.CommandButton cmdButton 
      Height          =   405
      Index           =   2
      Left            =   15
      Picture         =   "frmmain.frx":06C8
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "SAVE"
      ToolTipText     =   "Save Project"
      Top             =   1230
      Width           =   450
   End
   Begin VB.CommandButton cmdButton 
      Height          =   405
      Index           =   1
      Left            =   15
      Picture         =   "frmmain.frx":074A
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "OPEN"
      ToolTipText     =   "Open Project"
      Top             =   840
      Width           =   450
   End
   Begin VB.CommandButton cmdButton 
      Height          =   405
      Index           =   0
      Left            =   15
      Picture         =   "frmmain.frx":07C9
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "NEW"
      ToolTipText     =   "New Project"
      Top             =   450
      Width           =   450
   End
   Begin VB.ListBox lstUrls 
      Height          =   3150
      IntegralHeight  =   0   'False
      Left            =   495
      TabIndex        =   5
      Top             =   450
      Width           =   2610
   End
   Begin VB.PictureBox pTitle 
      AutoRedraw      =   -1  'True
      BackColor       =   &H009B6900&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000D&
      Height          =   345
      Index           =   0
      Left            =   15
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   206
      TabIndex        =   0
      Top             =   75
      Width           =   3090
      Begin VB.Label lblTitle1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current Tabs ::"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   75
         TabIndex        =   1
         Top             =   60
         Width           =   1230
      End
   End
   Begin VB.Line ln3d 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   480
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Line ln3d 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   0
      X2              =   480
      Y1              =   15
      Y2              =   15
   End
   Begin VB.Label lblFontStyle 
      AutoSize        =   -1  'True
      Caption         =   "Font Style"
      Height          =   195
      Left            =   60
      TabIndex        =   29
      Top             =   6735
      Width           =   705
   End
   Begin VB.Label lblFloat 
      AutoSize        =   -1  'True
      Caption         =   "Float"
      Height          =   195
      Left            =   60
      TabIndex        =   27
      Top             =   6345
      Width           =   345
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Text Hover Color"
      Height          =   195
      Index           =   4
      Left            =   420
      TabIndex        =   24
      Top             =   5580
      Width           =   1200
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Text Color"
      Height          =   195
      Index           =   3
      Left            =   420
      TabIndex        =   22
      Top             =   5325
      Width           =   720
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Tab Back Color"
      Height          =   195
      Index           =   2
      Left            =   420
      TabIndex        =   17
      Top             =   4620
      Width           =   1110
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Tab Hover Color"
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   16
      Top             =   4395
      Width           =   1170
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Tab Normal Color"
      Height          =   195
      Index           =   0
      Left            =   420
      TabIndex        =   14
      Top             =   4140
      Width           =   1230
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FirstRun As Boolean
Private Normal As cRgb
Private Hover As cRgb
Private HyperLinks As New Collection
Private TempFiles(2) As String

Private Sub Export(ByVal FolderSrc As String)
    'Refresh the page
    Call RefreshPage
    
    FileCopy TempFiles(0), FolderSrc & "Left.gif"
    FileCopy TempFiles(1), FolderSrc & "Right.gif"
    FileCopy TempFiles(2), FolderSrc & "index.htm"
    
End Sub

Private Function LoadProject(ByVal Filename As String) As Integer
Dim fp As Long
Dim Counter As Integer
Dim sPos As Integer
Dim sLine As String

    'load the project
    fp = FreeFile
    Open Filename For Binary As #fp
        Get #fp, , Project
        If (Project.ID <> "CTM") Then
            Exit Function
        Else
            Set HyperLinks = Nothing
            lstUrls.Clear
            'Set colors / Fonts
            pColor(0).BackColor = Project.Color1
            pColor(1).BackColor = Project.Color2
            pColor(2).BackColor = Project.Color3
            pColor(3).BackColor = Project.Color4
            pColor(4).BackColor = Project.Color5
            cboFloat.ListIndex = Project.TabFloat
            cboFontStyle.ListIndex = Project.fStyle
        
            'Load the tab items
            For Counter = 0 To UBound(Project.Items) - 1
                sLine = Project.Items(Counter)
                sPos = InStr(1, sLine, Chr(0), vbBinaryCompare)
                
                If (sPos > 0) Then
                    'Add url name
                    lstUrls.AddItem Left(sLine, sPos - 1)
                    'Add address
                    HyperLinks.Add Mid(sLine, sPos + 1)
                End If
            Next Counter
        End If
     Close #fp
    'Refresh the page
    Call RefreshPage
    LoadProject = 1
End Function

Private Sub SaveProject(ByVal Filename As String)
Dim Counter As Integer
Dim fp As Long

    'Build the project file
    Project.ID = "CTM"
    Project.Color1 = pColor(0).BackColor
    Project.Color2 = pColor(1).BackColor
    Project.Color3 = pColor(2).BackColor
    Project.Color4 = pColor(3).BackColor
    Project.Color5 = pColor(4).BackColor
    Project.TabFloat = cboFloat.ListIndex
    Project.fStyle = cboFontStyle.ListIndex
    'Redim the array to hold the items
    ReDim Preserve Project.Items(lstUrls.ListCount) As String
    'Store the items in the array
    For Counter = 0 To (lstUrls.ListCount - 1)
        Project.Items(Counter) = lstUrls.List(Counter) & Chr(0) & HyperLinks(Counter + 1)
    Next Counter
    
    fp = FreeFile
    
    'Save the project
    Open Filename For Binary As #fp
        Put #fp, , Project
    Close #fp
End Sub

Private Function GetDLGName(Optional ShowOpen As Boolean = True, Optional Title As String = "Open")
On Error GoTo CanErr:
        
    With CD1
        .CancelError = True
        .DialogTitle = Title
        .Filter = "DM CSS Tabs(*.ctm)|*.ctm|"
        
        If (ShowOpen) Then
            .ShowOpen
        Else
            .ShowSave
        End If
        
        GetDLGName = .Filename
        .Filename = vbNullString
    End With
    
    Exit Function
CanErr:

    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub PicBorder(TPicture As PictureBox)
    TPicture.Line (0, 0)-(TPicture.ScaleWidth - 1, TPicture.ScaleHeight - 1), vbWhite, B
    TPicture.Refresh
End Sub

Private Function ColorFromDLG() As Long
On Error GoTo CanErr:
    'Return color from dialog
    With CD1
        .CancelError = True
        .ShowColor
        ColorFromDLG = .Color
    End With
    
    Exit Function
CanErr:
    If (Err.Number = cdlCancel) Then
        ColorFromDLG = -1
    End If
End Function

Private Sub RefreshPage()
Dim fp As Long
Dim sBuff As String
Dim Counter As Integer
Dim sLine As String
Dim StrFinal As String
Dim RgbTmp As cRgb
Dim sHex As String

    'Make the temp files
    Call MakeTempFiles
    'Open the Temp html page
    fp = FreeFile
    Open TempFiles(2) For Binary As #fp
        sBuff = Space(LOF(fp))
        Get #fp, , sBuff
    Close #fp
    
    'Tab backcolor
    Call Long2Rgb(pColor(2).BackColor, RgbTmp)
    'Convert to Hex
    sHex = "#" & RgbToHex(RgbTmp.Red, RgbTmp.Green, RgbTmp.Blue)
    sBuff = Replace(sBuff, "$BK_Col", sHex, , , vbTextCompare)
    'Tab Forecolor
    Call Long2Rgb(pColor(3).BackColor, RgbTmp)
    sHex = "#" & RgbToHex(RgbTmp.Red, RgbTmp.Green, RgbTmp.Blue)
    sBuff = Replace(sBuff, "$TextCol", sHex, , , vbTextCompare)
    'Text hover color
    Call Long2Rgb(pColor(4).BackColor, RgbTmp)
    sHex = "#" & RgbToHex(RgbTmp.Red, RgbTmp.Green, RgbTmp.Blue)
    sBuff = Replace(sBuff, "$TextHover", sHex, , , vbTextCompare)
    'Set the align of the tabs
    sBuff = Replace(sBuff, "$Float", cboFloat.Text, , , vbTextCompare)
    'Set Text Style
    sBuff = Replace(sBuff, "$FStyle", cboFontStyle.Text, , , vbTextCompare)
    
    'Build the Links
    For Counter = 1 To HyperLinks.Count
        'Build each Link
        If (lstUrls.Selected(Counter - 1)) Then
            sLine = "<li id=" & VBQuote & "current" & VBQuote & "><a href=" & VBQuote & _
            HyperLinks(Counter) & VBQuote & "><span>" & lstUrls.List(Counter - 1) & "</span></a></li>"
        Else
            sLine = "<li><a href=" & VBQuote & HyperLinks(Counter) & VBQuote & _
            "><span>" & lstUrls.List(Counter - 1) & "</span></a></li>"
        End If
        'Build the final string
        StrFinal = StrFinal & sLine & vbCrLf
    Next Counter
    'Write to the temp html file
    sBuff = Replace(sBuff, "<!--Links-->", StrFinal, , , vbTextCompare)
    Open TempFiles(2) For Output As #fp
        Print #fp, sBuff
    Close #fp
    'Here we need to edit the two gif files and add the colors
    '
    Call Long2Rgb(pColor(0).BackColor, Normal)
    Call Long2Rgb(pColor(1).BackColor, Hover)
    'Write to the First and second gif file
    Call WriteToGif(TempFiles(0))
    Call WriteToGif(TempFiles(1))
    
    'Display the Html page
    WebView.Navigate FixPath(App.Path) & "index.htm"
    'Clear up
    sBuff = vbNullString
    sLine = vbNullString
    StrFinal = vbNullString
End Sub

Private Sub MakeTempFiles()
On Error Resume Next
    'Remove the old ones if they are found
    Call RemoveTempFiles
    'Copy the files over to the main folder to work with
    FileCopy FixPath(App.Path) & "data\Left.gif", TempFiles(0)
    FileCopy FixPath(App.Path) & "data\Right.gif", TempFiles(1)
    FileCopy FixPath(App.Path) & "data\index.htm", TempFiles(2)
End Sub

Private Sub RemoveTempFiles()
Dim Counter As Integer
On Error Resume Next
    'Remove the temp files if they are found
    For Counter = 0 To 2
        If FindFile(TempFiles(Counter)) Then
            Kill TempFiles(Counter)
        End If
    Next Counter
End Sub

Private Sub FillExamples()
    'Fill in with some examples
    lstUrls.Clear
    Set HyperLinks = Nothing
    'Add Hyperlinks Names
    lstUrls.AddItem "Products"
    lstUrls.AddItem "Download"
    lstUrls.AddItem "Purchase"
    lstUrls.AddItem "Support"
    lstUrls.AddItem "Company"
    'Add Address
    HyperLinks.Add "Products.htm"
    HyperLinks.Add "Download.htm"
    HyperLinks.Add "Purchase.htm"
    HyperLinks.Add "Support.htm"
    HyperLinks.Add "Company.htm"
End Sub

Private Sub WriteToGif(ByVal Filename As String)
Dim fp As Long
On Error Resume Next
    'This is a quick hack on how to chnage the colors
    'Note don't edit the offsets or the colors may not be displayd correct.
    fp = FreeFile
    Open Filename For Binary As #fp
        'Put the normal color
        Put #fp, 14, Normal.Red
        Put #fp, 15, Normal.Green
        Put #fp, 16, Normal.Blue
        'Now lets do the hover color part
        Put #fp, 17, Hover.Red
        Put #fp, 18, Hover.Green
        Put #fp, 19, Hover.Blue
    Close #fp
End Sub

Private Sub cboFloat_Click()
    'Refresh the page
    Call RefreshPage
End Sub

Private Sub cboFontStyle_Click()
    'Refresh the page
    Call RefreshPage
End Sub

Private Sub cmdButton_Click(Index As Integer)
Dim ColA As New Collection
Dim ColB As New Collection
Dim Counter As Integer
Dim lFile As String

    Select Case cmdButton(Index).Tag
        Case "NEW"
            'Start a new project
            If MsgBox("Are you sure you want to start a new project.", vbYesNo Or vbQuestion) = vbYes Then
                'Start new project
                cmdButton(4).Enabled = False
                cmdButton(5).Enabled = False
                Set HyperLinks = Nothing
                lstUrls.Clear
                'Refrsh the tabs
                Call RefreshPage
            End If
        Case "OPEN"
            'Load project
            lFile = GetDLGName()
            If Len(lFile) Then
                If Not LoadProject(lFile) = 1 Then
                    MsgBox "There was an error loading the project.", vbInformation, "Project Load Error"
                End If
            End If
        Case "SAVE"
            lFile = GetDLGName(False, "Save")
            If Len(lFile) Then
                Call SaveProject(lFile)
            End If
        Case "ADD"
            'Add new item
            EditOP = 0
            frmAdd.Show vbModal, Frmmain
            If (ButtonPress = vbOK) Then
                'Add the Item
                lstUrls.AddItem mUrlName
                HyperLinks.Add mUrlAddress
                'Refrsh the tabs
                Call RefreshPage
            End If
        Case "EDIT"
            'Edit Item
            EditOP = 1
            'Store the list index
            mUrlIndex = lstUrls.ListIndex
            'Store Item Info
            mUrlName = lstUrls.Text
            mUrlAddress = HyperLinks(mUrlIndex + 1)
            'Show the edit form
            frmAdd.Show vbModal, Frmmain
            '
            If (ButtonPress = vbOK) Then
                For Counter = 1 To HyperLinks.Count
                    If (mUrlIndex + 1) = Counter Then
                        'Add Urlname
                        Call ColA.Add(mUrlName)
                        'Add Address
                        Call ColB.Add(mUrlAddress)
                    Else
                        'Add Urlname
                        Call ColA.Add(lstUrls.List(Counter - 1))
                        'Add address
                        Call ColB.Add(HyperLinks(Counter))
                    End If
                Next Counter
                
                'Fill the listbox
                lstUrls.Clear
                For Counter = 1 To ColA.Count
                    lstUrls.AddItem ColA(Counter)
                Next Counter
                
                'Select the index we edited
                lstUrls.ListIndex = mUrlIndex
                Set HyperLinks = ColB
                Set ColA = Nothing
                Set ColB = Nothing
                'Refrsh the tabs
                Call RefreshPage
            End If
        Case "DELETE"
            'Delete Item
            If MsgBox("Are you sure you want to delete this item.", vbYesNo Or vbQuestion) = vbYes Then
                'Delete form collection
                Call HyperLinks.Remove(lstUrls.ListIndex + 1)
                'Delete from listbox
                lstUrls.RemoveItem lstUrls.ListIndex
                'Enable/disable edit and delete buttons
                cmdButton(4).Enabled = False
                cmdButton(5).Enabled = False
                'Refrsh the tabs
                Call RefreshPage
            End If
        Case "RELOAD"
            'Refrsh the tabs
            Call RefreshPage
        Case "EXPORT"
            lFile = FixPath(GetFolder(Frmmain.hWnd, "Export"))
            If (lFile <> "\") Then
                'Export the source files
                Call Export(lFile)
            End If
    End Select
    
    lFile = vbNullString
    ButtonPress = vbCancel
End Sub

Private Sub Form_Activate()
    If (Not FirstRun) Then
        lstUrls.ListIndex = 1
        cboFloat.ListIndex = 0
        FirstRun = True
    End If
End Sub

Private Sub Form_Load()
    'Temp Files
    TempFiles(0) = FixPath(App.Path) & "Left.gif"
    TempFiles(1) = FixPath(App.Path) & "Right.gif"
    TempFiles(2) = FixPath(App.Path) & "index.htm"
    'Add the fefault examples
    Call FillExamples
    'Refresh the page
    Call RefreshPage
    Call PicBorder(pTitle(0))
    Call PicBorder(pTitle(1))
    Call PicBorder(pTitle(2))
    Call PicBorder(pTitle(3))
    'Add align items
    cboFloat.AddItem "Left"
    cboFloat.AddItem "Right"
    'Add font styles
    cboFontStyle.AddItem "None"
    cboFontStyle.AddItem "Bold"
    cboFontStyle.AddItem "Italic"
    cboFontStyle.AddItem "Bold Italic"
    cboFontStyle.ListIndex = 1
End Sub

Private Sub Form_Resize()
    WebView.Width = (Frmmain.ScaleWidth - WebView.Left)
    WebView.Height = (Frmmain.ScaleHeight - pBar.Height - WebView.Top) - 30
    
    ln3d(2).X2 = Frmmain.ScaleWidth
    ln3d(3).X2 = Frmmain.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call RemoveTempFiles
    Set Frmmain = Nothing
End Sub

Private Sub lblInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call pColor_MouseDown(Index, Button, Shift, X, Y)
End Sub

Private Sub lstUrls_Click()
    cmdButton(4).Enabled = True
    cmdButton(5).Enabled = True
    'Refresh the page
    Call RefreshPage
End Sub

Private Sub lstUrls_DblClick()
    'Edit Item
    Call cmdButton_Click(4)
End Sub

Private Sub mnuAbout_Click()
    MsgBox Frmmain.Caption & " v1.0" & vbCrLf & vbTab & " By DreamVB" _
    & vbCrLf & vbTab & vbTab & "Please vote if you like this code", vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
    Unload Frmmain
End Sub

Private Sub pBar_Resize()
    ln3d(0).X2 = pBar.ScaleWidth
    ln3d(1).X2 = pBar.ScaleWidth
End Sub

Private Sub pColor_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iCol As Long
    If (Button = vbLeftButton) Then
        iCol = ColorFromDLG
        If (iCol <> -1) Then
            pColor(Index).BackColor = iCol
        End If
        'Refresh the page
        Call RefreshPage
    End If
End Sub

