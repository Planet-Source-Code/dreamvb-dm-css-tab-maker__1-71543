VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4005
      TabIndex        =   5
      Top             =   150
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   350
      Left            =   4005
      TabIndex        =   4
      Top             =   585
      Width           =   1000
   End
   Begin VB.TextBox txtUrl 
      Height          =   345
      Left            =   675
      TabIndex        =   3
      Top             =   615
      Width           =   3180
   End
   Begin VB.TextBox txtName 
      Height          =   345
      Left            =   675
      TabIndex        =   2
      Top             =   180
      Width           =   3180
   End
   Begin VB.Label lblUrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Url:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   240
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   465
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    'Unload form
    ButtonPress = vbCancel
    Unload frmAdd
End Sub

Private Sub cmdOk_Click()
    ButtonPress = vbOK
    'Setup URL Info
    mUrlName = txtName.Text
    mUrlAddress = txtUrl.Text
    'Unload form
    Unload frmAdd
End Sub

Private Sub Form_Load()
    Set frmAdd.Icon = Nothing
    
    If (EditOP = 0) Then
        frmAdd.Caption = "Add"
    Else
        frmAdd.Caption = "Edit"
        'Setup the textboxes.
        txtName.Text = mUrlName
        txtUrl.Text = mUrlAddress
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmAdd = Nothing
End Sub

Private Sub txtName_Change()
    cmdOk.Enabled = Len(Trim(txtName.Text))
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) And (cmdOk.Enabled) Then
        Call cmdOk_Click
    End If
End Sub

Private Sub txtUrl_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) And (cmdOk.Enabled) Then
        Call cmdOk_Click
    End If
End Sub
