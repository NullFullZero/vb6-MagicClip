VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MagicClip"
   ClientHeight    =   7170
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   11445
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start / Stop"
      Height          =   372
      Left            =   9960
      TabIndex        =   2
      Top             =   6720
      Width           =   1452
   End
   Begin VB.Timer tmrPaste 
      Interval        =   200
      Left            =   11040
      Top             =   6720
   End
   Begin VB.TextBox txtClip 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6612
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   11412
   End
   Begin VB.Label Label1 
      Caption         =   "Autor: NullFullZero"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   6720
      Width           =   3252
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next

tmrPaste.Enabled = Not tmrPaste.Enabled
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub tmrPaste_Timer()
On Error Resume Next

If Clipboard.GetText = "" Then Exit Sub

txtClip.Text = txtClip.Text & Clipboard.GetText & vbCrLf

Clipboard.Clear
End Sub
