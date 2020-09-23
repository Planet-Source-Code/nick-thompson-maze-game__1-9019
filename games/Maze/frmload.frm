VERSION 5.00
Begin VB.Form frmload 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load Maze"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   3105
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3210
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2895
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   2895
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "Start Maze"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Please select a level from the list below"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcancel_Click()
frmload.Visible = False
frmstart.Visible = True
Unload frmload
End Sub

Private Sub cmdstart_Click()
If File1.filename = "" Then
  success = MsgBox("Please select a level from the list!", vbOKOnly, "Error loading custom level")
  Exit Sub
End If
  levelpath = File1.Path & "\" & File1.filename
  level = 0
  frmmain.Visible = True
  frmload.Visible = False
  Unload frmload
End Sub

Private Sub Form_Load()
  frmload.Left = (Screen.Width - frmload.Width) / 2
  frmload.Top = (Screen.Height - frmload.Height) / 2
  File1.Path = App.Path & "\maps\user\"
  File1.Refresh
End Sub
