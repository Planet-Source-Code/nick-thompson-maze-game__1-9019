VERSION 5.00
Begin VB.Form frmstart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nick's 3D Maze"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdusermaze 
      Caption         =   "User Maze"
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
      Left            =   960
      TabIndex        =   9
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtpassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit Game"
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
      Left            =   3960
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
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
      Left            =   3960
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdinstructions 
      Caption         =   "Instructions"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1800
      Width           =   1815
   End
   Begin VB.CommandButton cmdeditor 
      Caption         =   "Maze Editor"
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
      Left            =   960
      TabIndex        =   1
      Top             =   3720
      Width           =   1815
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
      Left            =   960
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lbltitle 
      Alignment       =   2  'Center
      Caption         =   "Nick's 3D Maze"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label lblwelcome 
      Alignment       =   2  'Center
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmstart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdabout_Click()
frmabout.Visible = True
frmstart.Visible = False
Unload frmstart

End Sub

Private Sub cmdeditor_Click()
  frmeditor.Visible = True
  frmstart.Visible = False
  Unload frmstart
End Sub

Private Sub cmdexit_Click()
End

End Sub

Private Sub cmdinstructions_Click()
frminstructions.Visible = True
frmstart.Visible = False
Unload frmstart

End Sub

Private Sub cmdstart_Click()
level = 1
If txtpassword.Text = "HRSTEDRED" Then level = 2
If txtpassword.Text = "GOTOITNOW" Then level = 3
If txtpassword.Text = "STARTTHIS" Then level = 4
If txtpassword.Text = "LONGLEVEL" Then level = 5
If nlevels > 5 Then
  For cyclelevels = 6 To nlevels
    If txtpassword.Text = "LEVEL" & cyclelevels Then level = cyclelevels
  Next cyclelevels
End If
frmstart.Visible = False
frmmain.Visible = True
Unload frmstart
End Sub

Private Sub cmdusermaze_Click()
frmstart.Visible = False
frmload.Visible = True
Unload frmstart
End Sub

Private Sub Form_Load()
frmstart.Left = (Screen.Width - frmstart.Width) / 2
frmstart.Top = (Screen.Height - frmstart.Height) / 2
lblwelcome.Left = (frmstart.Width - lblwelcome.Width) / 2
lbltitle.Left = (frmstart.Width - lbltitle.Width) / 2
  Open App.Path & "\info.dat" For Input As #1 'open file
    Input #1, tftype 'Total number of types of floor
    Input #1, twtype 'Total number of types of wall
    Input #1, nlevels 'Total number of levels
  Close #1 'close file
End Sub
