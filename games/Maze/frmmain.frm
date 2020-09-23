VERSION 5.00
Begin VB.Form frmmain 
   BorderStyle     =   0  'None
   Caption         =   "Nick's 3D Maze"
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   10305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox picstart 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   8760
      ScaleHeight     =   630
      ScaleWidth      =   1530
      TabIndex        =   30
      Top             =   2760
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox picwall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   5880
      ScaleHeight     =   1455
      ScaleWidth      =   2640
      TabIndex        =   29
      Top             =   1680
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox picfloor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   5880
      ScaleHeight     =   1455
      ScaleWidth      =   2640
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Timer timerdie 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   840
      Top             =   4560
   End
   Begin VB.PictureBox piclifts 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   5880
      ScaleHeight     =   1110
      ScaleWidth      =   2550
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox picmenuback 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   0
      ScaleHeight     =   3375
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit to Menu"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2880
         Width           =   1455
      End
      Begin VB.Label lblfloor 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Floor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblmapname 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Close Encouters"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label lbllevel 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblfloors 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   9
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Floors"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1005
      End
      Begin VB.Label lblghosts 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   7
         Top             =   2160
         Width           =   495
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ghosts"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   1005
      End
      Begin VB.Label lbllives 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Lives Left"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label lblfruit 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fruits Left"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   1005
      End
   End
   Begin VB.PictureBox picghosts 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   8640
      ScaleHeight     =   630
      ScaleWidth      =   1530
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox picfruit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   8640
      ScaleHeight     =   1530
      ScaleWidth      =   1530
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox picbuggy 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3780
      Left            =   7440
      Picture         =   "frmmain.frx":0000
      ScaleHeight     =   3750
      ScaleWidth      =   2415
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Timer timerrefresh 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8880
      Top             =   5400
   End
   Begin VB.PictureBox picdest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2880
      ScaleHeight     =   1335
      ScaleWidth      =   1575
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picback 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   1680
      ScaleHeight     =   5535
      ScaleWidth      =   3975
      TabIndex        =   12
      Top             =   0
      Width           =   3975
      Begin VB.Frame framefin 
         Height          =   2895
         Left            =   0
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton cmdreturn2 
            Caption         =   "Return to Menu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   2400
            Width           =   2415
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Well Done you have completed all the levels in this game. View readme.txt for details on how to add levels."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2115
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.Frame framewelldone2 
         Height          =   1575
         Left            =   960
         TabIndex        =   31
         Top             =   3000
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton cmdreturn 
            Caption         =   "Return to Menu"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Well Done you have won this level."
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.Frame framewelldone 
         Height          =   3135
         Left            =   -240
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   2655
         Begin VB.CommandButton cmdnextlevel 
            Caption         =   "Start next level"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   2640
            Width           =   2415
         End
         Begin VB.Label lblpassword 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "HRSTEDRED"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   405
            Left            =   120
            TabIndex        =   21
            Top             =   2160
            Width           =   2400
         End
         Begin VB.Label lblpass 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "PASSWORD"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   20
            Top             =   1800
            Width           =   2400
         End
         Begin VB.Label lblwelldone 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Well Done you have won this level. Get ready for the next level.  Good Luck"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1485
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2400
         End
      End
      Begin VB.Label lblgameover 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Game Over"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   1215
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.Label lbldie 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0000C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "You died and lost a life"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   4755
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cycle As Integer     'Used in For...Next loops to cycle through things e.g. Ghosts,fruit
Dim movethem As Byte     'Move ghosts or not
Dim gameon As Byte       'Turns user control on or off
Dim oghostx As Integer   'Original ghost x position
Dim oghosty As Integer   'original ghost y position
Dim refreshnow As Byte   'Causes the graphics to refresh is set to 1
Dim fruitt() As Byte     'Has fruit been taken
Dim xpos As Long         'X position on screen for land 1,1
Dim ypos As Long         'Y position on screen for land 1,1
Sub loadmap() 'Load all information
  '****** LOAD LEVEL INFORMATION ******
  If level = 0 Then
    Open levelpath For Input As #1 'open custom file
  Else
    Open App.Path & "\maps\l" & level & ".map" For Input As #1 'open file
  End If
  Input #1, mapwidth  'Width
  Input #1, mapheight 'Height
  Input #1, mapfloors 'Number of Floors
  Input #1, mapname   'Map Name
  Input #1, ftype     'Floor Type
  Input #1, wtype     'Wall Type
  Input #1, nfruit    'Number of fruit
  Input #1, nghosts   'Number of ghosts
  Input #1, startxpos 'Start x position
  Input #1, startypos 'Start y position
  Input #1, startzpos 'Start floor
  'Set up dimensions for array
  ReDim mapinfo(mapheight * mapwidth * mapfloors)
  '**** Input data into array
  For cfloors = 1 To mapfloors
    For a = 1 To mapheight
      For b = 1 To mapwidth
        Input #1, mapinfo(((a - 1) * mapwidth) + b + ((cfloors - 1) * (mapwidth * mapheight)))
      Next b
    Next a
  Next cfloors
  'Set up dimensions for Fruit/Ghost arrays
  ReDim fruitt(nfruit)
  ReDim fruitx(nfruit)
  ReDim fruity(nfruit)
  ReDim fruitz(nfruit)
  ReDim ghostx(nghosts)
  ReDim ghosty(nghosts)
  ReDim ghostz(nghosts)
  cfruit = nfruit 'Decreased by one each time a fruit is eaten, when 0 you win
  '*** Input x,y,z info for fruit
  For a = 1 To nfruit Step 1
    Input #1, fruitx(a)
    Input #1, fruity(a)
    If mapfloors > 1 Then Input #1, fruitz(a) 'Only get infor for z if more than one floor
    If mapfloors = 1 Then fruitz(a) = 1       '                   "
    fruitt(a) = 0 'Fruit not taken yet so 0
  Next a
  '*** Input x,y,z info for ghosts
  For a = 1 To nghosts Step 1
    Input #1, ghostx(a)
    Input #1, ghosty(a)
    If mapfloors > 1 Then Input #1, ghostz(a)
    If mapfloors = 1 Then ghostz(a) = 1
  Next a
  Close #1 'Close File
  viewxpos = startxpos 'Set current X position as start position
  viewypos = startypos 'Set current Y position as start position
  viewzpos = startzpos 'Set current Z position as start position
  piclifts.Picture = LoadPicture(App.Path & "\pictures\lifts.bmp")
  picfloor.Picture = LoadPicture(App.Path & "\pictures\floor" & ftype & ".bmp")
  picwall.Picture = LoadPicture(App.Path & "\pictures\wall" & wtype & ".bmp")
  picghosts.Picture = LoadPicture(App.Path & "\pictures\ghost.bmp")
  picfruit.Picture = LoadPicture(App.Path & "\pictures\orange.bmp")
  picstart.Picture = LoadPicture(App.Path & "\pictures\start.bmp")

  Call refreshgraphics 'Draw all information on the screen
  '**** Display user information ****
    
  lbllevel.Caption = level      'Show current level (N/A if user created level)
  lblfloor.Caption = viewzpos   'Show current floor
  lblmapname.Caption = mapname  'Show map name
  lblfruit.Caption = nfruit     'Show number of fruit left
  lbllives.Caption = lives      'Show number of lives left
  lblghosts.Caption = nghosts   'Show number of ghosts
  lblfloors.Caption = mapfloors 'Show number of floors on level
  gameon = 1                    'Give user control / Start game
End Sub
Sub refreshgraphics()
restart:
xpos = ((frmmain.Width - picmenuback.Width) / 30) - (viewxpos * 42) + ((viewypos - 1) * 42)
ypos = (frmmain.Height / 30) - ((viewypos - 1) * 21) - (viewxpos * 21)
   
  For a = 1 To mapheight Step 1
    For b = 1 To mapwidth Step 1
    
      If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Then
        success = BitBlt(picdest.hdc, xpos, ypos, 84, 43, picfloor.hdc, 0, 0, SRCAND)
        success = BitBlt(picdest.hdc, xpos, ypos, 84, 43, picfloor.hdc, 84, 0, SRCPAINT)
      End If
     If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 2 Then
        success = BitBlt(picdest.hdc, xpos, ypos - 29, 84, 72, picwall.hdc, 0, 0, SRCAND)
        success = BitBlt(picdest.hdc, xpos, ypos - 29, 84, 72, picwall.hdc, 84, 0, SRCPAINT)
      End If
      If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
       success = BitBlt(picdest.hdc, xpos, ypos, 84, 72, piclifts.hdc, 0, 0, SRCAND)
       success = BitBlt(picdest.hdc, xpos, ypos, 84, 72, piclifts.hdc, 84, 0, SRCPAINT)
      End If
      If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 6 Then
        success = BitBlt(picdest.hdc, xpos, ypos, 84, 43, picstart.hdc, 0, 0, SRCAND)
        success = BitBlt(picdest.hdc, xpos, ypos, 84, 43, picstart.hdc, 84, 0, SRCPAINT)
      End If
 
      For cycle = 1 To nfruit Step 1
        If fruitx(cycle) = b And fruity(cycle) = a And viewzpos = fruitz(cycle) Then
          If viewxpos = fruitx(cycle) And viewypos = fruity(cycle) And viewzpos = fruitz(cycle) Then
            If fruitt(cycle) = 0 Then
              fruitt(cycle) = 1
              cfruit = cfruit - 1
              sndPlaySound App.Path & "\sound\eatfruit.wav", &H1
              lblfruit.Caption = cfruit
              If cfruit = 0 Then
                gameon = 0
                If level = 0 Then
                  framewelldone2.Visible = True
                  cmdreturn.SetFocus
                ElseIf level = nlevels Then framefin.Visible = True
                Else
                  If level = 1 Then lblpassword.Caption = "HRSTEDRED"
                  If level = 2 Then lblpassword.Caption = "GOTOITNOW"
                  If level = 3 Then lblpassword.Caption = "STATTHIS"
                  If level = 4 Then lblpassword.Caption = "LONGLEVEL"
                  If level > 4 Then lblpassword.Caption = "LEVEL" & level + 1
                  framewelldone.Visible = True
                  cmdnextlevel.SetFocus
                End If
              End If
             
            End If
          End If
          If fruitt(cycle) = 0 Then
            success = BitBlt(picdest.hdc, xpos + 24, ypos, 27, 29, picfruit.hdc, 0, 0, SRCAND)
            success = BitBlt(picdest.hdc, xpos + 24, ypos, 27, 29, picfruit.hdc, 27, 0, SRCPAINT)
          End If
        End If
      Next cycle
      
      For cycle = 1 To nghosts Step 1
        If ghostx(cycle) = b And ghosty(cycle) = a And viewzpos = ghostz(cycle) Then
          If viewxpos = ghostx(cycle) And viewypos = ghosty(cycle) And viewzpos = ghostz(cycle) Then
            If lives = 0 Then
              timerdie.Enabled = True
              lblgameover.Visible = True
              gameon = 0
              Exit Sub
            End If
               gameon = 0
            timerdie.Enabled = True
            lbldie.Visible = True
            sndPlaySound App.Path & "\sound\die.wav", &H1

            lives = lives - 1
            lbllives.Caption = lives
            viewxpos = startxpos
            viewypos = startypos
            viewzpos = startzpos
            picdest.Cls
            GoTo restart
          End If
            success = BitBlt(picdest.hdc, xpos + 22, ypos - 4, 40, 40, picghosts.hdc, 0, 0, SRCAND)
            success = BitBlt(picdest.hdc, xpos + 22, ypos - 4, 40, 40, picghosts.hdc, 55, 0, SRCPAINT)
          End If
      Next cycle
      
      
      If a = viewypos And b = viewxpos Then
        If buggydir = 1 Then
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 0, 0, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 81, 0, SRCPAINT)
        ElseIf buggydir = 2 Then
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 0, 61, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 81, 61, SRCPAINT)
        ElseIf buggydir = 3 Then
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 0, 124, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 81, 124, SRCPAINT)
        ElseIf buggydir = 4 Then
          success = BitBlt(picdest.hdc, xpos + 4, ypos - 12, 80, 54, picbuggy.hdc, 0, 180, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 4, ypos - 12, 80, 54, picbuggy.hdc, 81, 180, SRCPAINT)
        End If
      End If
      xpos = xpos + 42
      ypos = ypos + 21
    Next b
    xpos = xpos - (42 * (mapwidth + 1))
    ypos = ypos - (21 * (mapwidth - 1))
  Next a
  success = BitBlt(picback.hdc, 0, 0, picback.Width / 15, picback.Height / 15, picdest.hdc, 0, 0, SRCCOPY)
  

  For d = 1 To 10 Step 1
  

  Next d
picdest.Cls
End Sub


Private Sub cmdexit_Click() 'Exit to menu
  'Exit?
  success = MsgBox("Are you sure you want to exit, you will lose this game", vbYesNo, "Nicks's 3D Maze")
  'If yes then...
  If success = 6 Then
    frmmain.Visible = False
    frmstart.Visible = True
    Unload frmmain
  End If
End Sub
Private Sub cmdnextlevel_Click()
  level = level + 1 'move to next level
  framewelldone.Visible = False 'remove frame from view
  Call loadmap 'Load new level
End Sub


Private Sub cmdreturn_Click()
frmmain.Visible = False
frmstart.Visible = True
Unload frmmain
End Sub


Private Sub cmdreturn2_Click()
frmmain.Visible = False
frmstart.Visible = True
Unload frmmain

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If gameon = 0 Then Exit Sub 'Ignore all user keypresses if gameon is 0
If KeyCode = 33 Or KeyCode = 105 Then
  If mapinfo(((viewypos - 2) * mapwidth) + viewxpos + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Or mapinfo(((viewypos - 2) * mapwidth) + viewxpos + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
    buggydir = 1
    viewypos = viewypos - 1
    refreshnow = 1
  End If
ElseIf KeyCode = 34 Or KeyCode = 99 Then
  If mapinfo(((viewypos - 1) * mapwidth) + viewxpos + 1 + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Or mapinfo(((viewypos - 1) * mapwidth) + viewxpos + 1 + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
    buggydir = 2
    viewxpos = viewxpos + 1
    refreshnow = 1
  End If
ElseIf KeyCode = 35 Or KeyCode = 97 Then
  If mapinfo(((viewypos) * mapwidth) + viewxpos + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Or mapinfo(((viewypos) * mapwidth) + viewxpos + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
    buggydir = 3
    viewypos = viewypos + 1
    refreshnow = 1
  End If
ElseIf KeyCode = 36 Or KeyCode = 103 Then
  If mapinfo(((viewypos - 1) * mapwidth) + viewxpos - 1 + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Or mapinfo(((viewypos - 1) * mapwidth) + viewxpos - 1 + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
    buggydir = 4
    viewxpos = viewxpos - 1
    refreshnow = 1
  End If
ElseIf KeyCode = 38 Or KeyCode = 104 Then 'up floor
  If mapinfo(((viewypos - 1) * mapwidth) + viewxpos + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
    If mapfloors > 1 Then 'in case someone has made an error
      If viewzpos < mapfloors Then 'ensure you are not on top floor
        lblfloor.Caption = viewzpos
        viewzpos = viewzpos + 1
        refreshnow = 1
      End If
    End If
  End If
ElseIf KeyCode = 40 Or KeyCode = 98 Then 'down floor
  If mapinfo(((viewypos - 1) * mapwidth) + viewxpos + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
    If viewzpos > 1 Then 'ensure you are not on bottom floor
      viewzpos = viewzpos - 1
      lblfloor.Caption = viewzpos
      refreshnow = 1
    End If
  End If
ElseIf KeyCode = 12 Or KeyCode = 101 Then 'stand still
  refreshnow = 1
End If
If refreshnow = 1 Then
  
  For cycle = 1 To nghosts Step 1
 '  If ghostx(cycle) = b And ghosty(cycle) = a Then
     If viewxpos = ghostx(cycle) And viewypos = ghosty(cycle) And viewzpos = ghostz(cycle) Then
       If lives = 0 Then
         timerdie.Enabled = True
         lblgameover.Visible = True
         gameon = 0
         Exit Sub
       End If
    

       gameon = 0
       timerdie.Enabled = True
       lbldie.Visible = True
       lives = lives - 1
       sndPlaySound App.Path & "\sound\die.wav", &H1
     
       lbllives.Caption = lives
       viewxpos = startxpos
       viewypos = startypos
       viewzpos = startzpos
       picdest.Cls
     End If
 '  End If
 Next cycle
 If movethem = 1 Then
 For cycle = 1 To nghosts Step 1
    
    Randomize Timer
   m = Rnd
    oghostx = ghostx(cycle)
    oghosty = ghosty(cycle)
    If m >= 0 And m <= 0.2 Then
      ghostx(cycle) = ghostx(cycle) + 1
    End If
    If m > 0.2 And m <= 0.4 Then
      ghostx(cycle) = ghostx(cycle) - 1
    End If
    If m > 0.4 And m <= 0.6 Then
      ghosty(cycle) = ghosty(cycle) + 1
    End If
    If m > 0.6 And m <= 0.8 Then
      ghosty(cycle) = ghosty(cycle) - 1
    End If
      If mapinfo(((ghosty(cycle) - 1) * mapwidth) + ghostx(cycle) + ((ghostz(cycle) - 1) * (mapwidth * mapheight))) = 2 Or mapinfo(((ghosty(cycle) - 1) * mapwidth) + ghostx(cycle) + ((ghostz(cycle) - 1) * (mapwidth * mapheight))) = 5 Then
'      If mapinfo(((ghosty(cycle) - 1) * mapwidth) + ghostx(cycle) + ((viewzpos - 1) * (mapwidth * mapheight))) = 2 Or mapinfo(((ghosty(cycle) - 1) * mapwidth) + ghostx(cycle) + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then
        ghostx(cycle) = oghostx
        ghosty(cycle) = oghosty
       cycle = cycle - 1
     End If
  Next cycle
  End If
  refreshnow = 0
  
  

  Call refreshgraphics
End If

End Sub

Private Sub Form_Load()
  movethem = 1
  gameon = 0
  refreshnow = 0
  buggydir = 1
  lives = 5
  frmmain.Top = 0
  frmmain.Left = 0
  frmmain.Width = Screen.Width
  frmmain.Height = Screen.Height
  picback.Height = frmmain.Height
  picback.Top = 0
  picback.Left = 0
  picback.Width = frmmain.Width
  picdest.Width = picback.Width
  picdest.Height = picback.Height
  framewelldone.Left = (picback.Width - framewelldone.Width) / 2
  framewelldone.Top = (picback.Height - framewelldone.Height) / 2
  framewelldone2.Left = (picback.Width - framewelldone2.Width) / 2
  framewelldone2.Top = (picback.Height - framewelldone2.Height) / 2
  framefin.Left = (picback.Width - framefin.Width) / 2
  framefin.Top = (picback.Height - framefin.Height) / 2
  lbldie.Left = (picback.Width - lbldie.Width) / 2
  lbldie.Top = (picback.Height - lbldie.Height) / 2
  lblgameover.Left = (picback.Width - lblgameover.Width) / 2
  lblgameover.Top = (picback.Height - lblgameover.Height) / 2
  Call loadmap
End Sub






Private Sub picback_Paint()
  Call refreshgraphics
End Sub

Private Sub timerdie_Timer()
If lblgameover.Visible = True Then
  frmmain.Visible = False
  frmstart.Visible = True
  Unload frmmain
Else
  lbldie.Visible = False
  gameon = 1
  timerdie.Enabled = False

End If
End Sub

Private Sub timerrefresh_Timer()
Dim d As Long
Do

xpos = ((frmmain.Width - picmenuback.Width) / 30) - (viewxpos * 42) + ((viewypos - 1) * 42)
ypos = (frmmain.Height / 30) - ((viewypos - 1) * 21) - (viewxpos * 21)
  
  
  
  For a = 1 To mapheight Step 1
    For b = 1 To mapwidth Step 1
      If mapinfo(((a - 1) * mapwidth) + b) = 1 Then
        success = BitBlt(picdest.hdc, xpos, ypos, 84, 43, picland.hdc, 0, 0, SRCAND)
        success = BitBlt(picdest.hdc, xpos, ypos, 84, 43, picland.hdc, 84, 0, SRCPAINT)
      End If
     If mapinfo(((a - 1) * mapwidth) + b) = 2 Then
        success = BitBlt(picdest.hdc, xpos, ypos - 29, 84, 72, picland.hdc, 0, 43, SRCAND)
        success = BitBlt(picdest.hdc, xpos, ypos - 29, 84, 72, picland.hdc, 84, 43, SRCPAINT)
      End If
      If a = viewypos And b = viewxpos Then
        If buggydir = 1 Then
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 0, 0, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 81, 0, SRCPAINT)
        ElseIf buggydir = 2 Then
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 0, 61, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 81, 61, SRCPAINT)
        ElseIf buggydir = 3 Then
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 0, 124, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 6, ypos - 10, 80, 50, picbuggy.hdc, 81, 124, SRCPAINT)
        ElseIf buggydir = 4 Then
          success = BitBlt(picdest.hdc, xpos + 4, ypos - 12, 80, 54, picbuggy.hdc, 0, 180, SRCAND)
          success = BitBlt(picdest.hdc, xpos + 4, ypos - 12, 80, 54, picbuggy.hdc, 81, 180, SRCPAINT)
        End If
      End If
      xpos = xpos + 42
      ypos = ypos + 21
    Next b
    xpos = xpos - (42 * (mapwidth + 1))
    ypos = ypos - (21 * (mapwidth - 1))
  Next a
  success = BitBlt(picback.hdc, 0, 0, picback.Width / 15, picback.Height / 15, picdest.hdc, 0, 0, SRCCOPY)
  
  For d = 1 To 10 Step 1
  
  DoEvents
  Next d
picdest.Cls
Loop
End Sub
