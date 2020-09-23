VERSION 5.00
Begin VB.Form frmeditor 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Nick's 3D Maze"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9705
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   9705
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
      Left            =   7920
      ScaleHeight     =   630
      ScaleWidth      =   1530
      TabIndex        =   51
      Top             =   6720
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Timer timermove 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1560
      Top             =   7320
   End
   Begin VB.PictureBox picfruit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1530
      Left            =   8520
      ScaleHeight     =   1530
      ScaleWidth      =   1530
      TabIndex        =   33
      Top             =   4320
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox picghosts 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   630
      Left            =   7920
      ScaleHeight     =   630
      ScaleWidth      =   1530
      TabIndex        =   32
      Top             =   5880
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox piclifts 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   5160
      ScaleHeight     =   1110
      ScaleWidth      =   2550
      TabIndex        =   31
      Top             =   7320
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.PictureBox picfloor 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   2400
      ScaleHeight     =   1455
      ScaleWidth      =   2640
      TabIndex        =   30
      Top             =   6960
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox picwall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   4920
      ScaleHeight     =   1455
      ScaleWidth      =   2640
      TabIndex        =   29
      Top             =   7560
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.PictureBox picdest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   2160
      ScaleHeight     =   6495
      ScaleWidth      =   6495
      TabIndex        =   15
      Top             =   0
      Width           =   6495
      Begin VB.Frame frameload 
         BackColor       =   &H00E0E0E0&
         Height          =   5655
         Left            =   2040
         TabIndex        =   54
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         Begin VB.CommandButton cmdstart 
            Caption         =   "Load Maze"
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
            TabIndex        =   57
            Top             =   4440
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
            TabIndex        =   56
            Top             =   5040
            Width           =   2895
         End
         Begin VB.FileListBox File1 
            Height          =   3210
            Left            =   120
            TabIndex        =   55
            Top             =   1080
            Width           =   2895
         End
         Begin VB.Label Label14 
            BackColor       =   &H00E0E0E0&
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
            TabIndex        =   58
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame framesave 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Save Game"
         Height          =   1575
         Left            =   3240
         TabIndex        =   59
         Top             =   2280
         Visible         =   0   'False
         Width           =   3015
         Begin VB.CommandButton cmdcancelsave 
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
            Left            =   1560
            TabIndex        =   64
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdsavemap 
            Caption         =   "Save"
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
            TabIndex        =   63
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtsavename 
            Height          =   285
            Left            =   120
            TabIndex        =   60
            Top             =   600
            Width           =   2295
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Map Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label15 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   ".MAP"
            Height          =   285
            Left            =   2400
            TabIndex        =   61
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame framesetupmap 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Setup New Map"
         Height          =   2055
         Left            =   2880
         TabIndex        =   16
         Top             =   2520
         Width           =   3735
         Begin VB.TextBox txtfloors 
            Height          =   285
            Left            =   960
            TabIndex        =   27
            Text            =   "1"
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox txtheight 
            Height          =   285
            Left            =   960
            TabIndex        =   26
            Text            =   "40"
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtwidth 
            Height          =   285
            Left            =   960
            TabIndex        =   25
            Text            =   "40"
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtmapname 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1800
            TabIndex        =   23
            Top             =   600
            Width           =   1815
         End
         Begin VB.CommandButton cmdexitsetup 
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2640
            TabIndex        =   18
            Top             =   1560
            Width           =   975
         End
         Begin VB.CommandButton cmdcreatemap 
            Caption         =   "Create Map"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   17
            Top             =   1560
            Width           =   1575
         End
         Begin VB.CommandButton cmdloadmap 
            Caption         =   "Load Map"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label13 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Large maps may take a few seconds to make."
            Height          =   495
            Left            =   1800
            TabIndex        =   28
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label lblentermapname 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Map Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1800
            TabIndex        =   24
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Floors"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   22
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Height"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   21
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Width"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame framehelp 
         BackColor       =   &H00E0E0E0&
         Height          =   5295
         Left            =   360
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   6135
         Begin VB.CommandButton cmdok 
            Caption         =   "OK"
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
            Left            =   2640
            TabIndex        =   66
            Top             =   4680
            Width           =   1335
         End
         Begin VB.Label Label17 
            BackColor       =   &H00E0E0E0&
            Caption         =   $"frmeditor.frx":0000
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4455
            Left            =   120
            TabIndex        =   67
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.PictureBox pickey 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   0
         ScaleHeight     =   2265
         ScaleWidth      =   1545
         TabIndex        =   38
         Top             =   0
         Visible         =   0   'False
         Width           =   1575
         Begin VB.CommandButton cmdclosekey 
            Caption         =   "X"
            Height          =   255
            Left            =   1320
            TabIndex        =   71
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pickstart 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FF00FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   49
            Top             =   1200
            Width           =   255
         End
         Begin VB.PictureBox picklift 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H0000C000&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   47
            Top             =   840
            Width           =   255
         End
         Begin VB.PictureBox pickghost 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   120
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   46
            Top             =   2040
            Width           =   135
         End
         Begin VB.PictureBox pickorange 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000080FF&
            ForeColor       =   &H80000008&
            Height          =   135
            Left            =   120
            ScaleHeight     =   105
            ScaleWidth      =   105
            TabIndex        =   41
            Top             =   1680
            Width           =   135
         End
         Begin VB.PictureBox pickwall 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H000000FF&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   40
            Top             =   480
            Width           =   255
         End
         Begin VB.PictureBox pickfloor 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C0C000&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   39
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label9 
            BackColor       =   &H00FF0000&
            Caption         =   "Start"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   50
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FF0000&
            Caption         =   "Lift"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   48
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label7 
            BackColor       =   &H00FF0000&
            Caption         =   "Ghost"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   45
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00FF0000&
            Caption         =   "Orange"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   480
            TabIndex        =   44
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF0000&
            Caption         =   "Wall"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   43
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FF0000&
            Caption         =   "Floor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            TabIndex        =   42
            Top             =   120
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox picland 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2790
      Left            =   4320
      Picture         =   "frmeditor.frx":03E6
      ScaleHeight     =   2760
      ScaleWidth      =   3450
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.PictureBox picmenuback 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   7170
      ScaleWidth      =   2145
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      Begin VB.CommandButton cmdfloort 
         Caption         =   "Floor Types"
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
         TabIndex        =   70
         Top             =   3840
         Width           =   1935
      End
      Begin VB.CommandButton cmdwallt 
         Caption         =   "Wall Types"
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
         TabIndex        =   69
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "Add Others"
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
         TabIndex        =   68
         Top             =   4800
         Width           =   1935
      End
      Begin VB.CommandButton cmddown 
         Caption         =   "DOWN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   53
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton cmdup 
         Caption         =   "UP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   52
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton cmdkey 
         Caption         =   "Key"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   1560
         Width           =   1935
      End
      Begin VB.CommandButton cmdhelp 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   1935
      End
      Begin VB.CommandButton cmdsave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdload 
         Caption         =   "Load"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdnew 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.PictureBox picctype 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   360
         ScaleHeight     =   1275
         ScaleWidth      =   1395
         TabIndex        =   6
         Top             =   5760
         Width           =   1455
      End
      Begin VB.CommandButton cmdleft 
         Caption         =   "<"
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   5880
         Width           =   255
      End
      Begin VB.CommandButton cmdright 
         Caption         =   ">"
         Height          =   1095
         Left            =   1800
         TabIndex        =   8
         Top             =   5880
         Width           =   255
      End
      Begin VB.Label lblzpos 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Z"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblctype 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Floor Types"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   14
         Top             =   5280
         Width           =   1935
      End
      Begin VB.Label lblypos 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Y"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3240
         Width           =   255
      End
      Begin VB.Label lblxpos 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "X"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   255
      End
      Begin VB.Label lblmapname 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Map Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmeditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim remove As Integer
Dim editoron As Byte
Dim editorstate As Integer
Dim splaced As Byte
Dim dotemp As Byte
Dim oldx As Long
Dim tempmapinfo() As Long
Dim oldy As Long
Dim newx As Long
Dim newy As Long
Dim oldpos As Long
Dim newpos As Long
Dim cselected As Integer
Dim lefton As Byte
Dim righton As Byte
Dim upon As Byte
Dim downon As Byte

Sub refresheditor()
picdest.Cls
xpos = 0
ypos = 0
 
 For a = 1 To mapheight Step 1
    For b = 1 To mapwidth Step 1
      If a = 1 Or b = 1 Or a = mapheight Or b = mapwidth Then
        mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 2
        tempmapinfo(((a - 1) * mapwidth) + b) = 2
      End If
      If dotemp = 0 Then
        If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, pickfloor.hdc, 0, 0, SRCCOPY)
        If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 2 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, pickwall.hdc, 0, 0, SRCCOPY)
        If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 5 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, picklift.hdc, 0, 0, SRCCOPY)
        If mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = 6 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, pickstart.hdc, 0, 0, SRCCOPY)
      ElseIf dotemp = 1 Then
        If tempmapinfo(((a - 1) * mapwidth) + b) = 1 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, pickfloor.hdc, 0, 0, SRCCOPY)
        If tempmapinfo(((a - 1) * mapwidth) + b) = 2 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, pickwall.hdc, 0, 0, SRCCOPY)
        If tempmapinfo(((a - 1) * mapwidth) + b) = 5 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, picklift.hdc, 0, 0, SRCCOPY)
        If tempmapinfo(((a - 1) * mapwidth) + b) = 6 Then success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17), 17, 17, pickstart.hdc, 0, 0, SRCCOPY)
      End If
      If nfruit > 0 Then
        For cycle = 1 To nfruit Step 1
          If fruitx(cycle) = b And fruity(cycle) = a And viewzpos = fruitz(cycle) Then
              success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17), ypos - ((viewypos - 1) * 17) + 4, 9, 9, pickorange.hdc, 0, 0, SRCCOPY)
          End If
        Next cycle
      End If
      If nghosts > 0 Then
        For cycle = 1 To nghosts Step 1
          If ghostx(cycle) = b And ghosty(cycle) = a And viewzpos = ghostz(cycle) Then
            success = BitBlt(picdest.hdc, xpos - ((viewxpos - 1) * 17) + 8, ypos - ((viewypos - 1) * 17) + 4, 9, 9, pickghost.hdc, 0, 0, SRCCOPY)
          End If
        Next cycle
      End If
      xpos = xpos + 17
    Next b
    xpos = 0
    ypos = ypos + 17

  Next a
  refreshnow = 0
  
  
  
  picdest.Refresh
End Sub

Private Sub cmdadd_Click()
  lblctype.Caption = "Orange"
  cselected = 3
  picctype.Cls
  success = BitBlt(picctype.hdc, (picctype.Width - (picfruit.Width / 2)) / 30, (picctype.Height - picfruit.Height) / 30, (picfruit.Width / 30), picfruit.Height / 15, picfruit.hdc, 0, 0, SRCCOPY)
  picctype.Refresh

End Sub

Private Sub cmdcancel_Click()
If editorstate = 1 Then
  frameload.Visible = False
  framesetupmap.Visible = True
ElseIf editorstate = 2 Then
  frameload.Visible = False
  picmenuback.Visible = True
  Call refresheditor
  editoron = 1
End If
End Sub

Private Sub cmdcancelsave_Click()
  framesave.Visible = False
  picmenuback.Visible = True
  Call refresheditor
  editoron = 1

End Sub

Private Sub cmdclosekey_Click()
pickey.Visible = False
End Sub

Private Sub cmdcreatemap_Click()
splaced = 0
  If txtmapname.Text = "" Then
    MsgBox "Please enter a name for your map", vbOKOnly, "Maze Editor"
    txtmapname.SetFocus
    Exit Sub
  End If
  If IsNumeric(txtwidth.Text) = True Then
    If txtwidth.Text < 10 Or txtwidth.Text > 200 Then
      MsgBox "The map width and height must be an integer between 10 and 200", vbOKOnly, "Maze Editor"
      txtwidth.SetFocus
      Exit Sub
    End If
  Else
    MsgBox "The map width and height must be an integer between 10 and 200", vbOKOnly, "Maze Editor"
    txtwidth.SetFocus
    Exit Sub
  End If
   If IsNumeric(txtheight.Text) = True Then
    If txtheight.Text < 10 Or txtheight.Text > 200 Then
      MsgBox "The map width and height must be an integer between 10 and 200", vbOKOnly, "Maze Editor"
      txtheight.SetFocus
      Exit Sub
    End If
  Else
    MsgBox "The map width and height must be an integer between 10 and 200", vbOKOnly, "Maze Editor"
    txtheight.SetFocus
    Exit Sub
  End If
   If IsNumeric(txtfloors.Text) = True Then
    If txtfloors.Text < 1 Or txtfloors.Text > 10 Then
      MsgBox "The number of floors must be an integer between 1 and 10", vbOKOnly, "Maze Editor"
      txtfloors.SetFocus
      Exit Sub
    End If
  Else
      MsgBox "The number of floors must be an integer between 1 and 10", vbOKOnly, "Maze Editor"
      txtfloors.SetFocus
      Exit Sub
  End If
  mapwidth = Int(txtwidth.Text)
  mapheight = Int(txtheight.Text)
  mapfloors = Int(txtfloors.Text)
  mapname = txtmapname.Text
  ReDim mapinfo(mapheight * mapwidth * mapfloors)
  ReDim tempmapinfo(mapheight * mapwidth)
  
  '**** Input data into array
  For cfloors = 1 To mapfloors
    For a = 1 To mapheight
      For b = 1 To mapwidth
          mapinfo(((a - 1) * mapwidth) + b + ((cfloors - 1) * (mapwidth * mapheight))) = 2
      Next b
    Next a
  Next cfloors
  For a = 1 To mapheight
    For b = 1 To mapwidth
        tempmapinfo(((a - 1) * mapwidth) + b) = 2
    Next b
  Next a
  picmenuback.Visible = True
  framesetupmap.Visible = False
  lblmapname.Caption = mapname
'  If mapfloors = 1 Then
'    cmdup.Enabled = False
'    cmddown.Enabled = False
'    lblzpos.Caption = "N/A"
'  End If
  nghosts = 0
  nfruit = 0
  ReDim ghostx(500)
  ReDim ghosty(500)
  ReDim ghostz(500)
  ReDim fruitx(500)
  ReDim fruity(500)
  ReDim fruitz(500)
  viewzpos = 1
  viewxpos = 1
  viewypos = 1
 Call refresheditor
 picdest.Refresh
 timermove.Enabled = True
 pickey.Visible = True
  frmeditor.SetFocus
  editoron = 1
End Sub

Private Sub cmddown_Click()
If editoron = 0 Then Exit Sub
If viewzpos = 1 Then Exit Sub
viewzpos = viewzpos - 1
lblzpos.Caption = viewzpos
   For a = 0 To mapheight - 1
      For b = 1 To mapwidth
        tempmapinfo((a * mapwidth) + b) = mapinfo((a * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight)))
      Next b
    Next a
Call refresheditor
 For a = 0 To mapheight - 1
      For b = 1 To mapwidth
        tempmapinfo((a * mapwidth) + b) = mapinfo((a * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight)))
      Next b
    Next a
End Sub

Private Sub cmdexit_Click()
If editoron = 0 Then Exit Sub
success = MsgBox("Are you sure you want to exit? You will lose all unsaved information for this map", vbYesNo, "Maze Editor")
If success = 6 Then
  frmeditor.Visible = False
  frmstart.Visible = True
  Unload frmeditor
End If
End Sub

Private Sub cmdexitsetup_Click()
frmeditor.Visible = False
frmstart.Visible = True
Unload frmeditor

End Sub

Private Sub cmdfloort_Click()
  lblctype.Caption = cmdfloort.Caption
  picctype.Cls
  success = BitBlt(picctype.hdc, (picctype.Width - (picfloor.Width / 2)) / 30, (picctype.Height - picfloor.Height) / 30, (picfloor.Width / 30), picfloor.Height / 15, picfloor.hdc, 0, 0, SRCCOPY)
  picctype.Refresh
  cselected = 1

End Sub

Private Sub cmdhelp_Click()
If editoron = 0 Then Exit Sub
editoron = 0
picdest.Cls
framehelp.Visible = True
End Sub

Private Sub cmdkey_Click()
  If pickey.Visible = True Then
    pickey.Visible = False
  Else
    pickey.Visible = True
  End If
End Sub

Private Sub cmdleft_Click()
  If cselected = 1 Then
    If ftype = 1 Then Exit Sub 'exit if on first picture
    ftype = ftype - 1
    picfloor.Picture = LoadPicture(App.Path & "\pictures\floor" & ftype & ".bmp")
    picctype.Cls 'clear picturebox for new picture
    success = BitBlt(picctype.hdc, (picctype.Width - (picfloor.Width / 2)) / 30, (picctype.Height - picfloor.Height) / 30, (picfloor.Width / 30), picfloor.Height / 15, picfloor.hdc, 0, 0, SRCCOPY)
    picctype.Refresh
  ElseIf cselected = 2 Then
    If wtype = 1 Then Exit Sub 'exit if on first picture
    wtype = wtype - 1
    picwall.Picture = LoadPicture(App.Path & "\pictures\wall" & wtype & ".bmp")
    picctype.Cls 'clear picturebox for new picture
    success = BitBlt(picctype.hdc, (picctype.Width - (picwall.Width / 2)) / 30, (picctype.Height - picwall.Height) / 30, (picwall.Width / 30), picwall.Height / 15, picwall.hdc, 0, 0, SRCCOPY)
    picctype.Refresh
  ElseIf cselected >= 3 Then
    If cselected = 3 Then Exit Sub
    cselected = cselected - 1
    picctype.Cls
    If cselected = 3 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (picfruit.Width / 2)) / 30, (picctype.Height - picfruit.Height) / 30, (picfruit.Width / 30), picfruit.Height / 15, picfruit.hdc, 0, 0, SRCCOPY)
      lblctype.Caption = "Orange"
    End If
    If cselected = 4 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (picghosts.Width / 2)) / 30, (picctype.Height - picghosts.Height) / 30, (picghosts.Width / 30), picghosts.Height / 15, picghosts.hdc, picghosts.Width / 30, 0, SRCCOPY)
      lblctype.Caption = "Ghost"
    End If
    If cselected = 5 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (piclifts.Width / 2)) / 30, (picctype.Height - piclifts.Height) / 30, (piclifts.Width / 30), piclifts.Height / 15, piclifts.hdc, 0, 0, SRCCOPY)
      lblctype.Caption = "Lift"
    End If
    If cselected = 6 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (picstart.Width / 2)) / 30, (picctype.Height - picstart.Height) / 30, (picstart.Width / 30), picstart.Height / 15, picstart.hdc, 0, 0, SRCCOPY)
      lblctype.Caption = "Player Start"
    End If
    picctype.Refresh
  End If
 
End Sub

Private Sub cmdload_Click()
If editoron = 0 Then Exit Sub
editoron = 0
picdest.Cls
editorstate = 2
frameload.Visible = True
pickey.Visible = False
picmenuback.Visible = False
End Sub

Private Sub cmdloadmap_Click()
frameload.Visible = True
framesetupmap.Visible = False
editorstate = 1
End Sub

Private Sub cmdnew_Click()
If editoron = 0 Then Exit Sub
success = MsgBox("Are you sure you want to start a new map? You will lose all unsaved information for this map", vbYesNo, "Maze Editor")
If success = 6 Then
  picdest.Cls
  picmenuback.Visible = False
  pickey.Visible = False
  framesetupmap.Visible = True
  editoron = 0
End If
End Sub

Private Sub cmdok_Click()
  framehelp.Visible = False
  picmenuback.Visible = True
  Call refresheditor
  editoron = 1

End Sub

Private Sub cmdright_Click()
  If cselected = 1 Then
    If ftype = tftype Then Exit Sub 'exit if on last picture
    ftype = ftype + 1
    picfloor.Picture = LoadPicture(App.Path & "\pictures\floor" & ftype & ".bmp")
    picctype.Cls
    success = BitBlt(picctype.hdc, (picctype.Width - (picfloor.Width / 2)) / 30, (picctype.Height - picfloor.Height) / 30, (picfloor.Width / 30), picfloor.Height / 15, picfloor.hdc, 0, 0, SRCCOPY)
    picctype.Refresh
  ElseIf cselected = 2 Then
    If wtype = twtype Then Exit Sub 'exit if on last picture
    wtype = wtype + 1
    picwall.Picture = LoadPicture(App.Path & "\pictures\wall" & wtype & ".bmp")
    picctype.Cls
    success = BitBlt(picctype.hdc, (picctype.Width - (picwall.Width / 2)) / 30, (picctype.Height - picwall.Height) / 30, (picwall.Width / 30), picwall.Height / 15, picwall.hdc, 0, 0, SRCCOPY)
    picctype.Refresh
 ElseIf cselected >= 3 Then
   If cselected = 6 Then Exit Sub
   cselected = cselected + 1
    picctype.Cls
    If cselected = 3 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (picfruit.Width / 2)) / 30, (picctype.Height - picfruit.Height) / 30, (picfruit.Width / 30), picfruit.Height / 15, picfruit.hdc, 0, 0, SRCCOPY)
      lblctype.Caption = "Orange"
    End If
    If cselected = 4 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (picghosts.Width / 2)) / 30, (picctype.Height - picghosts.Height) / 30, (picghosts.Width / 30), picghosts.Height / 15, picghosts.hdc, picghosts.Width / 30, 0, SRCCOPY)
      lblctype.Caption = "Ghost"
    End If
    If cselected = 5 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (piclifts.Width / 2)) / 30, (picctype.Height - piclifts.Height) / 30, (piclifts.Width / 30), piclifts.Height / 15, piclifts.hdc, 0, 0, SRCCOPY)
      lblctype.Caption = "Lift"
    End If
    If cselected = 6 Then
      success = BitBlt(picctype.hdc, (picctype.Width - (picstart.Width / 2)) / 30, (picctype.Height - picstart.Height) / 30, (picstart.Width / 30), picstart.Height / 15, picstart.hdc, 0, 0, SRCCOPY)
      lblctype.Caption = "Player Start"
    End If
    picctype.Refresh
 End If
End Sub

Private Sub cmdsavemap_Click()
If txtsavename.Text = "" Then
  MsgBox "You need to enter a save name!", vbOKOnly, "Save Error"
  Exit Sub
End If
  Open App.Path & "\maps\user\" & txtsavename.Text & ".map" For Output As #1 'open file
  Write #1, mapwidth, mapheight, mapfloors 'Width and Height and floors
  Write #1, mapname   'Map Name
  Write #1, ftype, wtype;   'Floor Type  and wall type
  Write #1, nfruit;    'Number of fruit
  Write #1, nghosts   'Number of ghosts
  Write #1, startxpos; 'Start x position
  Write #1, startypos; 'Start y position
  Write #1, startzpos 'Start floor
  '**** Input data into array
  For cfloors = 1 To mapfloors
    For a = 1 To mapheight
      For b = 1 To mapwidth
        If a = mapheight And b = mapwidth Then
          Write #1, mapinfo(((a - 1) * mapwidth) + b + ((cfloors - 1) * (mapwidth * mapheight)))
        Else
          Write #1, mapinfo(((a - 1) * mapwidth) + b + ((cfloors - 1) * (mapwidth * mapheight)));
        End If
      Next b
    Next a
    Next cfloors
  'Input data from Fruit/Ghost arrays
  '*** Input x,y,z info for fruit
  For a = 1 To nfruit Step 1
    Write #1, fruitx(a);
    Write #1, fruity(a);
    If mapfloors > 1 Then Write #1, fruitz(a); 'Only get info for z if more than one floor
   Next a
  '*** Input x,y,z info for ghosts
  For a = 1 To nghosts Step 1
    Write #1, ghostx(a);
    Write #1, ghosty(a);
    If mapfloors > 1 Then Write #1, ghostz(a);
  Next a
  Close #1 'Close File
  framesave.Visible = False
  MsgBox "Map saved successfully as " & txtsavename.Text & ".map", vbOKOnly, "Save successful"
  picmenuback.Visible = True
  Call refresheditor
  editoron = 1

  
End Sub

Private Sub cmdstart_Click()
If File1.filename = "" Then
  success = MsgBox("Please select a level from the list!", vbOKOnly, "Error loading custom level")
  Exit Sub
End If
  Open File1.Path & "\" & File1.filename For Input As #1
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
    ReDim tempmapinfo(mapheight * mapwidth)

  '**** Input data into array
  For cfloors = 1 To mapfloors
    For a = 1 To mapheight
      For b = 1 To mapwidth
        Input #1, mapinfo(((a - 1) * mapwidth) + b + ((cfloors - 1) * (mapwidth * mapheight)))
      Next b
    Next a
  Next cfloors
  'Set up dimensions for Fruit/Ghost arrays
  ReDim fruitt(500)
  ReDim fruitx(500)
  ReDim fruity(500)
  ReDim fruitz(500)
  ReDim ghostx(500)
  ReDim ghosty(500)
  ReDim ghostz(500)
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
  viewxpos = 1
  viewypos = 1
  viewzpos = 1
  frameload.Visible = False
  picmenuback.Visible = True
  Call refresheditor
  editoron = 1
  splaced = 1
End Sub

Private Sub cmdup_Click()
If editoron = 0 Then Exit Sub
If viewzpos = mapfloors Then Exit Sub
viewzpos = viewzpos + 1
lblzpos.Caption = viewzpos
   For a = 0 To mapheight - 1
      For b = 1 To mapwidth
        tempmapinfo((a * mapwidth) + b) = mapinfo((a * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight)))
      Next b
    Next a
Call refresheditor
 For a = 0 To mapheight - 1
      For b = 1 To mapwidth
        tempmapinfo((a * mapwidth) + b) = mapinfo((a * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight)))
      Next b
    Next a
End Sub

Private Sub cmdsave_Click()
If editoron = 0 Then Exit Sub
If splaced = 0 Then
  MsgBox "You need to place a player start location first!", vbOKOnly, "Map Error"
  Exit Sub
End If
pickey.Visible = False
editoron = 0
picdest.Cls
framesave.Visible = True
picmenuback.Visible = False
End Sub





Private Sub cmdwallt_Click()
  lblctype.Caption = cmdwallt.Caption
  picctype.Cls
  success = BitBlt(picctype.hdc, (picctype.Width - (picwall.Width / 2)) / 30, (picctype.Height - picwall.Height) / 30, (picwall.Width / 30), picwall.Height / 15, picwall.hdc, 0, 0, SRCCOPY)
  picctype.Refresh
  cselected = 2

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If editoron = 0 Then Exit Sub
  If KeyCode = 37 Then lefton = 1
  If KeyCode = 39 Then righton = 1
  If KeyCode = 40 Then downon = 1
  If KeyCode = 38 Then upon = 1
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If editoron = 0 Then Exit Sub
  If KeyCode = 37 Then lefton = 0
  If KeyCode = 39 Then righton = 0
  If KeyCode = 40 Then downon = 0
  If KeyCode = 38 Then upon = 0
End Sub

Private Sub Form_Load()
File1.Path = App.Path & "\maps\user\"
editoron = 0
  splaced = 0
  dotemp = 0
  cselected = 1
  lefton = 0
  righton = 0
  downon = 0
  upon = 0
  frmeditor.Top = 0
  frmeditor.Left = 0
  frmeditor.Width = Screen.Width
  frmeditor.Height = Screen.Height
  picmenuback.Top = 0
  picmenuback.Left = 0
  picmenuback.Height = frmeditor.Height
  picdest.Left = picmenuback.Width
  picdest.Top = 0
  picdest.Height = frmeditor.Height
  picdest.Width = frmeditor.Width - picmenuback.Width
  framesetupmap.Top = (picdest.Height - framesetupmap.Height) / 2
  framesetupmap.Left = (picdest.Width - framesetupmap.Width) / 2 - (picmenuback.Width / 2)
  frameload.Top = (picdest.Height - frameload.Height) / 2
  frameload.Left = (picdest.Width - frameload.Width) / 2 - (picmenuback.Width / 2)
  framesave.Top = (picdest.Height - framesave.Height) / 2
  framesave.Left = (picdest.Width - framesave.Width) / 2 - (picmenuback.Width / 2)
  framehelp.Top = (picdest.Height - framehelp.Height) / 2
  framehelp.Left = (picdest.Width - framehelp.Width) / 2 - (picmenuback.Width / 2)
 ftype = 1
  wtype = 1
  piclifts.Picture = LoadPicture(App.Path & "\pictures\lifts.bmp")
  picfloor.Picture = LoadPicture(App.Path & "\pictures\floor" & ftype & ".bmp")
  picwall.Picture = LoadPicture(App.Path & "\pictures\wall" & wtype & ".bmp")
  picghosts.Picture = LoadPicture(App.Path & "\pictures\ghost.bmp")
  picfruit.Picture = LoadPicture(App.Path & "\pictures\orange.bmp")
  picstart.Picture = LoadPicture(App.Path & "\pictures\start.bmp")

  success = BitBlt(picctype.hdc, (picctype.Width - (picfloor.Width / 2)) / 30, (picctype.Height - picfloor.Height) / 30, (picfloor.Width / 30), picfloor.Height / 15, picfloor.hdc, 0, 0, SRCCOPY)
  picctype.Refresh
  
 
End Sub



Private Sub picdest_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If editoron = 0 Then Exit Sub
   
  If Button = 2 Then
    If cselected > 2 Then Exit Sub
    oldx = Int(viewxpos + (X / 15 / 17))
    oldy = Int(viewypos + (Y / 15 / 17))
    For a = 0 To mapheight - 1
      For b = 1 To mapwidth
        tempmapinfo((a * mapwidth) + b) = mapinfo((a * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight)))
      Next b
    Next a

    tempmapinfo(((oldy - 1) * mapwidth) + oldx) = cselected
    dotemp = 1
    Call refresheditor
     dotemp = 0
   ElseIf Button = 1 Then
     If cselected = 1 Or cselected = 2 Or cselected = 5 Then
       oldx = Int(viewxpos + (X / 15 / 17))
       oldy = Int(viewypos + (Y / 15 / 17))
       If oldx = 1 Then Exit Sub
       If oldy = 1 Then Exit Sub
       If framesetupmap.Visible = True Then Exit Sub
       If cselected = 5 Then
         For cfloors = 1 To mapfloors Step 1
           mapinfo(((oldy - 1) * mapwidth) + (oldx) + ((cfloors - 1) * (mapwidth * mapheight))) = cselected
         Next cfloors
       Else
         mapinfo(((oldy - 1) * mapwidth) + (oldx) + ((viewzpos - 1) * (mapwidth * mapheight))) = cselected
       End If
       Call refresheditor
     ElseIf cselected = 3 Then
       oldx = Int(viewxpos + (X / 15 / 17))
       oldy = Int(viewypos + (Y / 15 / 17))
       If oldx = 1 Or oldy = 1 Or oldx = mapwidth Or oldy = mapheight Then Exit Sub
          If nfruit > 0 Then
          For cycle = 1 To nfruit Step 1
            If fruitx(cycle) = oldx And fruity(cycle) = oldy And viewzpos = fruitz(cycle) Then
              For remove = cycle To nfruit
                fruitx(remove) = fruitx(remove + 1)
                fruity(remove) = fruity(remove + 1)
                fruitz(remove) = fruitz(remove + 1)
              Next remove
              nfruit = nfruit - 1
              Call refresheditor
              Exit Sub
            End If
          Next cycle
        End If
        If mapinfo(((oldy - 1) * mapwidth) + oldx + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Then
          nfruit = nfruit + 1
          fruitx(nfruit) = oldx
          fruity(nfruit) = oldy
          fruitz(nfruit) = viewzpos
          Call refresheditor
        End If
      ElseIf cselected = 4 Then
        oldx = Int(viewxpos + (X / 15 / 17))
        oldy = Int(viewypos + (Y / 15 / 17))
        If oldx = 1 Or oldy = 1 Or oldx = mapwidth Or oldy = mapheight Then Exit Sub
        If nghosts > 0 Then
          For cycle = 1 To nghosts Step 1
            If ghostx(cycle) = oldx And ghosty(cycle) = oldy And viewzpos = ghostz(cycle) Then
              For remove = cycle To nghosts
                ghostx(remove) = ghostx(remove + 1)
                ghosty(remove) = ghosty(remove + 1)
                ghostz(remove) = ghostz(remove + 1)
              Next remove
              nghosts = nghosts - 1
              Call refresheditor
              Exit Sub
            End If
          Next cycle
        End If
        If mapinfo(((oldy - 1) * mapwidth) + oldx + ((viewzpos - 1) * (mapwidth * mapheight))) = 1 Then
          nghosts = nghosts + 1
          ghostx(nghosts) = oldx
          ghosty(nghosts) = oldy
          ghostz(nghosts) = viewzpos
          Call refresheditor
        End If
      ElseIf cselected = 6 Then
        oldx = Int(viewxpos + (X / 15 / 17))
        oldy = Int(viewypos + (Y / 15 / 17))
        If oldx = 1 Or oldy = 1 Or oldx = mapwidth Or oldy = mapheight Then Exit Sub

        If splaced = 1 Then
          mapinfo(((startypos - 1) * mapwidth) + startxpos + ((startzpos - 1) * (mapwidth * mapheight))) = 1
        End If
        splaced = 1
        mapinfo(((oldy - 1) * mapwidth) + oldx + ((viewzpos - 1) * (mapwidth * mapheight))) = cselected
        startxpos = oldx
        startypos = oldy
        startzpos = viewzpos
        Call refresheditor
      End If
   End If
End Sub

Private Sub picdest_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If editoron = 0 Then Exit Sub

If X <= 0 Then Exit Sub
If framesetupmap.Visible = True Then Exit Sub
If Button = 2 Then
  If cselected > 2 Then Exit Sub
  If (Int(viewxpos + (X / 15 / 17))) > mapwidth Then Exit Sub
  If (Int(viewypos + (Y / 15 / 17))) > mapheight Then Exit Sub
  newpos = (Int(viewypos - 1 + (Y / 15 / 17)) * mapwidth) + (Int((viewxpos) + (X / 15 / 17)))
  newx = Int(viewxpos + (X / 15 / 17))
  newy = Int(viewypos + (Y / 15 / 17))
  For a = 0 To mapheight - 1
    For b = 1 To mapwidth
      tempmapinfo((a * mapwidth) + b) = mapinfo((a * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight)))
    Next b
  Next a
  If newx >= oldx And newy >= oldy Then
    For a = oldy To newy
      For b = oldx To newx
        tempmapinfo(((a - 1) * mapwidth) + b) = cselected
      Next b
    Next a
  ElseIf newx < oldx And newy < oldy Then
    For a = newy To oldy
      For b = newx To oldx
        tempmapinfo(((a - 1) * mapwidth) + b) = cselected
      Next b
    Next a
 ElseIf newx < oldx And newy >= oldy Then
    For a = oldy To newy
      For b = newx To oldx
        tempmapinfo(((a - 1) * mapwidth) + b) = cselected
      Next b
    Next a
 ElseIf newx >= oldx And newy < oldy Then
    For a = newy To oldy
      For b = oldx To newx
        tempmapinfo(((a - 1) * mapwidth) + b) = cselected
      Next b
    Next a
  End If
  dotemp = 1
  Call refresheditor
  dotemp = 0
  picdest.Refresh
ElseIf Button = 1 Then
  If cselected = 1 Or cselected = 2 Then
    oldx = Int(viewxpos + (X / 15 / 17))
    oldy = Int(viewypos + (Y / 15 / 17))
    If oldx = 1 Then Exit Sub
    If oldy = 1 Then Exit Sub
    mapinfo(((oldy - 1) * mapwidth) + (oldx) + ((viewzpos - 1) * (mapwidth * mapheight))) = cselected
    Call refresheditor
    picdest.Refresh
  End If
End If
  
If Int(viewxpos + (X / 15 / 17)) < mapwidth + 1 Then lblxpos.Caption = Int(viewxpos + (X / 15 / 17))
If Int(viewypos + (Y / 15 / 17)) < mapheight + 1 Then lblypos.Caption = Int(viewypos + (Y / 15 / 17))
End Sub

Private Sub picdest_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If editoron = 0 Then Exit Sub

If (Int(viewxpos + (X / 15 / 17))) > mapwidth Or (Int(viewypos + (Y / 15 / 17))) > mapheight Or X <= 0 Then
 Call refresheditor

  picdest.Refresh
  

  Exit Sub
End If
If Button = 2 Then
      If cselected > 2 Then Exit Sub

  newpos = (Int(viewypos - 1 + (Y / 15 / 17)) * mapwidth) + (Int((viewxpos) + (X / 15 / 17)))
  newx = Int(viewxpos + (X / 15 / 17))
  newy = Int(viewypos + (Y / 15 / 17))
  If newx >= oldx And newy >= oldy Then
    For a = oldy To newy
      For b = oldx To newx
        mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = cselected
      Next b
    Next a
  ElseIf newx < oldx And newy < oldy Then
    For a = newy To oldy
      For b = newx To oldx
        mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = cselected
      Next b
    Next a
 ElseIf newx < oldx And newy >= oldy Then
    For a = oldy To newy
      For b = newx To oldx
        mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = cselected
      Next b
    Next a
 ElseIf newx >= oldx And newy < oldy Then
    For a = newy To oldy
      For b = oldx To newx
        mapinfo(((a - 1) * mapwidth) + b + ((viewzpos - 1) * (mapwidth * mapheight))) = cselected
      Next b
    Next a
  End If
'  Call refresheditor
  '  For a = 1 To mapheight
  '    For b = 1 To mapwidth
  '      tempmapinfo(((a - 1) * mapwidth) + b) = mapinfo((a * mapwidth) + b)
  '    Next b
  '  Next a
  Call refresheditor

  picdest.Refresh
End If
End Sub

Private Sub timermove_Timer()
If lefton = 1 And righton = 0 Then
  If viewxpos > 1 Then
    viewxpos = viewxpos - 1
    lblxpos.Caption = lblxpos.Caption - 1
    refreshnow = 1
  End If
ElseIf righton = 1 And lefton = 0 Then
  If viewxpos < mapwidth Then
    viewxpos = viewxpos + 1
   If lblxpos.Caption < mapwidth Then lblxpos.Caption = lblxpos.Caption + 1
    refreshnow = 1
  End If
End If
If upon = 1 And downon = 0 Then
  If viewypos > 1 Then
    viewypos = viewypos - 1
    refreshnow = 1
    lblypos.Caption = lblypos.Caption - 1
  End If
ElseIf downon = 1 And upon = 0 Then
  If viewypos < mapheight Then
    viewypos = viewypos + 1
    refreshnow = 1
   If lblypos.Caption < mapheight Then lblypos.Caption = lblypos.Caption + 1
  End If
End If
If refreshnow = 1 Then
  Call refresheditor
End If


End Sub
