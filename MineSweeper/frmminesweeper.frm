VERSION 5.00
Begin VB.Form frmminesweeper 
   BackColor       =   &H00C0C0C0&
   Caption         =   "MineSweeper"
   ClientHeight    =   4380
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   3525
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrtime 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3480
      Top             =   4320
   End
   Begin VB.PictureBox picgame 
      BackColor       =   &H00E0E0E0&
      Height          =   3660
      Left            =   120
      ScaleHeight     =   3600
      ScaleWidth      =   3240
      TabIndex        =   0
      Top             =   600
      Width           =   3300
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   89
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   88
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   87
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   86
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   85
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   84
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   83
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   82
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   81
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   3240
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   80
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   79
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   78
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   77
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   76
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   75
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   74
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   73
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   72
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   2880
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   71
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   70
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   69
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   68
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   67
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   66
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   65
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   64
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   63
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2520
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   62
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   61
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   60
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   59
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   58
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   57
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   56
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   55
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   54
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2160
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   53
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   52
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   51
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   50
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   49
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   48
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   47
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   46
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   45
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   1800
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   44
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   43
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   42
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   41
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   40
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   39
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   38
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   37
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   36
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1440
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   35
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   34
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   33
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   32
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   31
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   30
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   29
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   28
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   27
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   26
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   25
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   24
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   23
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   22
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   21
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   20
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   19
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   18
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   17
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   16
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   15
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   14
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   13
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   12
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   11
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   10
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   9
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   8
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   7
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   6
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   5
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   4
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   3
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   2
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.CheckBox tile 
         Height          =   375
         Index           =   0
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   5
      Height          =   3675
      Left            =   120
      Top             =   585
      Width           =   3315
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   2040
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image imgsmile 
      Height          =   360
      Left            =   1560
      Picture         =   "frmminesweeper.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   315
   End
   Begin VB.Label lbltime 
      Alignment       =   2  'Center
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Diablo"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2160
      TabIndex        =   92
      Top             =   0
      Width           =   1245
   End
   Begin VB.Label lblBombs 
      Alignment       =   2  'Center
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Diablo"
         Size            =   24
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      TabIndex        =   91
      Top             =   0
      Width           =   1245
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnufilenew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnufilehigh 
         Caption         =   "&HighScores"
      End
      Begin VB.Menu mnuFileDifficulty 
         Caption         =   "&Difficulty"
         Begin VB.Menu mnuFiledifficultyeasy 
            Caption         =   "&Easy"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFiledifficultymedium 
            Caption         =   "&Medium"
         End
         Begin VB.Menu mnuFiledifficultyhard 
            Caption         =   "&Hard"
         End
         Begin VB.Menu mnuFileDifficultyCustom 
            Caption         =   "&Custom"
         End
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmminesweeper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'MineSweeper V2.1
'By: David Torrey
'Friday, June 6, 2003
'No Revisions
'Basic MineSweeper Game

'Declare Multi-Dimensional Array Variable Representing Board
'Declare the Difficulty Variable showing how hard it will be
'Declare amount of mines on map, based on Difficulty level.
'also, make one that checks the flag positions
Private Board(0 To 10, 0 To 11) As String
Private FlagPos(0 To 10, 0 To 11) As String
Private Difficulty As String
Private Mines As Integer
Private BlankSpots As Integer
Private FMines As Integer
Private Time As Integer
Private HighScoreNames(1 To 11) As String
Private HighScores(1 To 11) As Integer
Private HighDifficulty(1 To 11) As String
Private TempScoreName As String
Private TempScore As Integer
Private TempDifficulty As String
'determines if it's a new game
Private NewGame As Boolean
'Set constants of the difficulty level
'Those being Easy, Medium, and Hard. then there
'is custom
Const Easy As String = "Easy"
Const Medium As String = "Medium"
Const Hard As String = "Hard"
Const Custom As String = "Custom"
Const Flag As String = "Flag"

'Set constant for board info
'called "MINE"

Const Mine As String = "Mine"

'Add nine for y and subtract 1
'Subtract 1 for x

Private Sub Form_Load()
    
    'Allows for more truely random amounts of mines
    Randomize Timer
    
    'show high scores
    
    If Dir(App.Path & "\HighScore.mhs") <> "HighScores.mhs" Then
        Open App.Path & "\HighScore.mhs" For Append As #1
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Write #1, "Blank", "0", "Easy"
        Close #1
    End If
    
    'Initialize board to ""
    Dim I As Integer
    Dim J As Integer
    
    'Start out as a new game
    
    NewGame = True
    
    For I = 1 To 10
        For J = 1 To 9
            Board(J, I) = ""
            FlagPos(J, I) = ""
        Next
    Next
    
    Difficulty = Easy
    
    'Set picturebox to disabled until new game is
    'started
    
    picgame.Enabled = False

End Sub

Private Sub imgsmile_Click()
    'start new game
    Call mnufilenew_Click
End Sub

Private Sub imgsmile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgsmile.Picture = LoadPicture(App.Path & "\smilefaceo.bmp")
End Sub

Private Sub imgsmile_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgsmile.Picture = LoadPicture(App.Path & "\smileface.bmp")
End Sub



Private Sub mnuexit_Click()
    'Exit Program
    End
End Sub

Private Sub mnuFileDifficultyCustom_Click()
    'Make other menu lists not checked, and
    'make the difficulty custom
    mnuFiledifficultyeasy.Checked = False
    mnuFiledifficultymedium.Checked = False
    mnuFiledifficultyhard.Checked = False
    mnuFileDifficultyCustom.Checked = True
    
    Difficulty = Custom
End Sub

Private Sub mnuFiledifficultyeasy_Click()
    'Make other menu lists not checked, and
    'make the difficulty easy
    mnuFiledifficultyeasy.Checked = True
    mnuFiledifficultymedium.Checked = False
    mnuFiledifficultyhard.Checked = False
    mnuFileDifficultyCustom.Checked = False
    
    Difficulty = Easy
End Sub

Private Sub mnuFiledifficultyhard_Click()
    'Make other menu lists not checked, and
    'make the difficulty hard
    mnuFiledifficultyeasy.Checked = False
    mnuFiledifficultymedium.Checked = False
    mnuFiledifficultyhard.Checked = True
    mnuFileDifficultyCustom.Checked = False
    
    Difficulty = Hard
End Sub

Private Sub mnuFiledifficultymedium_Click()
    'Make other menu lists not checked, and
    'make the difficulty medium
    mnuFiledifficultyeasy.Checked = False
    mnuFiledifficultymedium.Checked = True
    mnuFiledifficultyhard.Checked = False
    mnuFileDifficultyCustom.Checked = False
    
    Difficulty = Medium
End Sub

Private Sub mnufilehigh_Click()
    Dim Q As Integer
    Dim L As Integer
    
    Open App.Path & "\HighScore.mhs" For Input As #1
        For Q = 1 To 10
            Input #1, HighScoreNames(Q), HighScores(Q), HighDifficulty(Q)
        Next
    Close #1
    
    For Q = 1 To 10
        For L = 1 To 10
            If HighScores(L) < HighScores(Q) Then
                TempScore = HighScores(Q)
                TempScoreName = HighScoreNames(Q)
                TempDifficulty = HighDifficulty(Q)
                HighScores(Q) = HighScores(L)
                HighScoreNames(Q) = HighScoreNames(L)
                HighDifficulty(Q) = HighDifficulty(L)
                HighScores(L) = TempScore
                HighScoreNames(L) = TempScoreName
                HighDifficulty(L) = TempDifficulty
            End If
        Next
    Next
    
    MsgBox "High Scores" & vbCrLf & "==========" & vbCrLf & vbCrLf & "Name" & vbTab & vbTab & "Score" & vbTab & vbTab & "Difficulty" & vbCrLf & _
    "=======================================" & vbCrLf & HighScoreNames(1) & vbTab & vbTab & HighScores(1) & vbTab & vbTab & HighDifficulty(1) & vbCrLf & _
    HighScoreNames(2) & vbTab & vbTab & HighScores(2) & vbTab & vbTab & HighDifficulty(2) & vbCrLf & _
    HighScoreNames(3) & vbTab & vbTab & HighScores(3) & vbTab & vbTab & HighDifficulty(3) & vbCrLf & _
    HighScoreNames(4) & vbTab & vbTab & HighScores(4) & vbTab & vbTab & HighDifficulty(4) & vbCrLf & _
    HighScoreNames(5) & vbTab & vbTab & HighScores(5) & vbTab & vbTab & HighDifficulty(5) & vbCrLf & _
    HighScoreNames(6) & vbTab & vbTab & HighScores(6) & vbTab & vbTab & HighDifficulty(6) & vbCrLf & _
    HighScoreNames(7) & vbTab & vbTab & HighScores(7) & vbTab & vbTab & HighDifficulty(7) & vbCrLf & _
    HighScoreNames(8) & vbTab & vbTab & HighScores(8) & vbTab & vbTab & HighDifficulty(8) & vbCrLf & _
    HighScoreNames(9) & vbTab & vbTab & HighScores(9) & vbTab & vbTab & HighDifficulty(9) & vbCrLf & _
    HighScoreNames(10) & vbTab & vbTab & HighScores(10) & vbTab & vbTab & HighDifficulty(10) & vbCrLf, , "MineSweeper V2.1 High Scores"
End Sub

Private Sub mnufilenew_Click()
    'Depending on difficulty, set a random amount
    'of mines on board in random positions.
    NewGame = True
    
    Dim MineXPos As Integer
    Dim MineYPos As Integer
    Dim I As Integer
    Dim J As Integer
      
    'Set Time
    Time = 0
    lbltime.Caption = "000"

    'Initialize board to "" and tiles

    For I = 0 To tile.UBound
        If tile(I).Value = vbChecked Then
            tile(I).Value = vbUnchecked
        End If
        tile(I).Picture = Nothing
    Next
    
    For I = 1 To 10
        For J = 1 To 9
            Board(J, I) = ""
            FlagPos(J, I) = ""
        Next
    Next
    
    
    'Set number of mines based on difficulty
    
    If Difficulty = Easy Then
        Mines = Int(Rnd * 5) + 5
    ElseIf Difficulty = Medium Then
        Mines = Int(Rnd * 5) + 10
    ElseIf Difficulty = Hard Then
        Mines = Int(Rnd * 5) + 20
    ElseIf Difficulty = Custom Then
        Do
            Mines = InputBox("Please enter the number of mines (1-30):", "Mines", "15")
        Loop Until Mines >= 1 And Mines <= 30
    End If
    'display mines on screen
    FMines = Mines
    
    If Len(LTrim(FMines)) = 1 Then
        lblBombs.Caption = "00" & FMines
    ElseIf Len(LTrim(FMines)) = 2 Then
        lblBombs.Caption = "0" & FMines
    End If
    
    Dim Number As Integer
    
    'Put all of the mines onto the board, and if
    'an area is already taken, then put it in another
    'location randomly.  Also, set up the numbers around
    'the mines so that the person knows how many mines
    'are around

    For I = 1 To Mines
        Do
            MineXPos = Int(Rnd * 9) + 1
            MineYPos = Int(Rnd * 10) + 1
        Loop Until (Board(MineXPos, MineYPos) <> Mine) And (MineXPos <> 1 And MineYPos <> 1)
        Board(MineXPos, MineYPos) = Mine
        'left
        If Board(MineXPos - 1, MineYPos) <> Mine Then
            Number = Val(Board(MineXPos - 1, MineYPos))
            Number = Number + 1
            Board(MineXPos - 1, MineYPos) = Number
        End If
        'up left
        If Board(MineXPos - 1, MineYPos - 1) <> Mine Then
            Number = Val(Board(MineXPos - 1, MineYPos - 1))
            Number = Number + 1
            Board(MineXPos - 1, MineYPos - 1) = Number
        End If
        'up
        If Board(MineXPos, MineYPos - 1) <> Mine Then
            Number = Val(Board(MineXPos, MineYPos - 1))
            Number = Number + 1
            Board(MineXPos, MineYPos - 1) = Number
        End If
        'up right
        If Board(MineXPos + 1, MineYPos - 1) <> Mine Then
            Number = Val(Board(MineXPos + 1, MineYPos - 1))
            Number = Number + 1
            Board(MineXPos + 1, MineYPos - 1) = Number
        End If
        'right
        If Board(MineXPos + 1, MineYPos) <> Mine Then
            Number = Val(Board(MineXPos + 1, MineYPos))
            Number = Number + 1
            Board(MineXPos + 1, MineYPos) = Number
        End If
        'down right
        If Board(MineXPos + 1, MineYPos + 1) <> Mine Then
            Number = Val(Board(MineXPos + 1, MineYPos + 1))
            Number = Number + 1
            Board(MineXPos + 1, MineYPos + 1) = Number
        End If
        'down
        If Board(MineXPos, MineYPos + 1) <> Mine Then
            Number = Val(Board(MineXPos, MineYPos + 1))
            Number = Number + 1
            Board(MineXPos, MineYPos + 1) = Number
        End If
        'downleft
        If Board(MineXPos - 1, MineYPos + 1) <> Mine Then
            Number = Val(Board(MineXPos - 1, MineYPos + 1))
            Number = Number + 1
            Board(MineXPos - 1, MineYPos + 1) = Number
        End If
    Next
    
    'Initialize board to "" and tiles

    For I = 0 To tile.UBound
        If tile(I).Value = vbChecked Then
            tile(I).Value = vbUnchecked
        End If
        tile(I).Picture = Nothing
    Next
    
        
    For I = 1 To 10
        For J = 1 To 9
            If Trim(Board(J, I)) = "" Then
                BlankSpots = BlankSpots + 1
            End If
        Next
    Next
    'Make the user able to click buttons
    picgame.Enabled = True
    tmrtime.Enabled = True
    NewGame = False
End Sub

'Transfer Object Index Code to Board array code

Sub GetCoordinates(Index As Integer, ByRef Xpos As Integer, ByRef Ypos As Integer)
    Dim I As Integer
    Ypos = 0
    Xpos = Index
    If Index >= 9 Then
        Do
            Ypos = Ypos + 1
            Xpos = Xpos - 9
        Loop Until Xpos <= 8
            Xpos = Xpos + 1
            Ypos = Ypos + 1
    Else
        Ypos = 1
        Xpos = Index + 1
    End If
End Sub


Private Sub tile_Click(Index As Integer)
simulation:
    If NewGame = False Then
        Dim Flagx As Integer
        Dim Flagy As Integer
             
        Call GetCoordinates(Index, Flagx, Flagy)
        
        If FlagPos(Flagx, Flagy) = Flag Then
            tile(Index).Value = vbUnchecked
            Exit Sub
        End If
        
        'this is to make it look better appearance wise
        If tile(Index).Value = vbUnchecked Then
            tile(Index).Value = vbChecked
            Exit Sub
        End If
        'Checks and sees if a flag is there
        
        
        
    End If
    'coordinates you clicked
    Dim X As Integer, Y As Integer
    
    Dim I As Integer
    Dim J As Integer
    'Get board coordinates
    Call GetCoordinates(Index, X, Y)
    'If the tile you clicked on happens to be a mine
    If Board(X, Y) = Mine Then
        tmrtime.Enabled = False
        picgame.Enabled = False
        'This is to display red background for bomb actually hit
        Dim HitIndex As Integer
        HitIndex = Index
        For I = 0 To tile.UBound
            Call GetCoordinates(I, X, Y)
            If Board(X, Y) = Mine Then
                tile(I).Picture = LoadPicture(App.Path & "\Mine.bmp")
            End If
        Next
        tile(HitIndex).Picture = LoadPicture(App.Path & "\Minehit.bmp")
    'Otherwise, if it's a number
    ElseIf Val(Board(X, Y)) = 1 Or Val(Board(X, Y)) = 2 Or Val(Board(X, Y)) = 3 Or Val(Board(X, Y)) = 4 Or Val(Board(X, Y)) = 5 Or Val(Board(X, Y)) = 6 _
    Or Val(Board(X, Y)) = 7 Or Val(Board(X, Y)) = 8 Then
        tile(Index).Picture = LoadPicture(App.Path & "\" & Board(X, Y) & ".bmp")
    ElseIf Trim(Board(X, Y)) = "" And NewGame = False Then
            'checks to see if squares around it are blank too
            'and if they are, then show them also.
            If Index = 0 Then
                If Board(X + 1, Y) <> Mine Then
                    tile(Index + 1).Value = vbChecked
                End If
                If Board(X, Y + 1) <> Mine Then
                    tile(Index + 9).Value = vbChecked
                End If
                If Board(X + 1, Y + 1) <> Mine Then
                    tile(Index + 10).Value = vbChecked
                End If
            ElseIf Index >= 1 And Index <= 7 Then
                If Board(X - 1, Y) <> Mine Then
                    tile(Index - 1).Value = vbChecked
                End If
                If Board(X + 1, Y) <> Mine Then
                    tile(Index + 1).Value = vbChecked
                End If
                If Board(X - 1, Y + 1) <> Mine Then
                    tile(Index + 8).Value = vbChecked
                End If
                If Board(X, Y + 1) <> Mine Then
                    tile(Index + 9).Value = vbChecked
                End If
                If Board(X + 1, Y + 1) <> Mine Then
                    tile(Index + 10).Value = vbChecked
                End If
            ElseIf Index = 8 Then
                If Board(X - 1, Y) <> Mine Then
                    tile(Index - 1).Value = vbChecked
                End If
                If Board(X, Y + 1) <> Mine Then
                    tile(Index + 9).Value = vbChecked
                End If
                If Board(X - 1, Y - 1) <> Mine Then
                    tile(Index + 8).Value = vbChecked
                End If
            ElseIf Index = 9 Or Index = 18 Or Index = 27 Or Index = 36 Or Index = 45 Or Index = 54 Or _
            Index = 63 Or Index = 72 Then
                If Board(X, Y - 1) <> Mine Then
                    tile(Index - 9).Value = vbChecked
                End If
                If Board(X + 1, Y) <> Mine Then
                    tile(Index + 1).Value = vbChecked
                End If
                If Board(X + 1, Y - 1) <> Mine Then
                    tile(Index - 8).Value = vbChecked
                End If
                If Board(X + 1, Y + 1) <> Mine Then
                    tile(Index + 10).Value = vbChecked
                End If
                If Board(X, Y + 1) <> Mine Then
                    tile(Index + 9).Value = vbChecked
                End If
             ElseIf Index = 81 Then
                If Board(X + 1, Y) <> Mine Then
                    tile(Index + 1).Value = vbChecked
                End If
                If Board(X, Y - 1) <> Mine Then
                    tile(Index - 9).Value = vbChecked
                End If
                If Board(X + 1, Y - 1) <> Mine Then
                    tile(Index - 8).Value = vbChecked
                End If
             ElseIf Index >= 82 And Index <= 88 Then
                If Board(X - 1, Y) <> Mine Then
                    tile(Index - 1).Value = vbChecked
                End If
                If Board(X + 1, Y) <> Mine Then
                    tile(Index + 1).Value = vbChecked
                End If
                If Board(X - 1, Y - 1) <> Mine Then
                    tile(Index - 10).Value = vbChecked
                End If
                If Board(X, Y - 1) <> Mine Then
                    tile(Index - 9).Value = vbChecked
                End If
                If Board(X + 1, Y - 1) <> Mine Then
                    tile(Index - 8).Value = vbChecked
                End If
            ElseIf Index = 89 Then
                If Board(X - 1, Y) <> Mine Then
                    tile(Index - 1).Value = vbChecked
                End If
                If Board(X, Y - 1) <> Mine Then
                    tile(Index - 9).Value = vbChecked
                End If
                If Board(X - 1, Y - 1) <> Mine Then
                    tile(Index - 10).Value = vbChecked
                End If
            ElseIf Index = 17 Or Index = 26 Or Index = 35 Or Index = 44 Or Index = 53 Or Index = 62 Or Index = 71 Or Index = 80 Then
                If Board(X, Y - 1) <> Mine Then
                    tile(Index - 9).Value = vbChecked
                End If
                If Board(X, Y + 1) <> Mine Then
                    tile(Index + 9).Value = vbChecked
                End If
                If Board(X - 1, Y - 1) <> Mine Then
                    tile(Index - 10).Value = vbChecked
                End If
                If Board(X - 1, Y + 1) <> Mine Then
                    tile(Index - 8).Value = vbChecked
                End If
                If Board(X - 1, Y) <> Mine Then
                    tile(Index - 1).Value = vbChecked
                End If
            Else
                If Board(X - 1, Y - 1) <> Mine Then
                    tile(Index - 10).Value = vbChecked
                End If
                If Board(X, Y - 1) <> Mine Then
                    tile(Index - 9).Value = vbChecked
                End If
                If Board(X + 1, Y - 1) <> Mine Then
                    tile(Index - 8).Value = vbChecked
                End If
                If Board(X - 1, Y) <> Mine Then
                    tile(Index - 1).Value = vbChecked
                End If
                If Board(X + 1, Y) <> Mine Then
                    tile(Index + 1).Value = vbChecked
                End If
                If Board(X - 1, Y + 1) <> Mine Then
                    tile(Index + 8).Value = vbChecked
                End If
                If Board(X, Y + 1) <> Mine Then
                    tile(Index + 9).Value = vbChecked
                End If
                If Board(X + 1, Y + 1) <> Mine Then
                    tile(Index + 10).Value = vbChecked
                End If
            End If
                    
    'code to find out if blocks next to you are blank
    End If
    Call CheckWin
End Sub
'checks to see if you won
Sub CheckWin()
Dim Won As Boolean
Won = False
Dim I As Integer
Dim J As Integer
Dim L As Integer
Dim X As Integer
Dim Q As Integer
Dim Y As Integer
Dim Clicked As Integer
Clicked = 0
For I = 0 To tile.UBound
    If tile(I).Value = vbChecked Then
        'add a clicked area if it has already been clicked
        Clicked = Clicked + 1
    End If
Next
Clicked = Clicked + Mines
'if all the areas except bombs are clicked, you win
If Clicked = 90 Then
    Won = True
End If
If Won = True Then
    'show cool face
    imgsmile.Picture = LoadPicture(App.Path & "\Smilefacesg.bmp")
    MsgBox "You Win!"
    tmrtime.Enabled = False
    For I = 0 To tile.UBound
            'display all the bombs
            Call GetCoordinates(I, X, Y)
            If Board(X, Y) = Mine Then
                tile(I).Picture = LoadPicture(App.Path & "\Mine.bmp")
            End If
    Next
    picgame.Enabled = False
    'Check High Score
    HighScores(11) = Time
    HighScoreNames(11) = ""
    HighDifficulty(11) = Difficulty
    
    Open App.Path & "\HighScore.mhs" For Input As #1
        For Q = 1 To 10
            Input #1, HighScoreNames(Q), HighScores(Q), HighDifficulty(Q)
        Next
    Close #1
    
    For Q = 1 To 11
        For L = 1 To 11
            If HighScores(L) < HighScores(Q) Then
                TempScore = HighScores(Q)
                TempScoreName = HighScoreNames(Q)
                TempDifficulty = HighDifficulty(Q)
                HighScores(Q) = HighScores(L)
                HighScoreNames(Q) = HighScoreNames(L)
                HighDifficulty(Q) = HighDifficulty(L)
                HighScores(L) = TempScore
                HighScoreNames(L) = TempScoreName
                HighDifficulty(L) = TempDifficulty
            End If
        Next
    Next
    'You got a highscore
    If Time <> HighScores(11) Then
        For Q = 1 To 10
            If HighScoreNames(Q) = "" Then
                HighScoreNames(Q) = InputBox("High Score! Enter your name:", "HighScore")
            End If
        Next
    End If
    
    Open App.Path & "\HighScore.mhs" For Output As #1
        For Q = 1 To 10
            Write #1, HighScoreNames(Q), HighScores(Q), HighDifficulty(Q)
        Next
    Close #1
    
    Call mnufilehigh_Click
    
End If
End Sub



Private Sub tile_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if you right click, then a flag will appear.
    If Button = 2 Then
        If tile(Index).Value <> vbChecked Then
            Dim Xes As Integer
            Dim Yes As Integer
            Call GetCoordinates(Index, Xes, Yes)
            If FlagPos(Xes, Yes) <> Flag Then
                FlagPos(Xes, Yes) = Flag
                tile(Index).Picture = LoadPicture(App.Path & "\flag.bmp")
                FMines = FMines - 1
                If Len(LTrim(FMines)) = 1 Then
                    lblBombs.Caption = "00" & FMines
                ElseIf Len(LTrim(FMines)) = 2 Then
                    lblBombs.Caption = "0" & FMines
                End If
            ElseIf FlagPos(Xes, Yes) = Flag Then
                FMines = FMines + 1
                If Len(LTrim(FMines)) = 1 Then
                    lblBombs.Caption = "00" & FMines
                ElseIf Len(LTrim(FMines)) = 2 Then
                    lblBombs.Caption = "0" & FMines
                End If
                FlagPos(Xes, Yes) = ""
                tile(Index).Picture = Nothing
            End If
        End If
    End If
End Sub

Private Sub tmrtime_Timer()
    Time = Time + 1
    If Len(LTrim(Time)) = 1 Then
        lbltime.Caption = "00" & Time
    ElseIf Len(LTrim(Time)) = 2 Then
        lbltime.Caption = "0" & Time
    Else
        lbltime.Caption = Time
    End If
    If Time >= 999 Then
        Time = 999
    End If
End Sub
