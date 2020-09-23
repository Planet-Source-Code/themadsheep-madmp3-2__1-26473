VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form main 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4335
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7230
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   4335
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Minimize"
      Height          =   300
      Left            =   1080
      TabIndex        =   24
      Top             =   1440
      Width           =   855
   End
   Begin VB.ListBox songlist 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FFFF&
      Height          =   1200
      Left            =   2400
      TabIndex        =   22
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ListBox playlist 
      BackColor       =   &H80000007&
      Height          =   840
      Left            =   2640
      TabIndex        =   21
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton exit_but 
      Caption         =   "Exit"
      Height          =   300
      Left            =   1080
      TabIndex        =   15
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton list_but 
      Caption         =   "Show"
      Height          =   350
      Left            =   5400
      TabIndex        =   14
      Top             =   1560
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5040
      Top             =   5160
   End
   Begin VB.HScrollBar balance 
      Height          =   200
      LargeChange     =   100
      Left            =   1320
      Max             =   5000
      Min             =   -5000
      SmallChange     =   50
      TabIndex        =   8
      Top             =   3840
      Value           =   1
      Width           =   1695
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   200
      LargeChange     =   10
      Left            =   4200
      Max             =   2500
      SmallChange     =   10
      TabIndex        =   6
      Top             =   3840
      Value           =   1875
      Width           =   1695
   End
   Begin VB.Frame list_frame 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   6120
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
      Begin VB.CommandButton load_but 
         Caption         =   "Load"
         Height          =   350
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton save_but 
         Caption         =   "Save"
         Height          =   350
         Left            =   120
         TabIndex        =   19
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton clear_but 
         Caption         =   "Clear"
         Height          =   350
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton rem_but 
         Caption         =   "Remove"
         Height          =   350
         Left            =   120
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton add_but 
         Caption         =   "Add"
         Height          =   350
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox loop_but 
         BackColor       =   &H00000000&
         Caption         =   "LOOP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         MaskColor       =   &H00808080&
         TabIndex        =   13
         Top             =   3600
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FFFF&
         X1              =   120
         X2              =   960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         X1              =   120
         X2              =   960
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Caption         =   "PLAYLIST"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   4320
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.mp3"
   End
   Begin VB.Image p2 
      Height          =   420
      Left            =   4920
      Picture         =   "Form1.frx":6AE4
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image p1 
      Height          =   420
      Left            =   4320
      Picture         =   "Form1.frx":6EF1
      Top             =   5760
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image end_but 
      Height          =   420
      Left            =   4320
      Picture         =   "Form1.frx":73BF
      ToolTipText     =   "Finish"
      Top             =   2760
      Width           =   465
   End
   Begin VB.Image beg_but 
      Height          =   420
      Left            =   2950
      Picture         =   "Form1.frx":7842
      ToolTipText     =   "Start"
      Top             =   2760
      Width           =   465
   End
   Begin VB.Label song_left 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "0.00"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5280
      TabIndex        =   12
      Top             =   720
      Width           =   615
   End
   Begin VB.Label song_dur 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "0.00"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   720
      Width           =   615
   End
   Begin VB.Label balance_pos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Center"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      ToolTipText     =   "Balance"
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1320
      TabIndex        =   9
      ToolTipText     =   "Balance"
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Volume:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      ToolTipText     =   "Balance"
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label vol 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "75"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5280
      TabIndex        =   5
      ToolTipText     =   "Volume"
      Top             =   3600
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   5640
      TabIndex        =   4
      ToolTipText     =   "Volume"
      Top             =   3600
      Width           =   255
   End
   Begin VB.Label songdir 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "EMPTY"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1560
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Image stop_but 
      Height          =   420
      Left            =   4800
      Picture         =   "Form1.frx":7CC3
      ToolTipText     =   "Stop"
      Top             =   2760
      Width           =   600
   End
   Begin VB.Image rskip_but 
      Height          =   420
      Left            =   3860
      Picture         =   "Form1.frx":809E
      ToolTipText     =   "Skip"
      Top             =   2760
      Width           =   465
   End
   Begin VB.Image lskip_but 
      Height          =   420
      Left            =   3400
      Picture         =   "Form1.frx":8478
      ToolTipText     =   "Skip"
      Top             =   2760
      Width           =   465
   End
   Begin VB.Image pause_but 
      Height          =   420
      Left            =   2380
      Picture         =   "Form1.frx":8858
      ToolTipText     =   "Pause"
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image play_but 
      Height          =   420
      Left            =   1800
      Picture         =   "Form1.frx":8D26
      ToolTipText     =   "Play"
      Top             =   2760
      Width           =   585
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   705
      Left            =   1560
      TabIndex        =   2
      Top             =   6600
      Visible         =   0   'False
      Width           =   3975
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   0
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   -1  'True
      VideoBorderWidth=   1
      VideoBorderColor=   12632256
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label filetitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      Caption         =   "EMPTY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Top             =   360
      Width           =   4575
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sel

Sub songdur()

a = Int(MediaPlayer1.Duration) / 60
Let song_dur.Caption = Left(a, 4)
Timer1.Enabled = True

End Sub

Private Sub add_but_Click()

On Error Resume Next 'continue on error

cd1.DialogTitle = "Select MP3 to add to PlayList" 'Title
cd1.InitDir = App.Path 'start at app path
cd1.Filter = "Music(*.mp3;*.wav)|*.mp3;*.wav" 'only allow MP3 files
cd1.CancelError = True 'cancel error is true
cd1.ShowOpen 'show open screen

If cd1.filename = "" Then Exit Sub 'if cancel selected then add nothing to playlist
playlist.AddItem cd1.filename 'add selected MP3 file
songlist.AddItem cd1.filetitle 'add selected MP3 file
cd1.filename = "" 'reset common dialog filename

End Sub

Private Sub balance_Change()

On Error Resume Next 'exit on error
Let MediaPlayer1.balance = balance.Value 'set balance level

'show balance position
If balance.Value > -500 And balance.Value < 500 Then balance_pos.Caption = "Center"
If balance.Value < -500 Then balance_pos.Caption = "Left"
If balance.Value > 500 Then balance_pos.Caption = "Right"


End Sub

Private Sub beg_but_Click()
On Error Resume Next
Let MediaPlayer1.CurrentPosition = 0

End Sub

Private Sub clear_but_Click()

'ask if user wants to clear list
Const MB_OK = 0, MB_OKCANCEL = 1
    Const MB_YESNOCANCEL = 3, MB_YESNO = 4
    Const MB_ICONSTOP = 16, MB_ICONQUESTION = 32
    Const MB_ICONEXCLAMATION = 48, MB_ICONINFORMATION = 64
    Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7
    Dim DgDef, Msg, Response, Title
    Msg = "ARE YOU SURE YOU WANT TO CLEAR PLAYLIST?"
    
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, Title)
    'if yes clear list
    If Response = IDYES Then
        playlist.Clear
        songlist.Clear
        MediaPlayer1.filename = ""
        Exit Sub
    Else
    'if no then exit the sub
        Exit Sub
    End If



End Sub

















Private Sub Command1_Click()

Let main.WindowState = 1 - minimized


End Sub

Private Sub exit_but_Click()
'ask user if they wana exit program
Const MB_OK = 0, MB_OKCANCEL = 1
    Const MB_YESNOCANCEL = 3, MB_YESNO = 4
    Const MB_ICONSTOP = 16, MB_ICONQUESTION = 32
    Const MB_ICONEXCLAMATION = 48, MB_ICONINFORMATION = 64
    Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7
    Dim DgDef, Msg, Response, Title
    Msg = "ARE YOU SURE YOU WANT TO LEAVE?"
    
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2

Response = MsgBox(Msg, DgDef, Title)
    If Response = IDYES Then
        'if yes exit and stop playing any songs
        MediaPlayer1.Stop
        MediaPlayer1.filename = ""
        End
        Exit Sub
    Else
        'if no return to program
        Exit Sub
    End If
    
End Sub

Private Sub end_but_Click()
On Error Resume Next
Let MediaPlayer1.CurrentPosition = MediaPlayer1.Duration

If filetitle.Caption = "EMPTY" Then
    Timer1.Enabled = False
    DoEvents
    song_left.Caption = "0.00"
End If
End Sub

Private Sub Form_Load()

'set sel to -1 (no song selected from list)
Let sel = -1
Let pause_but.Picture = p1.Picture

End Sub

Private Sub HScroll1_Change()

On Error Resume Next

yy = HScroll1.Value - 2500
MediaPlayer1.Volume = yy
Dim a As Integer, b As Integer
b = HScroll1.Min
a = HScroll1.Value
vol.Caption = a \ 25


End Sub

Private Sub HScroll1_Scroll()

On Error Resume Next
Dim a As Integer, b As Integer

yy = HScroll1.Value - 2500
MediaPlayer1.Volume = yy
b = HScroll1.Min
a = HScroll1.Value
vol.Caption = a \ 25


End Sub


Private Sub list_but_Click()

If list_but.Caption = "Hide" Then
    Let list_frame.Visible = False
    Let list_but.Caption = "Show"
    Exit Sub
Else
    Let list_frame.Visible = True
    Let list_but.Caption = "Hide"
    Exit Sub
End If

End Sub

Private Sub load_but_Click()
On Error Resume Next
Dim filedata

cd1.DialogTitle = "Open PlayList" 'Title
cd1.InitDir = App.Path 'start at app path
cd1.Filter = "Playlist Files|*.mpl" 'only allow MP3 files
cd1.CancelError = True 'cancel error is true
cd1.ShowOpen 'show open screen

If cd1.filename = "" Then Exit Sub 'if cancel selected then add nothing to playlist

Close #1
Open cd1.filename For Input As #1
    Do While EOF(1) = False
        DoEvents
        Input #1, filedata  'input song path
        Input #1, filedata2 'input song title
        playlist.AddItem filedata   'add song path to list
        songlist.AddItem filedata2  'add song title to list
    Loop
Close #1

cd1.filename = ""

End Sub

Private Sub lskip_but_Click()
On Error Resume Next
MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition - 10

End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)


Let MediaPlayer1.CurrentPosition = 0 'reset mp position
If songlist.ListIndex = songlist.ListCount - 1 Then
    'if loop no true then stop playing from list
    If loop_but.Value = 0 - Unchecked Then
        Let songlist.ListIndex = 0
        Let playlist.ListIndex = 0
        Let filetitle.Caption = "EMPTY"
        Let song_dur.Caption = "0.00"
        Exit Sub
    Else
    'if loop checked then go back to begining
        Let playlist.ListIndex = 0
        Let filetitle.Caption = songlist.List(playlist.ListIndex)
        MediaPlayer1.filename = playlist.List(0)
        songlist.ListIndex = playlist.ListIndex
        Call songdur
        Exit Sub
    End If
End If

'play next song in list
Dim num
Let num = playlist.ListIndex + 1
Let playlist.ListIndex = num
Let songlist.ListIndex = num
Let filetitle.Caption = songlist.List(num)
Let MediaPlayer1.filename = playlist.List(num)
Call songdur






End Sub





Private Sub pause_but_Click()

On Error Resume Next
If pause_but.Picture = p1.Picture Then
    MediaPlayer1.Pause
    Let pause_but.Picture = p2.Picture
    Exit Sub
Else
    MediaPlayer1.Play
    Let pause_but.Picture = p1.Picture
    Exit Sub
End If


End Sub

Private Sub play_but_Click()
On Error Resume Next


    sel = songlist.ListIndex
    MediaPlayer1.filename = playlist.List(sel) 'load file path from list of paths
If MediaPlayer1.filename = "" Then
    MsgBox "Please Load a song from the PlayList"
    Exit Sub
End If
    Call songdur
    MediaPlayer1.Play   'play song
    Let filetitle.Caption = songlist.List(sel) 'show song title
    Let playlist.ListIndex = sel


End Sub
Private Sub rem_but_Click()

'inform user no song selected
If songlist.ListIndex <= -1 Then
    MsgBox "PLEASE SELECT A SONG TO REMOVE"
    Exit Sub
End If

'remove selected song
playlist.RemoveItem (sel)
songlist.RemoveItem (sel)
sel = -1


End Sub


Private Sub rskip_but_Click()
On Error Resume Next
MediaPlayer1.CurrentPosition = MediaPlayer1.CurrentPosition + 10

End Sub

Private Sub save_but_Click()

On Error Resume Next
cd1.DialogTitle = "Save PlayList" 'Title
cd1.InitDir = App.Path 'start at app path
cd1.Filter = "Playlist Files|*.mpl" 'only allow MP3 files
cd1.CancelError = True 'cancel error is true
cd1.ShowSave 'show open screen

If cd1.filename = "" Then Exit Sub 'if cancel selected then add nothing to playlist

Close #1
Open cd1.filename For Output As #1
    For X = 0 To songlist.ListCount - 1
        Print #1, playlist.List(X)  'write song path to file
        Print #1, songlist.List(X)  'write song title to file
    Next X
Close #1

cd1.filename = ""

End Sub







Private Sub songlist_Click()


sel = songlist.ListIndex 'set value (sel) to selected item in list


End Sub


Private Sub songlist_DblClick()

sel = songlist.ListIndex
MediaPlayer1.filename = playlist.List(sel) 'load file path from list of paths
Call songdur
MediaPlayer1.Play   'play song
Let filetitle.Caption = songlist.List(sel) 'show song title
Let playlist.ListIndex = sel



End Sub









Private Sub stop_but_Click()
On Error Resume Next
MediaPlayer1.Stop
MediaPlayer1.CurrentPosition = 0
MediaPlayer1.filename = ""
filetitle.Caption = "EMPTY"
song_dur.Caption = "0.00"
song_left.Caption = "0.00"


End Sub

Private Sub Timer1_Timer()

a = MediaPlayer1.CurrentPosition
b = MediaPlayer1.Duration - a
c = b / 60

Let song_left.Caption = Left(c, 4)


End Sub


