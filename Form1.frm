VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BartNet Music Player"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "Form1.frx":0442
   ScaleHeight     =   6555
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin Project1.chameleonButton cmdForward 
      Height          =   375
      Left            =   6720
      TabIndex        =   23
      Top             =   6000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ">>"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":45DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdBack 
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   6000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "<<"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":45F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdDown 
      Height          =   285
      Left            =   2760
      TabIndex        =   18
      Top             =   4680
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Down"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4615
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdUp 
      Height          =   285
      Left            =   2760
      TabIndex        =   17
      Top             =   4320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "Up"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4631
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar pVolume 
      Height          =   1455
      Left            =   2520
      TabIndex        =   16
      Top             =   3480
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   2566
      _Version        =   393216
      Appearance      =   0
      Max             =   2500
      Orientation     =   1
      Scrolling       =   1
   End
   Begin Project1.chameleonButton cmdShuffle 
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4560
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Shuffle"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":464D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdRepeat 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Repeat"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4669
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdNormal 
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   3600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Normal"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4685
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdPause 
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   6000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Pause"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":46A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdStop 
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   6000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Stop"
      ENAB            =   0   'False
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":46BD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdPlay 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   6000
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Play"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":46D9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdAddFolder 
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2400
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Add All Files From Folder To PlayList"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":46F5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdAddFile 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   1920
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Add Single File To PlayList"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4711
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdDeletePlayList 
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   1320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Delete Existing PlayList"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":472D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdCreatePlayList 
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Create New PlayList"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Form1.frx":4749
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Timer timCheckCD 
      Interval        =   1000
      Left            =   43200
      Top             =   4920
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   840
      Top             =   58800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4765
            Key             =   "Icon"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   58800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4AB7
            Key             =   "CD"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4DD9
            Key             =   "PlayList"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   40800
   End
   Begin MSComDlg.CommonDialog c1 
      Left            =   960
      Top             =   50000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView l1 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4683
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ColHdrIcons     =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageCombo cboLists 
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      ImageList       =   "ImageList1"
   End
   Begin MSComctlLib.ProgressBar pTrack 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5640
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin VB.Line Border3 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   3840
      X2              =   7320
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Border1 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   7320
      X2              =   4680
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Border5 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   3960
      X2              =   3830
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label lblFileInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "File Info"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   4060
      TabIndex        =   28
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Line Border2 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   7320
      X2              =   7320
      Y1              =   5040
      Y2              =   3360
   End
   Begin VB.Line Border4 
      BorderColor     =   &H80000005&
      Visible         =   0   'False
      X1              =   3840
      X2              =   3840
      Y1              =   5040
      Y2              =   3360
   End
   Begin VB.Label lblDuration 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3960
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3960
      TabIndex        =   25
      Top             =   3480
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MediaPlayerCtl.MediaPlayer mCheck 
      Height          =   735
      Left            =   4680
      TabIndex        =   24
      Top             =   36000
      Visible         =   0   'False
      Width           =   855
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
      DisplayForeColor=   16777215
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
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblPosition 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   6000
      TabIndex        =   21
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label lblPlaying 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5400
      Width           =   5775
   End
   Begin VB.Line Line10 
      BorderColor     =   &H80000005&
      X1              =   3480
      X2              =   3240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line9 
      BorderColor     =   &H80000005&
      X1              =   2640
      X2              =   2400
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line8 
      BorderColor     =   &H80000005&
      X1              =   3480
      X2              =   3480
      Y1              =   5040
      Y2              =   3360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H80000005&
      X1              =   2400
      X2              =   2400
      Y1              =   3360
      Y2              =   5040
   End
   Begin VB.Line Line6 
      BorderColor     =   &H80000005&
      X1              =   2400
      X2              =   3480
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblVolume 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2760
      TabIndex        =   19
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   3240
      Width           =   615
   End
   Begin VB.Line Line5 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   240
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   1320
      X2              =   1440
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   120
      Y1              =   3360
      Y2              =   5040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   1440
      X2              =   1440
      Y1              =   3360
      Y2              =   5040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   0
      X1              =   120
      X2              =   1450
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label lblPlayingMode 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Playing Mode"
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3240
      Width           =   1095
   End
   Begin MediaPlayerCtl.MediaPlayer m1 
      Height          =   2775
      Left            =   1200
      TabIndex        =   2
      Top             =   30000
      Width           =   3855
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
      DisplayForeColor=   16777215
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
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000005&
      Height          =   975
      Left            =   3960
      TabIndex        =   27
      Top             =   3960
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "a"
      Visible         =   0   'False
      Begin VB.Menu mnuPlay 
         Caption         =   "Play"
      End
      Begin VB.Menu jk 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove From List"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private fso As New FileSystemObject
Private Drives As Drives
Private Files As Files
Private Drive As Drive
Private File As File
Private strm As TextStream
Private Folder As Folder

Private Sub GetInfo()
On Error GoTo a
    Dim TimePast As Integer
    Dim Minutes As Integer
    Dim Seconds As Integer

    lblFileInfo.Visible = True
    Border1.Visible = True
    Border2.Visible = True
    Border3.Visible = True
    Border4.Visible = True
    Border5.Visible = True
    lblName.Visible = True
    lblDuration.Visible = True
    lblPath.Visible = True
    lblName.Caption = l1.SelectedItem.Text
    mCheck.Mute = True
    mCheck.FileName = l1.SelectedItem.Key
    TimePast = mCheck.Duration
    Minutes = TimePast / 60
    Seconds = TimePast - (Minutes * 60)
    If Seconds < 0 Then Seconds = 0
    If Seconds < 10 Then
        lblDuration.Caption = "Duration : " & Minutes & ":0" & Seconds
    Else
        lblDuration.Caption = "Duration : " & Minutes & ":" & Seconds
    End If
    lblPath.Caption = l1.SelectedItem.Key
    
    Exit Sub
    
a:
    lblFileInfo.Visible = False
    Border1.Visible = False
    Border2.Visible = False
    Border3.Visible = False
    Border4.Visible = False
    Border5.Visible = False
    lblName.Visible = False
    lblDuration.Visible = False
    lblPath.Visible = False
End Sub

Private Sub cboLists_Click()
On Error Resume Next

    l1.ListItems.Clear
    
    If Mid(cboLists.SelectedItem.Key, 1, 4) = "List" Then
        Set strm = fso.OpenTextFile(App.Path & "\Settings\" & Mid(cboLists.SelectedItem.Key, 5, Len(cboLists.SelectedItem.Key) - 4) & ".BartNet")
        Dim a As String
        With strm
            a = .ReadLine
            If a = "none" Then
                Exit Sub
            Else
                l1.ListItems.Add , a, .ReadLine, "Icon", "Icon"
            End If
            Do Until .AtEndOfStream
                l1.ListItems.Add , .ReadLine, .ReadLine, "Icon", "Icon"
            Loop
            .Close
        End With
    Else
        Set Folder = fso.GetFolder(cboLists.SelectedItem.Key & ":\")
        Set Files = Folder.Files
        
        For Each File In Files
            Select Case File.Type
                Case "MP3 Format Sound"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
                Case "Windows Media Audio File"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
                Case "Wave Sound"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
                Case "CD Audio Track"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
            End Select
        Next
    End If
End Sub

Private Sub cboLists_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    l1.ListItems.Clear
    
    If Mid(cboLists.SelectedItem.Key, 1, 4) = "List" Then
        Set strm = fso.OpenTextFile(App.Path & "\Settings\" & Mid(cboLists.SelectedItem.Key, 5, Len(cboLists.SelectedItem.Key) - 4) & ".BartNet")
        Dim a As String
        With strm
            a = .ReadLine
            If a = "none" Then
                Exit Sub
            Else
                l1.ListItems.Add , a, .ReadLine, "Icon", "Icon"
            End If
            Do Until .AtEndOfStream
                l1.ListItems.Add , .ReadLine, .ReadLine, "Icon", "Icon"
            Loop
            .Close
        End With
    Else
        Set Folder = fso.GetFolder(cboLists.SelectedItem.Key & ":\")
        Set Files = Folder.Files
        
        For Each File In Files
            Select Case File.Type
                Case "MP3 Format Sound"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
                Case "Windows Media Audio File"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
                Case "Wave Sound"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
                Case "CD Audio Track"
                    l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
            End Select
        Next
    End If
End Sub


Private Sub cmdAddFile_Click()
On Error GoTo a
    
    If Mid(cboLists.SelectedItem.Key, 1, 4) = "List" Then
        c1.Filter = "All Music Files|*.mp3;*.wav;*.wma|All Files|*.*"
        c1.DialogTitle = "Select File"
        c1.ShowOpen
        
        Set strm = fso.OpenTextFile(App.Path & "\Settings\" & Mid(cboLists.SelectedItem.Key, 5, Len(cboLists.SelectedItem.Key) - 4) & ".BartNet", ForAppending)
        strm.WriteLine c1.FileName
        strm.WriteLine c1.FileTitle
        strm.Close
        l1.ListItems.Add , c1.FileName, c1.FileTitle, "Icon", "Icon"
    Else
    
    End If
    
a:

End Sub

Private Sub cmdAddFolder_Click()
On Error GoTo a

    If Mid(cboLists.SelectedItem.Key, 1, 4) = "List" Then
        Form2.Show vbModal, Me
    Else
    
    End If
    
a:

End Sub

Private Sub cmdBack_Click()
    Dim Item As ListItem
    Dim Items As ListItems
    
    m1.Stop
    lblPosition.Caption = ""
    lblPlaying.Caption = ""
    pTrack.Value = 0
    cmdPlay.Enabled = True
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    Timer1.Enabled = False
    
    Set Items = l1.ListItems
    For Each Item In Items
        If Item.Key = lblPlaying.Tag Then
            If Item.Index = 1 Then
                m1.FileName = Items.Item(Items.Count).Key
                lblPlaying.Caption = Items.Item(Items.Count).Text
                lblPlaying.Tag = Items.Item(Items.Count).Key
                m1.Play
                pTrack.Min = 0
                pTrack.Max = m1.Duration
                pTrack.Value = m1.CurrentPosition
                cmdPlay.Enabled = False
                cmdStop.Enabled = True
                cmdPause.Enabled = True
                Timer1.Enabled = True
                
                Exit Sub
            Else
                m1.FileName = Items.Item(Item.Index - 1).Key
                lblPlaying.Caption = Items.Item(Item.Index - 1).Text
                lblPlaying.Tag = Items.Item(Item.Index - 1).Key
                m1.Play
                pTrack.Min = 0
                pTrack.Max = m1.Duration
                pTrack.Value = m1.CurrentPosition
                cmdPlay.Enabled = False
                cmdStop.Enabled = True
                cmdPause.Enabled = True
                Timer1.Enabled = True
                
                Exit Sub
            End If
        Else
        
        End If
    Next
End Sub

Private Sub cmdCreatePlayList_Click()
    Dim a As String
    
    a = InputBox("Enter A Name Fore The New Playlist", "New Playlist", "New Playlist", Me.Left + 500, Me.Top + 500)
    If a = "" Then
    
    Else
        Set strm = fso.OpenTextFile(App.Path & "\Settings\PlayLists.BartNet", ForReading)
        If strm.ReadLine = "none" Then
            strm.Close
            Set strm = fso.OpenTextFile(App.Path & "\Settings\PlayLists.BartNet", ForWriting)
            strm.WriteLine a
            strm.Close
        Else
            strm.Close
            Set strm = fso.OpenTextFile(App.Path & "\Settings\PlayLists.BartNet", ForAppending)
            strm.WriteLine a
            strm.Close
        End If
        
        Set strm = fso.CreateTextFile(App.Path & "\Settings\" & a & ".BartNet")
        strm.Close
        
        cboLists.ComboItems.Add , "List" & a, a, "PlayList", "PlayList"
    End If
End Sub

Private Sub cmdDeletePlayList_Click()
    If Mid(cboLists.SelectedItem.Key, 1, 4) = "List" Then
        Dim a As ComboItem
        Dim B As ComboItems
        Dim c As Boolean
        
        c = False
        Kill App.Path & "\Settings\" & Mid(cboLists.SelectedItem.Key, 5, Len(cboLists.SelectedItem.Key) - 4) & ".BartNet"
        cboLists.ComboItems.Remove (cboLists.SelectedItem.Index)
        Set B = cboLists.ComboItems
        Set strm = fso.OpenTextFile(App.Path & "\Settings\PlayLists.BartNet", ForWriting)
        For Each a In B
            If Mid(a.Key, 1, 4) = "List" Then
                strm.WriteLine Mid(a.Key, 5, Len(a.Key) - 4)
                c = True
            Else
            
            End If
        Next
        
        If c = True Then
            strm.Close
        Else
            strm.WriteLine "none"
            strm.Close
        End If
    Else
        MsgBox "You Can't Delete A CD-ROM Drive From The List", vbOKOnly + vbInformation, "Alert"
    End If
    
    cboLists.Text = ""
    l1.ListItems.Clear
End Sub

Private Sub cmdDown_Click()
    cmdUp.Enabled = True
    pVolume.Value = pVolume.Value - 100
    
    If pVolume.Value < 100 Then
        cmdDown.Enabled = False
        m1.Mute = True
    Else
    
    End If
    
    m1.Volume = pVolume.Value - 2500
    lblVolume.Caption = pVolume.Value / 25 & "%"
End Sub

Private Sub cmdForward_Click()
    Dim Item As ListItem
    Dim Items As ListItems

    If cmdShuffle.Enabled = False Then
        m1.Stop
        lblPosition.Caption = ""
        lblPlaying.Caption = ""
        pTrack.Value = 0
        cmdPlay.Enabled = True
        cmdStop.Enabled = False
        cmdPause.Enabled = False
        Timer1.Enabled = False
        
        Dim a As Integer
        
        Set Items = l1.ListItems
        
        Select Case Items.Count
            Case 1 To 10
                a = Round(Rnd * 10, 0)
                
                If a = 0 Then a = 1
                
                If a > Items.Count Then
                    cmdForward_Click
                    Exit Sub
                Else
                
                End If
                
                m1.FileName = Items.Item(a).Key
                lblPlaying.Caption = Items.Item(a).Text
                lblPlaying.Tag = Items.Item(a).Key
                m1.Play
                pTrack.Min = 0
                pTrack.Max = m1.Duration
                pTrack.Value = m1.CurrentPosition
                cmdPlay.Enabled = False
                cmdStop.Enabled = True
                cmdPause.Enabled = True
                Timer1.Enabled = True
                
                Exit Sub
            Case 11 To 100
                a = Round(Rnd * 100, 0)
                
                If a > Items.Count Then
                    cmdForward_Click
                    Exit Sub
                Else
                
                End If
                
                m1.FileName = Items.Item(a).Key
                lblPlaying.Caption = Items.Item(a).Text
                lblPlaying.Tag = Items.Item(a).Key
                m1.Play
                pTrack.Min = 0
                pTrack.Max = m1.Duration
                pTrack.Value = m1.CurrentPosition
                cmdPlay.Enabled = False
                cmdStop.Enabled = True
                cmdPause.Enabled = True
                Timer1.Enabled = True
                
                Exit Sub
            Case 101 To 1000
                a = Round(Rnd * 1000, 0)
                
                If a > Items.Count Then
                    cmdForward_Click
                    Exit Sub
                Else
                
                End If
                
                m1.FileName = Items.Item(a).Key
                lblPlaying.Caption = Items.Item(a).Text
                lblPlaying.Tag = Items.Item(a).Key
                m1.Play
                pTrack.Min = 0
                pTrack.Max = m1.Duration
                pTrack.Value = m1.CurrentPosition
                cmdPlay.Enabled = False
                cmdStop.Enabled = True
                cmdPause.Enabled = True
                Timer1.Enabled = True
                
                Exit Sub
        End Select
    Else
        m1.Stop
        lblPosition.Caption = ""
        lblPlaying.Caption = ""
        pTrack.Value = 0
        cmdPlay.Enabled = True
        cmdStop.Enabled = False
        cmdPause.Enabled = False
        Timer1.Enabled = False
        
        Set Items = l1.ListItems
        For Each Item In Items
            If Item.Key = lblPlaying.Tag Then
                If Item.Index = Items.Count Then
                    m1.FileName = Items.Item(1).Key
                    lblPlaying.Caption = Items.Item(1).Text
                    lblPlaying.Tag = Items.Item(1).Key
                    m1.Play
                    pTrack.Min = 0
                    pTrack.Max = m1.Duration
                    pTrack.Value = m1.CurrentPosition
                    cmdPlay.Enabled = False
                    cmdStop.Enabled = True
                    cmdPause.Enabled = True
                    Timer1.Enabled = True
                    
                    Exit Sub
                Else
                    m1.FileName = Items.Item(Item.Index + 1).Key
                    lblPlaying.Caption = Items.Item(Item.Index + 1).Text
                    lblPlaying.Tag = Items.Item(Item.Index + 1).Key
                    m1.Play
                    pTrack.Min = 0
                    pTrack.Max = m1.Duration
                    pTrack.Value = m1.CurrentPosition
                    cmdPlay.Enabled = False
                    cmdStop.Enabled = True
                    cmdPause.Enabled = True
                    Timer1.Enabled = True
                    
                    Exit Sub
                End If
            Else
            
            End If
        Next
    End If
End Sub

Private Sub cmdNormal_Click()
    cmdRepeat.Enabled = True
    cmdShuffle.Enabled = True
    cmdNormal.Enabled = False
    cmdForward.Visible = True
    cmdBack.Visible = True
End Sub

Private Sub cmdPause_Click()
    m1.Pause
    
    cmdPause.Enabled = False
    cmdPlay.Enabled = True
End Sub

Private Sub cmdPlay_Click()
On Error GoTo a

    If cmdStop.Enabled = True Then
        m1.Play
        
        cmdPlay.Enabled = False
        cmdPause.Enabled = True
    Else
        m1.FileName = l1.SelectedItem.Key
        lblPlaying.Caption = l1.SelectedItem.Text
        lblPlaying.Tag = l1.SelectedItem.Key
        m1.Play
        pTrack.Min = 0
        pTrack.Max = m1.Duration
        pTrack.Value = m1.CurrentPosition
        cmdPlay.Enabled = False
        cmdStop.Enabled = True
        cmdPause.Enabled = True
        Timer1.Enabled = True
    End If
      
    Exit Sub
    
a:
    MsgBox "An error has ocurred, make sure the file hasn't been moved.", vbOKOnly + vbCritical, "BartNet Music Player"
End Sub

Private Sub cmdRepeat_Click()
    cmdRepeat.Enabled = False
    cmdShuffle.Enabled = True
    cmdNormal.Enabled = True
    cmdForward.Visible = True
    cmdBack.Visible = True
End Sub

Private Sub cmdShuffle_Click()
    cmdRepeat.Enabled = True
    cmdShuffle.Enabled = False
    cmdNormal.Enabled = True
    cmdForward.Visible = True
    cmdBack.Visible = False
End Sub

Private Sub cmdStop_Click()
On Error Resume Next

    m1.Stop

    lblPosition.Caption = ""
    lblPlaying.Caption = ""
    lblPlaying.Tag = ""
    pTrack.Value = 0
    
    cmdPlay.Enabled = True
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    
    Timer1.Enabled = False
End Sub

Private Sub cmdUp_Click()
    cmdDown.Enabled = True
    m1.Mute = False
    pVolume.Value = pVolume.Value + 100
    
    If pVolume.Value > pVolume.Max - 100 Then
        cmdUp.Enabled = False
    Else
    
    End If
    
    m1.Volume = pVolume.Value - 2500
    lblVolume.Caption = pVolume.Value / 25 & "%"
End Sub
Private Sub Form_Click()
On Error Resume Next

    l1_LostFocus
    l1.SelectedItem.Selected = False
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    pVolume.Value = pVolume.Max
    lblName.Caption = ""
    lblDuration.Caption = ""
    lblPath.Caption = ""

    Set Drives = fso.Drives
    For Each Drive In Drives
        If Drive.DriveType = CDRom Then
            If Drive.IsReady = True Then
                cboLists.ComboItems.Add , Drive.DriveLetter, Drive.VolumeName, "CD", "CD"
            Else
                cboLists.ComboItems.Add , Drive.DriveLetter, Drive.DriveLetter & " - Please Insert Disk", "CD", "CD"
            End If
        Else
        
        End If
    Next
    
    Set strm = fso.OpenTextFile(App.Path & "\Settings\PlayLists.BartNet", ForReading)
    
    Dim a As String
    With strm
        a = .ReadLine
        If a = "none" Then
            Exit Sub
        Else
            cboLists.ComboItems.Add , "List" & a, a, "PlayList", "PlayList"
        End If
        
        Do Until .AtEndOfStream
            a = .ReadLine
            cboLists.ComboItems.Add , "List" & a, a, "PlayList", "PlayList"
        Loop
        .Close
    End With
    
    cmdNormal.Enabled = False
    cmdUp.Enabled = False
    pVolume.Value = 2500
    pTrack.Value = 0
    lblPlaying.Caption = ""
    lblPosition.Caption = ""
    cmdStop.Enabled = False
    cmdPause.Enabled = False
    
    m1.Volume = 0
    lblVolume.Caption = "100%"
End Sub

Private Sub l1_KeyUp(KeyCode As Integer, Shift As Integer)
    GetInfo
End Sub

Private Sub l1_LostFocus()
    lblFileInfo.Visible = False
    Border1.Visible = False
    Border2.Visible = False
    Border3.Visible = False
    Border4.Visible = False
    Border5.Visible = False
    lblName.Visible = False
    lblDuration.Visible = False
    lblPath.Visible = False
End Sub


Private Sub l1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuPopup
    Else
        GetInfo
    End If
End Sub


Private Sub m1_EndOfStream(ByVal Result As Long)
    Dim Item As ListItem
    Dim Items As ListItems

    If cmdNormal.Enabled = False Then
        m1.Stop
        lblPosition.Caption = ""
        lblPlaying.Caption = ""
        pTrack.Value = 0
        cmdPlay.Enabled = True
        cmdStop.Enabled = False
        cmdPause.Enabled = False
        Timer1.Enabled = False
        
        Set Items = l1.ListItems
        For Each Item In Items
            If Item.Key = lblPlaying.Tag Then
                If Item.Index = Items.Count Then
                
                Else
                    m1.FileName = Items.Item(Item.Index + 1).Key
                    lblPlaying.Caption = Items.Item(Item.Index + 1).Text
                    lblPlaying.Tag = Items.Item(Item.Index + 1).Key
                    m1.Play
                    pTrack.Min = 0
                    pTrack.Max = m1.Duration
                    pTrack.Value = m1.CurrentPosition
                    cmdPlay.Enabled = False
                    cmdStop.Enabled = True
                    cmdPause.Enabled = True
                    Timer1.Enabled = True
                    
                    Exit Sub
                End If
            Else
            
            End If
        Next
    Else
        If cmdRepeat.Enabled = False Then
            m1.Stop
            lblPosition.Caption = ""
            lblPlaying.Caption = ""
            pTrack.Value = 0
            cmdPlay.Enabled = True
            cmdStop.Enabled = False
            cmdPause.Enabled = False
            Timer1.Enabled = False
            
            Set Items = l1.ListItems
            For Each Item In Items
                If Item.Key = lblPlaying.Tag Then
                    If Item.Index = Items.Count Then
                        m1.FileName = Items.Item(1).Key
                        lblPlaying.Caption = Items.Item(1).Text
                        lblPlaying.Tag = Items.Item(1).Key
                        m1.Play
                        pTrack.Min = 0
                        pTrack.Max = m1.Duration
                        pTrack.Value = m1.CurrentPosition
                        cmdPlay.Enabled = False
                        cmdStop.Enabled = True
                        cmdPause.Enabled = True
                        Timer1.Enabled = True
                        
                        Exit Sub
                    Else
                        m1.FileName = Items.Item(Item.Index + 1).Key
                        lblPlaying.Caption = Items.Item(Item.Index + 1).Text
                        lblPlaying.Tag = Items.Item(Item.Index + 1).Key
                        m1.Play
                        pTrack.Min = 0
                        pTrack.Max = m1.Duration
                        pTrack.Value = m1.CurrentPosition
                        cmdPlay.Enabled = False
                        cmdStop.Enabled = True
                        cmdPause.Enabled = True
                        Timer1.Enabled = True
                        
                        Exit Sub
                    End If
                Else
                
                End If
            Next
        Else
            m1.Stop
            lblPosition.Caption = ""
            lblPlaying.Caption = ""
            pTrack.Value = 0
            cmdPlay.Enabled = True
            cmdStop.Enabled = False
            cmdPause.Enabled = False
            Timer1.Enabled = False
            
            Dim a As Integer
            
            Set Items = l1.ListItems
            
            Select Case Items.Count
                Case 1 To 10
                    a = Round(Rnd * 10, 0)
                    
                    If a = 0 Then a = 1
                    
                    If a > Items.Count Then
                        m1_EndOfStream (25)
                        Exit Sub
                    Else
                    
                    End If
                    
                    m1.FileName = Items.Item(a).Key
                    lblPlaying.Caption = Items.Item(a).Text
                    lblPlaying.Tag = Items.Item(a).Key
                    m1.Play
                    pTrack.Min = 0
                    pTrack.Max = m1.Duration
                    pTrack.Value = m1.CurrentPosition
                    cmdPlay.Enabled = False
                    cmdStop.Enabled = True
                    cmdPause.Enabled = True
                    Timer1.Enabled = True
                    
                    Exit Sub
                Case 11 To 100
                    a = Round(Rnd * 100, 0)
                    
                    If a > Items.Count Then
                        m1_EndOfStream (25)
                        Exit Sub
                    Else
                    
                    End If
                    
                    m1.FileName = Items.Item(a).Key
                    lblPlaying.Caption = Items.Item(a).Text
                    lblPlaying.Tag = Items.Item(a).Key
                    m1.Play
                    pTrack.Min = 0
                    pTrack.Max = m1.Duration
                    pTrack.Value = m1.CurrentPosition
                    cmdPlay.Enabled = False
                    cmdStop.Enabled = True
                    cmdPause.Enabled = True
                    Timer1.Enabled = True
                    
                    Exit Sub
                Case 101 To 1000
                    a = Round(Rnd * 1000, 0)
                    
                    If a > Items.Count Then
                        m1_EndOfStream (25)
                        Exit Sub
                    Else
                    
                    End If
                    
                    m1.FileName = Items.Item(a).Key
                    lblPlaying.Caption = Items.Item(a).Text
                    lblPlaying.Tag = Items.Item(a).Key
                    m1.Play
                    pTrack.Min = 0
                    pTrack.Max = m1.Duration
                    pTrack.Value = m1.CurrentPosition
                    cmdPlay.Enabled = False
                    cmdStop.Enabled = True
                    cmdPause.Enabled = True
                    Timer1.Enabled = True
                    
                    Exit Sub
            End Select
        End If
    End If
End Sub


Private Sub mnuPlay_Click()
    cmdStop_Click
    cmdPlay_Click
End Sub

Private Sub mnuRemove_Click()
    Dim a As Integer
    Dim List As ListItem
    Dim Lists As ListItems
    
    a = MsgBox("Are you sure you want to remove this file from this playlist?", vbYesNo + vbInformation, "BartNet Music Player")
    
    If a = 6 Then
        cmdStop_Click
    
        l1.ListItems.Remove (l1.SelectedItem.Index)
        
        Set Lists = l1.ListItems
        Set strm = fso.OpenTextFile(App.Path & "\Settings\" & cboLists.Text & ".BartNet", ForWriting)
        
        With strm
            For Each List In Lists
                .WriteLine List.Key
                .WriteLine List.Text
            Next
            .Close
        End With
    Else
    
    End If
End Sub


Private Sub timCheckCD_Timer()
    Dim ComboItem As ComboItem
    Dim ComboItems As ComboItems
    
    Set ComboItems = cboLists.ComboItems
    Set Drives = fso.Drives
    
    For Each ComboItem In ComboItems
        For Each Drive In Drives
            If ComboItem.Key = Drive.DriveLetter Then
                If Drive.IsReady = True Then
                    If ComboItem.Text = Drive.VolumeName Then
                    
                    Else
                        ComboItem.Text = Drive.VolumeName
                    End If
                Else
                    If ComboItem.Text = Drive.DriveLetter & " - Please Insert Disk" Then
                    
                    Else
                        ComboItem.Text = Drive.DriveLetter & " - Please Insert Disk"
                    End If
                End If
            Else
            
            End If
        Next
    Next
End Sub

Private Sub Timer1_Timer()
    Dim TimePast As Integer
    Dim Minutes As Integer
    Dim Seconds As Integer
    
    If m1.CurrentPosition >= 0 Then
        pTrack.Value = m1.CurrentPosition
    Else
        pTrack.Value = 0
    End If
    
    TimePast = m1.CurrentPosition
    Minutes = TimePast \ 60
    Seconds = TimePast - (Minutes * 60)
    If Seconds = "-1" Then Seconds = "0"
    
    If Seconds < 10 Then
        lblPosition.Caption = Minutes & ":0" & Seconds
    Else
        lblPosition.Caption = Minutes & ":" & Seconds
    End If
    
    Exit Sub
End Sub
