VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Folder"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3255
   StartUpPosition =   1  'CenterOwner
   Begin Project1.chameleonButton cmdCancel 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Cancel"
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
      MICON           =   "Form2.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Project1.chameleonButton cmdOK 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "OK"
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
      MICON           =   "Form2.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim File As File
    Dim Files As Files
    Dim Folder As Folder
    Dim fso As New FileSystemObject
    Dim strm As TextStream
    
    Set Folder = fso.GetFolder(Dir1.Path)
    Set Files = Folder.Files
    Set strm = fso.OpenTextFile(App.Path & "\Settings\" & Mid(Form1.cboLists.SelectedItem.Key, 5, Len(Form1.cboLists.SelectedItem.Key) - 4) & ".BartNet", ForAppending)
    
    For Each File In Files
        If Right(File.Name, 4) = ".mp3" Or Right(File.Name, 4) = ".wav" Or Right(File.Name, 4) = ".wma" Then
            strm.WriteLine File.Path
            strm.WriteLine File.Name
            Form1.l1.ListItems.Add , File.Path, File.Name, "Icon", "Icon"
        Else
        
        End If
    Next
    
    strm.Close
    
    Unload Me
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:\"
    Dir1.Path = "c:\"
End Sub
