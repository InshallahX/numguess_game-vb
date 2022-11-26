VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form Main 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "4-2_Game-1.frx":0000
   ScaleHeight     =   3360
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton DevButton 
      BackColor       =   &H00FFFFC0&
      Caption         =   "จัดทำโดย"
      BeginProperty Font 
         Name            =   "CordiaUPC"
         Size            =   20.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   495
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Image Image1 
      Height          =   750
      Left            =   1560
      Picture         =   "4-2_Game-1.frx":5EEC42
      Top             =   480
      Width           =   2250
   End
   Begin VB.Image TutorialButton 
      Height          =   750
      Left            =   3000
      Picture         =   "4-2_Game-1.frx":5F3251
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Image StartButton 
      Height          =   750
      Left            =   720
      Picture         =   "4-2_Game-1.frx":5F6E70
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1500
   End
   Begin VB.Image BackgroundImg1 
      Height          =   4500
      Left            =   1920
      Picture         =   "4-2_Game-1.frx":5FAF91
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Main.Caption = "Welcome"
    SetFocusAPI Main.hWnd
    WindowsMediaPlayer1.URL = "assets\music.mp3"
End Sub

Private Sub StartButton_Click()
    Main.Hide
    Login.Show
End Sub

Private Sub TutorialButton_Click()
    Call Shell("explorer.exe " & "https://inshallah.cc/app/vb-game/")
End Sub

Private Sub DevButton_Click()
    Call Shell("explorer.exe " & "https://inshallah.cc/app/vb-game/dev")
End Sub
