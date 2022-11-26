VERSION 5.00
Begin VB.Form SettingsPage 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox PresentationMode 
      BackColor       =   &H00000000&
      Caption         =   "โหมดนำเสนอ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "CordiaUPC"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3960
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   7
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox QAmount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   12
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      MaxLength       =   2
      TabIndex        =   5
      Text            =   "10"
      Top             =   1320
      Width           =   1455
   End
   Begin VB.OptionButton graph_style 
      BackColor       =   &H00000000&
      Caption         =   "แบบที่ 2"
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   1095
   End
   Begin VB.OptionButton graph_style 
      BackColor       =   &H00000000&
      Caption         =   "แบบที่ 1"
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   1
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Image PlayButton 
      Height          =   750
      Left            =   1680
      Picture         =   "4-2_Game-3.frx":0000
      Top             =   2400
      Width           =   2250
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "เพิ่มเติม"
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   495
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "จำนวนข้อ"
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "รูปแบบกราฟ"
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   15.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF80&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "การตั้งค่า"
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   26.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "SettingsPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SETTING_GRAPH_STYLE As Integer
Dim SETTING_QUESTION_AMOUNT As Integer

Private Sub Form_Activate()
    SettingsPage.Caption = "Logged in as " & Login.UserID.Text
End Sub

Private Sub QAmount_Change()
    Dim arg As String
    arg = QAmount.Text
    If arg = "" Then
        Exit Sub
    End If
    If Not IsNumeric(arg) Then
        MsgBox "ใส่ตัวเลขเท่านั้น"
        QAmount.Text = ""
    End If
End Sub

Sub Check()
    Select Case True
        Case graph_style(0).Value
            GUI.Graph.Picture = LoadPicture("assets/bar-style-1.jpg")
        Case graph_style(1).Value
            GUI.Graph.Picture = LoadPicture("assets/bar-style-2.jpg")
    End Select
    
    Dim qa As Integer 'question amount(input)
    Dim mq As Integer 'max question
    If QAmount.Text = "" Then
        MsgBox "กรุณาใส่จำนวนข้อ"
        Exit Sub
    End If
    qa = QAmount.Text
    mq = 30
    If Len(QAmount.Text) > 0 Then
        If qa < 5 Then
            MsgBox "จำนวนข้อต้องมากกว่า 5 ข้อ"
            SetFocusAPI QAmount.hWnd
            Exit Sub
        End If
        If qa <= mq Then
            SETTING_QUESTION_AMOUNT = qa
            SettingsPage.Hide
            GUI.Show
        Else
            MsgBox "จำนวนข้อไม่สามารถมากกว่า " & mq & " ข้อ"
            SetFocusAPI QAmount.hWnd
        End If
    End If
End Sub

Private Sub PlayButton_Click()
    Call Check
End Sub

