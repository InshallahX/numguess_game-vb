VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BackButton 
      BackColor       =   &H00C0C0FF&
      Caption         =   "กลับ"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton LoginButton 
      BackColor       =   &H00C0FFC0&
      Caption         =   "เข้าสู่ระบบ"
      BeginProperty Font 
         Name            =   "Leelawadee UI"
         Size            =   14.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox UserID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   21.75
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "ID"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   4020
      Left            =   0
      Picture         =   "4-2_Game-2.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//DEPRECATED Dim Con As New ADODB.Connection
'//DEPRECATED Dim Rec As New ADODB.Recordset

Private Sub Form_Load()
    Login.Caption = "                           Authenticator"
    SetFocusAPI Login.hWnd
End Sub

Private Sub BackButton_Click()
    Login.Hide
    Main.Show
End Sub

Private Sub LoginButton_Click()
    If InStr(UserID.Text, " ") Then
        MsgBox "ไม่สามารถเว้นวรรคได้"
        Exit Sub
    End If
    If UserID.Text = "ID" Then
        MsgBox "ไอดีต้องเป็นตัวเลข 5 หลัก"
        Exit Sub
    End If
    If Len(UserID.Text) = 5 Then
        Login.Hide
        SettingsPage.Show
        If UserID.Text = "01337" Then
            SettingsPage.PresentationMode.Enabled = True
        End If
    Else
        MsgBox "ไอดีต้องเป็นตัวเลข 5 หลัก"
    End If
End Sub

Private Sub UserID_Change()
    UserID.PasswordChar = "*"
    Dim arg As String
    arg = UserID.Text
    If arg = "" Then
        Exit Sub
    End If
    If Not IsNumeric(arg) Then
        MsgBox "ไอดีจะต้องเป็นตัวเลขเท่านั้น"
        UserID.Text = ""
    End If
End Sub

Private Sub UserID_GotFocus()
    If UserID.Text = "ID" Then
        UserID.Text = ""
    End If
End Sub
