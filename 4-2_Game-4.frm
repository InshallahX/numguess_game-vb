VERSION 5.00
Begin VB.Form GUI 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton GiveupButton 
      BackColor       =   &H008080FF&
      Caption         =   "ยอมแพ้"
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton AnswerButton 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Index           =   5
      Left            =   3000
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton AnswerButton 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Index           =   4
      Left            =   3000
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton AnswerButton 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Index           =   3
      Left            =   3000
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
   End
   Begin VB.CommandButton AnswerButton 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Index           =   2
      Left            =   600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton AnswerButton 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Index           =   1
      Left            =   600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton AnswerButton 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Index           =   0
      Left            =   600
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Image Graph 
      Height          =   495
      Left            =   960
      Picture         =   "4-2_Game-4.frx":0000
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   960
      Picture         =   "4-2_Game-4.frx":28FD
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label HintDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "BrowalliaUPC"
         Size            =   36
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1080
      TabIndex        =   3
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "GUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CURRENT_Q_POS As Integer
Dim TOTAL_Q As Integer
Dim TOTAL_ANSWER_BTN As Integer
Dim ANSWER As Integer
Dim CURRENT_ANSWER_POS As Integer
Dim CORRECT_COUNT As Integer
Dim GRAPH_LEN As Double
Dim GRAPH_SCALE As Double

Sub Init()
    CURRENT_Q_POS = 1
    TOTAL_Q = SettingsPage.QAmount.Text
    TOTAL_ANSWER_BTN = AnswerButton.Count - 1
    GRAPH_LEN = 0
    ANSWER = RndAns
    CURRENT_ANSWER_POS = RndCh
    CORRECT_COUNT = 0
    Graph.Visible = False
    AnswerButton(CURRENT_ANSWER_POS).Caption = ANSWER
    GRAPH_SCALE = 3615 / TOTAL_Q
    For i = 0 To TOTAL_ANSWER_BTN
        Randomize
        If (AnswerButton(i).Caption = None) Then
            r = Int(Rnd * 1)
            If (r = 0) Then
                AnswerButton(i).Caption = Int(Rnd * ANSWER - 1)
            End If
            If (AnswerButton(i).Caption = ANSWER) Then
                AnswerButton(i).Caption = Int(Rnd * (ANSWER - 1) + 1)
            End If
        End If
    Next i
    Call Self_Refresh
End Sub

Sub Destroy()
    CURRENT_Q_POS = 0
    GRAPH_LEN = 0
    ANSWER = 0
    CURRENT_ANSWER_POS = 0
    CORRECT_COUNT = 0
    Graph.Visible = False
    GRAPH_SCALE = 0
End Sub

Sub Self_Refresh()
    For i = 0 To TOTAL_ANSWER_BTN
        Randomize
        AnswerButton(i).Caption = Int(Rnd * 100) + 10
    Next i
    ANSWER = RndAns
    CURRENT_ANSWER_POS = RndCh
    AnswerButton(CURRENT_ANSWER_POS).Caption = ANSWER
    Randomize
    HintDisplay.Caption = "น้อยกว่า " & (Int(Rnd * 50) + (Int(Rnd * 9) + 1)) + ANSWER
    Call RefreshC
End Sub

Sub RefreshC()
    If SettingsPage.PresentationMode = 1 Then
        GUI.Caption = "CurPos: " & CURRENT_Q_POS & "/" & TOTAL_Q & " " & " Answer: " & ANSWER & " Correct: " & CORRECT_COUNT
    Else
        GUI.Caption = "ข้อที่ " & CURRENT_Q_POS & "/" & TOTAL_Q & " " & " ตอบถูก " & CORRECT_COUNT & " ข้อ"
    End If
End Sub

Sub TryCheck()
    If (GRAPH_LEN < GRAPH_SCALE) Then
        Graph.Width = 0
        Graph.Visible = False
        MsgBox "คุณแพ้แล้ว"
        GUI.Hide
        SettingsPage.Show
        Call Destroy
        Exit Sub
    End If
    If (GRAPH_LEN >= 3615) Then
        MsgBox "เกมจบแล้ว ด้วยคะแนนเต็ม " & TOTAL_Q, 0, "Perfect!"
        GUI.Hide
        SettingsPage.Show
        Call Destroy
        Exit Sub
    Else
        If (CURRENT_Q_POS > TOTAL_Q) Then
            Graph.Width = 0
            Graph.Visible = False
            MsgBox "เกมจบแล้ว ด้วยคะแนน " & CORRECT_COUNT & " คะแนน"
            GUI.Hide
            SettingsPage.Show
            Call Destroy
            Exit Sub
        End If
    End If
    Call Self_Refresh
End Sub

Private Sub AnswerButton_Click(Index As Integer)
    Dim SELECTED_ANSWER As String
    SELECTED_ANSWER = Val(AnswerButton(Index).Caption)
    Graph.Visible = True
    Select Case SELECTED_ANSWER
        Case Is = ANSWER
            CURRENT_Q_POS = CURRENT_Q_POS + 1
            CORRECT_COUNT = CORRECT_COUNT + 1
            GRAPH_LEN = GRAPH_LEN + GRAPH_SCALE
            Graph.Width = GRAPH_LEN
        Case Else
            If (GRAPH_LEN <= GRAPH_SCALE) Then
                Graph.Visible = False
                GRAPH_LEN = 0
                Graph.Width = GRAPH_LEN
            Else
                CURRENT_Q_POS = CURRENT_Q_POS + 1
                GRAPH_LEN = GRAPH_LEN - GRAPH_SCALE
                Graph.Width = GRAPH_LEN
            End If
    End Select
    Call TryCheck
End Sub

Private Sub Form_Activate()
    SetFocusAPI GUI.hWnd
    Call Init
End Sub

Private Sub GiveupButton_Click()
    Call Destroy
    GUI.Hide
    SettingsPage.Show
End Sub
