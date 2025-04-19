VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "Timers"
   ClientHeight    =   5175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ComboBox CmbFPS 
      Height          =   345
      Left            =   6240
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   0
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   773
      TabIndex        =   0
      Top             =   360
      Width           =   11655
   End
   Begin VB.OptionButton OptDTypLong 
      Caption         =   "Timer (Long)"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   0
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton OptDTypCurrency 
      Caption         =   "Timer (Currency)"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton BtnStop 
      Caption         =   "Stop []"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton BtnPlayPause 
      Caption         =   "Play |>"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   855
   End
   Begin VB.Label LblFPS 
      Caption         =   "FPS:"
      Height          =   255
      Left            =   7440
      TabIndex        =   7
      Top             =   60
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "FPS:"
      Height          =   375
      Left            =   5760
      TabIndex        =   5
      Top             =   60
      Width           =   375
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IListenXTimer
Private m_Thread As Thread
Private m_XTimer As XTimer
Private X As Long
Private Y As Long
Private m_Initializing As Boolean
Private Const VK_MEDIA_PLAY_PAUSE As Long = &HB3

Private Sub BtnNext_KeyDown(KeyCode As Integer, Shift As Integer):         HandleMediaPlayPause KeyCode: End Sub
Private Sub BtnPlayPause_KeyDown(KeyCode As Integer, Shift As Integer):    HandleMediaPlayPause KeyCode: End Sub
Private Sub BtnStop_KeyDown(KeyCode As Integer, Shift As Integer):         HandleMediaPlayPause KeyCode: End Sub
Private Sub CmbFPS_KeyDown(KeyCode As Integer, Shift As Integer):          HandleMediaPlayPause KeyCode: End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer):            HandleMediaPlayPause KeyCode: End Sub
Private Sub OptDTypCurrency_KeyDown(KeyCode As Integer, Shift As Integer): HandleMediaPlayPause KeyCode: End Sub
Private Sub OptDTypLong_KeyDown(KeyCode As Integer, Shift As Integer):     HandleMediaPlayPause KeyCode: End Sub
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer):        HandleMediaPlayPause KeyCode: End Sub
Private Sub HandleMediaPlayPause(KeyCode As Integer)
    If KeyCode = VK_MEDIA_PLAY_PAUSE Then BtnPlayPause_Click
End Sub

Private Sub Form_Load()
    m_Initializing = True
    Me.Caption = "Timers v" & App.Major & "." & App.Minor & "." & App.Revision
    Set m_Thread = MNew.Thread(EThreadPriority.PRIORITY_TIME_CRITICAL, 4)
    Set m_XTimer = MNew.XTimerL(Me, 1000 / 60)
    BtnPlayPause.Caption = "Play |>"
    BtnStop.Caption = "Stop []"
    FillCmbFPS CmbFPS
    CmbFPS.Text = CLng(m_XTimer.FPS)
    m_Initializing = False
End Sub

Private Sub Form_Resize()
    Dim L As Single, T As Single: T = Picture1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Picture1.Move L, T, W, H
End Sub

Private Sub FillCmbFPS(CmbFPS As ComboBox)
    With CmbFPS
        .Clear
        Dim i As Long
        For i = 1 To 1000
            CmbFPS.AddItem i
        Next
    End With
End Sub

Private Sub CmbFPS_Click()
    If m_Initializing Then Exit Sub
    m_XTimer.FPS = CSng(CmbFPS.Text)
End Sub

Private Sub BtnPlayPause_Click()
    If BtnPlayPause.Caption = "Play |>" Then
        BtnPlayPause.Caption = "Pause ||"
        m_XTimer.Enabled = True
    Else
        BtnPlayPause.Caption = "Play |>"
        m_XTimer.Enabled = False
    End If
End Sub

Private Sub BtnStop_Click()
    m_XTimer.Enabled = False
    BtnPlayPause.Caption = "Play |>"
    Reset
End Sub

Sub Reset()
    X = 0: Y = 0
    IListenXTimer_XTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BtnStop_Click
End Sub

Private Sub IListenXTimer_Frames(ByVal FPS As Long)
    LblFPS.Caption = "FPS: " & FPS & "; Interval: " & Format(m_XTimer.Interval, "0.000") & "ms"
End Sub

Private Sub IListenXTimer_XTimer()
    If X = Picture1.ScaleWidth Then X = 0
    Picture1.Cls
    For Y = 0 To Picture1.ScaleHeight - 1 Step 2
        Picture1.PSet (X, Y), vbWhite
    Next
    X = X + 1
End Sub

Private Sub OptDTypCurrency_Click()
    Dim en As Boolean: en = m_XTimer.Disable
    Set m_XTimer = MNew.XTimer(Me, m_XTimer.Interval)
    m_XTimer.Enabled = en
End Sub

Private Sub OptDTypLong_Click()
    Dim en As Boolean: en = m_XTimer.Disable
    Set m_XTimer = MNew.XTimerL(Me, m_XTimer.Interval)
    m_XTimer.Enabled = en
End Sub

