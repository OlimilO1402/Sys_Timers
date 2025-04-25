VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Atari Space Race 1973"
   ClientHeight    =   7335
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   489
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "Restart"
      Height          =   435
      Left            =   4200
      TabIndex        =   11
      Top             =   6720
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      Caption         =   "Player 2"
      Height          =   1095
      Left            =   7440
      TabIndex        =   5
      Top             =   6120
      Width           =   2055
      Begin VB.Label Label2 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   8
         Top             =   660
         Width           =   345
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Controls: Up/Down"
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1665
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Points:"
         Height          =   255
         Left            =   180
         TabIndex        =   6
         Top             =   660
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player 1"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   6120
      Width           =   2055
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Points:"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   660
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Controls: W/S"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   660
         Width           =   345
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   0
      ScaleHeight     =   792.079
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Round: "
      Height          =   255
      Left            =   5880
      TabIndex        =   14
      Top             =   6360
      Width           =   675
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "FPS: "
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   6720
      Width           =   405
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Asteroids: "
      Height          =   255
      Left            =   2400
      TabIndex        =   12
      Top             =   6360
      Width           =   930
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      Caption         =   "Time remaining:"
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   6060
      Width           =   1410
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Zentriert
      Caption         =   "45"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   6300
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Implements IListenXTimer

Private Declare Sub RtlMoveMemory Lib "kernel32" (pDst As Any, pSrc As Any, ByVal bytlength As Long)
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long
Private Declare Function timeBeginPeriod Lib "Winmm.dll" (ByVal uPeriod As Long) As Long
Private Declare Function timeGetTime Lib "Winmm.dll" () As Long
Private Declare Function GetCurrentThread Lib "kernel32.dll" () As Long
Private Declare Function SetThreadPriority Lib "kernel32.dll" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Private Declare Function SetThreadAffinityMask Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwThreadAffinityMask As Long) As Long

Private Type Rocket
    X As Long
    Y As Long
    Points As Long
End Type

Private Type Star
    X As Long
    Y As Long
    GoingRight As Boolean
End Type

Private Const THREAD_BASE_PRIORITY_LOWRT As Long = 15&
Private Const THREAD_BASE_PRIORITY_MAX As Long = 2&
Private Const THREAD_PRIORITY_HIGHEST As Long = THREAD_BASE_PRIORITY_MAX
Private Const THREAD_PRIORITY_TIME_CRITICAL As Long = THREAD_BASE_PRIORITY_LOWRT

Private Const VK_W = &H57
Private Const VK_S = &H53
Private Const VK_LBUTTON = &H1
Private Const VK_RBUTTON = &H2
Private Const VK_UP = &H26
Private Const VK_DOWN = &H28

Private Const FRAMEWIDTH As Long = 640
Private Const FRAMEHEIGHT As Long = 400

Private TIMERINTERVAL As Long '= 20
Private NUMSTARS As Long '= 32 '36
Private MAXTIME  As Long

Private Const ROCKETWIDTH As Long = 20
Private Const ROCKETHEIGHT As Long = 28
Private Const ROCKETHEADROOM As Long = 5

Private Const DELAYFRAMES As Long = 75

Private Player1 As Rocket
Private Player2 As Rocket
Private Starfield() As Star
Private RespawnDelay1 As Boolean
Private RespawnDelay2 As Boolean
Private RespawnDelay1Counter As Long
Private RespawnDelay2Counter As Long
Private Rounds As Byte
Private XTimer As XTimer

Private Sub HandleInput()

Dim RetVal As Long
Dim P1Done As Boolean

RetVal = GetAsyncKeyState(VK_W)
If CBool(RetVal And &H8000) Then
    Call P1Up
    P1Done = True
End If

RetVal = GetAsyncKeyState(VK_S)
If CBool(RetVal And &H8000) Then
    Call P1Down
    P1Done = True
End If

If Not P1Done Then
    RetVal = GetAsyncKeyState(VK_LBUTTON)
    If CBool(RetVal And &H8000) Then
        Call P1Up
    End If
    
    RetVal = GetAsyncKeyState(VK_RBUTTON)
    If CBool(RetVal And &H8000) Then
        Call P1Down
    End If
End If

RetVal = GetAsyncKeyState(VK_UP)
If CBool(RetVal And &H8000) Then
    Call P2Up
End If

RetVal = GetAsyncKeyState(VK_DOWN)
If CBool(RetVal And &H8000) Then
    Call P2Down
End If

End Sub

Private Sub P1Up()

If RespawnDelay1 Then Exit Sub

If Player1.Y <= 0 Then
    Player1.Points = Player1.Points + 1
    Form1.Label1 = Player1.Points
    Call Reset1
    Exit Sub
End If
Player1.Y = Player1.Y - 1

End Sub

Private Sub P1Down()

If Player1.Y >= FRAMEHEIGHT - ROCKETHEIGHT - ROCKETHEADROOM Then Exit Sub
Player1.Y = Player1.Y + 1

End Sub

Private Sub P2Up()

If RespawnDelay2 Then Exit Sub

If Player2.Y <= 0 Then
    Player2.Points = Player2.Points + 1
    Form1.Label2 = Player2.Points
    Call Reset2
    Exit Sub
End If
Player2.Y = Player2.Y - 1

End Sub

Private Sub P2Down()

If Player2.Y >= FRAMEHEIGHT - ROCKETHEIGHT - ROCKETHEADROOM Then Exit Sub
Player2.Y = Player2.Y + 1

End Sub

Private Sub Reset1()

Player1.X = (FRAMEWIDTH \ 3) - (ROCKETWIDTH \ 2)
Player1.Y = FRAMEHEIGHT - ROCKETHEIGHT - ROCKETHEADROOM

End Sub

Private Sub Reset2()

Player2.X = ((FRAMEWIDTH \ 3) * 2) - (ROCKETWIDTH \ 2)
Player2.Y = FRAMEHEIGHT - ROCKETHEIGHT - ROCKETHEADROOM

End Sub

Private Sub Command1_Click()
Set XTimer = New XTimer: XTimer.New_ Me, 1000 / TIMERINTERVAL
Label9.Caption = "Asteroids: " & NUMSTARS
Label10.Caption = "FPS: " & TIMERINTERVAL
Label11.Caption = "Round: " & Rounds
Form1.Command1.Enabled = False

Call Reset1
'Player1.Points = 0
Call Reset2
'Player2.Points = 0

Call InitStarfield

Form1.Label1 = Player1.Points '"0"
Form1.Label2 = Player2.Points '"0"
Form1.Label8 = MAXTIME

RespawnDelay1 = False
RespawnDelay2 = False

'Call MainLoop
XTimer.Enabled = True
TIMERINTERVAL = TIMERINTERVAL * 1.2
NUMSTARS = NUMSTARS * 1.2
MAXTIME = MAXTIME * 1.2
End Sub

Private Sub Form_Load()
Me.Caption = "Atari Space Race 1973 v" & App.Major & "." & App.Minor & "." & App.Revision
TIMERINTERVAL = 20
NUMSTARS = 30
MAXTIME = 60
Label8.Caption = MAXTIME

Dim FileNum As Integer
Dim hThread As Long
Dim RetVal As Long

Randomize

Call timeBeginPeriod(1)
hThread = GetCurrentThread
RetVal = SetThreadPriority(hThread, THREAD_PRIORITY_TIME_CRITICAL)
RetVal = SetThreadAffinityMask(hThread, 1)

Call Init(Form1.Picture1)
Call InitStarfield

ReDim RocketBitmap(ROCKETWIDTH * ROCKETHEIGHT)

FileNum = FreeFile
Dim pfn As String: pfn = App.Path & "\Rocket.bin"
If FileExists(pfn) Then
    Open pfn For Binary As FileNum
        Get #FileNum, , RocketBitmap
    Close
Else
    Dim bytes() As Byte: bytes = LoadResData(2, "CUSTOM")
    RtlMoveMemory RocketBitmap(0), bytes(0), UBound(bytes) + 1
End If

Call Reset1
Call Reset2
Set XTimer = New XTimer: XTimer.New_ Me, 1000 / TIMERINTERVAL
Rounds = 1
End Sub

Private Function FileExists(ByVal FileName As String) As Boolean
    On Error Resume Next
    FileExists = Not CBool(GetAttr(FileName) And (vbDirectory Or vbVolume))
    On Error GoTo 0
End Function

Private Function MakeStar(GoRight As Boolean) As Star

Dim MyStar As Star

MyStar.X = Rnd * FRAMEWIDTH
MyStar.Y = Rnd * (FRAMEHEIGHT - ROCKETHEIGHT - ROCKETHEADROOM - 2)
MyStar.GoingRight = GoRight

MakeStar = MyStar

End Function

Private Sub InitStarfield()

Dim i As Long

ReDim Starfield(0 To NUMSTARS - 1)

For i = 0 To (NUMSTARS \ 2) - 1
    Starfield(i) = MakeStar(True)
Next

For i = (NUMSTARS \ 2) To NUMSTARS - 1
    Starfield(i) = MakeStar(False)
Next

End Sub

Private Sub UpdateStarfield()

Dim i As Long

For i = 0 To UBound(Starfield)
    If Starfield(i).GoingRight Then
        Starfield(i).X = Starfield(i).X + 1
        If Starfield(i).X >= FRAMEWIDTH Then
            Starfield(i) = MakeStar(True)
            Starfield(i).X = 0
        End If
    Else
        Starfield(i).X = Starfield(i).X - 1
        If Starfield(i).X <= 0 Then
            Starfield(i) = MakeStar(False)
            Starfield(i).X = FRAMEWIDTH - 1
        End If
    End If
Next

End Sub

Private Sub DrawRockets()

Dim XPix As Long
Dim YPix As Long
Dim SourceOffset As Long
Dim TargetOffset As Long

For XPix = 0 To ROCKETWIDTH - 1
    For YPix = 0 To ROCKETHEIGHT - 1
        SourceOffset = (YPix * ROCKETWIDTH) + XPix
        If Not (RocketBitmap(SourceOffset).Red = 255 And RocketBitmap(SourceOffset).Green = 0 And RocketBitmap(SourceOffset).Blue = 255) Then
            If Not RespawnDelay1 Then
                TargetOffset = ((YPix + Player1.Y) * FRAMEWIDTH) + (XPix + Player1.X)
                FrameBuffer(TargetOffset) = RocketBitmap(SourceOffset)
            End If
            If Not RespawnDelay2 Then
                TargetOffset = ((YPix + Player2.Y) * FRAMEWIDTH) + (XPix + Player2.X)
                FrameBuffer(TargetOffset) = RocketBitmap(SourceOffset)
            End If
        End If
    Next
Next

End Sub

Private Sub DrawStarfield()

Dim i As Long
Dim Offset As Long

For i = 0 To UBound(Starfield)
    Offset = (Starfield(i).Y * FRAMEWIDTH) + Starfield(i).X
    FrameBuffer(Offset).Red = 200
    FrameBuffer(Offset).Green = 200
    FrameBuffer(Offset).Blue = 255
    If Starfield(i).GoingRight Then
        Offset = (Starfield(i).Y * FRAMEWIDTH) + (Starfield(i).X - 1)
        If Offset < 0 Then Offset = 0
        FrameBuffer(Offset).Red = 192
        FrameBuffer(Offset).Green = 192
        FrameBuffer(Offset).Blue = 32
    Else
        Offset = (Starfield(i).Y * FRAMEWIDTH) + (Starfield(i).X + 1)
        FrameBuffer(Offset).Red = 192
        FrameBuffer(Offset).Green = 192
        FrameBuffer(Offset).Blue = 32
    End If
Next

End Sub

Private Function CheckCollision() As Long

Dim XPix As Long
Dim YPix As Long
Dim i As Long
Dim SourceOffset As Long

For XPix = 0 To ROCKETWIDTH - 1
    For YPix = 0 To ROCKETHEIGHT - 1
        SourceOffset = (YPix * ROCKETWIDTH) + XPix
        If Not (RocketBitmap(SourceOffset).Red = 255 And RocketBitmap(SourceOffset).Green = 0 And RocketBitmap(SourceOffset).Blue = 255) Then
            For i = 0 To UBound(Starfield)
                If Starfield(i).X = XPix + Player1.X And Starfield(i).Y = YPix + Player1.Y Then
                    CheckCollision = 1
                    Exit Function
                End If
                If Starfield(i).X = XPix + Player2.X And Starfield(i).Y = YPix + Player2.Y Then
                    CheckCollision = 2
                    Exit Function
                End If
            Next
        End If
    Next
Next

End Function

Friend Sub MakeFrame()

Dim Collision As Long

FrameBuffer = Background

Call UpdateStarfield
Call DrawStarfield

If RespawnDelay1 Then
    If RespawnDelay1Counter >= DELAYFRAMES Then
        RespawnDelay1Counter = 0
        RespawnDelay1 = False
        Call Reset1
    Else
        RespawnDelay1Counter = RespawnDelay1Counter + 1
    End If
End If

If RespawnDelay2 Then
    If RespawnDelay2Counter >= DELAYFRAMES Then
        RespawnDelay2Counter = 0
        RespawnDelay2 = False
        Call Reset2
    Else
        RespawnDelay2Counter = RespawnDelay2Counter + 1
    End If
End If

Collision = CheckCollision
If Collision = 1 Then RespawnDelay1 = True
If Collision = 2 Then RespawnDelay2 = True

Call DrawRockets

Call Draw(Form1.Picture1)

End Sub

Private Sub IListenXTimer_Frames(ByVal FPS As Long)
    Form1.Label8 = Form1.Label8 - 1
    If Form1.Label8 < 0 Then Form1.Label8 = 0
End Sub

Private Sub IListenXTimer_XTimer()
    Call HandleInput
    Call MakeFrame
    If Form1.Label8 = "0" Then
        XTimer.Enabled = False
        If Player1.Points <> Player2.Points Then
            MsgBox "Player " & IIf(Player1.Points > Player2.Points, "1", "2") & " wins.", vbInformation
        Else
            MsgBox "The game is a tie.", vbInformation
        End If
        Rounds = Rounds + 1
        Form1.Command1.Enabled = True
        Command1.SetFocus
        Exit Sub
    End If
    DoEvents
End Sub

'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Dim en As Boolean: If Not XTimer Is Nothing Then en = XTimer.Disable
'    Dim mr As VbMsgBoxResult: mr = MsgBox("Are you sure you want to quit?", vbOKCancel)
'    If mr = vbCancel Then
'        Cancel = 1
'        If Not XTimer Is Nothing Then XTimer.Enabled = en
'    'Else
'    End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not XTimer Is Nothing Then XTimer.Enabled = False
    Set XTimer = Nothing
    End
End Sub
