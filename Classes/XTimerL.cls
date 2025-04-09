VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XTimerL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ActiveVB Class XTimer New Version 2025-03, by BAGZZLash, idiv and OlimilO
Implements XTimer
'https://learn.microsoft.com/en-us/windows/win32/api/timeapi/nf-timeapi-timebeginperiod
Private Declare Function timeBeginPeriod Lib "winmm" (ByVal uPeriod As Long) As Long
'The timeBeginPeriod function requests a minimum resolution for periodic timers
'https://learn.microsoft.com/en-us/windows/win32/api/timeapi/nf-timeapi-timegettime
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
'The timeGetTime function retrieves the system time, in milliseconds. The system time is the time elapsed since Windows was started.
'Private m_Listeners As Collection 'Of IListenXTimer






Private m_Listener As IListenXTimer
Private m_Interval As Single
Private m_Enabled  As Boolean
Private m_Abort    As Boolean


Private Sub Class_Initialize()
    timeBeginPeriod 1 'get a low time-slice of 1ms from time-scheduler
End Sub
Private Sub Class_Terminate()
    Me.Enabled = False
End Sub

Friend Sub New_(Listener As IListenXTimer, Interval_ms As Single)
    Set m_Listener = Listener: m_Interval = Interval_ms
End Sub

'Copy to MNew:
'Public Function XTimerL(Listener As IListenXTimer, ByVal Interval_ms As Single) As XTimerL
'    Set XTimerL = New XTimerL: XTimerL.New_ Listener, Interval_ms
'End Function
Friend Sub NewC(other As XTimer)
    m_Interval = other.Interval
    Set m_Listener = other.Listener
End Sub

Friend Function Copy() As XTimerL
    Set Copy = New XTimerL: Copy.NewC Me
End Function

Public Property Get Listener() As IListenXTimer
    Set Listener = m_Listener
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property
Public Property Let Enabled(ByVal Value As Boolean)
    m_Enabled = Value: m_Abort = Not m_Enabled
    If m_Enabled Then Run
End Property

Public Function Disable() As Boolean
    Disable = m_Enabled
    Me.Enabled = False
End Function

Public Property Get Interval() As Single
    Interval = m_Interval
End Property
Public Property Let Interval(ByVal Value_ms As Single)
    If Value_ms <= 0 Then Value_ms = 1
    m_Interval = Value_ms
End Property

Public Property Get FPS() As Single
    FPS = 1 / (m_Interval / 1000)
End Property
Public Property Let FPS(ByVal Value As Single)
    If Value <= 0 Then Value = 1
    m_Interval = 1000 / Value
End Property

Private Sub Run()
    Dim FPS As Long
    Dim StartTime   As Long
    Dim ComputeTime As Long
    Dim EndTime     As Long
    Dim StartSecond As Date: StartSecond = Now
    Do Until m_Abort
        StartTime = timeGetTime()
        m_Listener.XTimer
        FPS = FPS + 1
        If Now - StartSecond Then
            m_Listener.Frames FPS
            FPS = 0
            StartSecond = Now
        End If
        DoEvents
        EndTime = timeGetTime
        ComputeTime = EndTime - StartTime
        StartTime = timeGetTime
        Do
            EndTime = timeGetTime()
        Loop Until m_Interval - ComputeTime <= EndTime - StartTime
    Loop
End Sub

' v ' ############################## ' v '    Implements XTimer    ' v ' ############################## ' v '
Private Function XTimer_Disable() As Boolean
    XTimer_Disable = Me.Disable
End Function

Private Property Get XTimer_Enabled() As Boolean
    XTimer_Enabled = Me.Enabled
End Property
Private Property Let XTimer_Enabled(ByVal RHS As Boolean)
    Me.Enabled = RHS
End Property

Private Property Get XTimer_FPS() As Single
    XTimer_FPS = Me.FPS
End Property
Private Property Let XTimer_FPS(ByVal RHS As Single)
    Me.FPS = RHS
End Property

Private Property Get XTimer_Interval() As Single
    XTimer_Interval = Me.Interval
End Property
Private Property Let XTimer_Interval(ByVal RHS As Single)
    Me.Interval = RHS
End Property

Private Property Get XTimer_Listener() As IListenXTimer
    Set XTimer_Listener = Me.Listener
End Property
' ^ ' ############################## ' ^ '    Implements XTimer    ' ^ ' ############################## ' ^ '

