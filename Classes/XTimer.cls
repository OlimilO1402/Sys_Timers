VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ActiveVB Class XTimer New Version 2025-03, by BAGZZLash, idiv and OlimilO
'https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancecounter
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
'Retrieves the current value of the performance counter, which is a high resolution (<1us) time stamp that can be used for time-interval measurements.
'https://learn.microsoft.com/en-us/windows/win32/api/profileapi/nf-profileapi-queryperformancefrequency
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
'Retrieves the frequency of the performance counter. The frequency of the performance counter is fixed at system boot and is consistent across all processors. Therefore, the frequency need only be queried upon application initialization, and the result can be cached.

'https://learn.microsoft.com/en-us/windows/win32/api/realtimeapiset/nf-realtimeapiset-queryunbiasedinterrupttime
'Private Declare Sub QueryUnbiasedInterruptTime Lib "kernel32" (ByRef UnbiasedTime As Currency)
'Gets the current unbiased interrupt-time count, in units of 100 nanoseconds. The unbiased interrupt-time count does not include time the system spends in sleep or hibernation.
'https://learn.microsoft.com/en-us/windows/win32/api/realtimeapiset/nf-realtimeapiset-queryunbiasedinterrupttimeprecise
'Private Declare Sub QueryUnbiasedInterruptTimePrecise Lib "kernel32" (ByRef lpUnbiasedInterruptTimePrecise As Currency)
'Gets the current unbiased interrupt-time count, in a more precise form than QueryUnbiasedInterruptTime does. The unbiased interrupt-time count does not include time the system spends in sleep or hibernation.
'Private m_Listeners As Collection 'Of IListenXTimer
Private m_Listener As IListenXTimer
Private m_Interval As Currency
Private m_Enabled  As Boolean
Private m_Abort    As Boolean
Private m_Frequ    As Currency

Private Sub Class_Initialize()
    'timeBeginPeriod 1 'eine niedrige Time Slice von 1ms beim Scheduler anfordern
    QueryPerformanceFrequency m_Frequ
End Sub
Private Sub Class_Terminate()
    Me.Enabled = False
End Sub

Friend Sub New_(Listener As IListenXTimer, Interval_ms As Single)
    Set m_Listener = Listener: m_Interval = Interval_ms
End Sub

'Copy to MNew:
'Public Function XTimer(Listener As IListenXTimer, ByVal Interval_ms As Single) As XTimer
'    Set XTimer = New XTimer: XTimer.New_ Listener, Interval_ms
'End Function
Friend Sub NewC(other As XTimer)
    m_Interval = other.Interval
    Set m_Listener = other.Listener
    'Me.Enabled = other.Disable 'Enabled
End Sub

Friend Function Copy() As XTimer
    Set Copy = New XTimer: Copy.NewC Me
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
    'returns the current Enable state and disables the timer
    Disable = m_Enabled
    Me.Enabled = False
End Function

Public Property Get Interval() As Single
    'get or set the interval in milliseconds
    Interval = m_Interval
End Property
Public Property Let Interval(ByVal Value_ms As Single)
    If Value_ms <= 0 Then Value_ms = 1
    m_Interval = Value_ms
End Property

Public Property Get FPS() As Single
    'get or set the interval in terms of Frames per second
    FPS = 1 / (m_Interval / 1000)
End Property
Public Property Let FPS(ByVal Value As Single)
    If Value <= 0 Then Value = 1
    m_Interval = 1000 / Value
End Property

Private Sub Run()
    Dim FPS As Long
    Dim StartTime   As Currency
    Dim ComputeTime As Currency
    Dim EndTime     As Currency
    Dim StartSecond As Date: StartSecond = Now
    Do Until m_Abort
        QueryPerformanceCounter StartTime
        m_Listener.XTimer
        FPS = FPS + 1
        If Now - StartSecond Then
            m_Listener.Frames FPS
            FPS = 0
            StartSecond = Now
        End If
        DoEvents
        QueryPerformanceCounter EndTime
        ComputeTime = EndTime - StartTime
        QueryPerformanceCounter StartTime
        Do
            QueryPerformanceCounter EndTime
        Loop Until m_Interval - ComputeTime <= EndTime - StartTime
    Loop
End Sub
