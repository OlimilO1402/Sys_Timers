VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const THREAD_BASE_PRIORITY_IDLE     As Long = -15&
Private Const THREAD_BASE_PRIORITY_MIN      As Long = -2&
Private Const THREAD_BASE_PRIORITY_MAX      As Long = 2&
Private Const THREAD_BASE_PRIORITY_LOWRT    As Long = 15&

Private Const THREAD_PRIORITY_IDLE          As Long = THREAD_BASE_PRIORITY_IDLE
Private Const THREAD_PRIORITY_LOWEST        As Long = THREAD_BASE_PRIORITY_MIN
Private Const THREAD_PRIORITY_BELOW_NORMAL  As Long = (THREAD_PRIORITY_LOWEST + 1)
Private Const THREAD_PRIORITY_NORMAL        As Long = 0
Private Const THREAD_PRIORITY_HIGHEST       As Long = THREAD_BASE_PRIORITY_MAX
Private Const THREAD_PRIORITY_ABOVE_NORMAL  As Long = (THREAD_PRIORITY_HIGHEST - 1)
Private Const THREAD_PRIORITY_TIME_CRITICAL As Long = THREAD_BASE_PRIORITY_LOWRT

Private Const THREAD_PRIORITY_ERROR_RETURN  As Long = 2147483647 '(MAXLONG) '2147483647 ?

Public Enum EThreadPriority
    PRIORITY_IDLE = -15
    PRIORITY_LOWEST = -2
    PRIORITY_BELOW_NORMAL = -1
    PRIORITY_NORMAL = 0
    PRIORITY_ABOVE_NORMAL = 1
    PRIORITY_HIGHEST = 2
    PRIORITY_TIME_CRITICAL = 15
End Enum

Private Const THREAD_GET_CONTEXT            As Long = (&H8)
Private Const THREAD_IDLE_TIMEOUT           As Long = 10
Private Const THREAD_IMPERSONATE            As Long = (&H100)
Private Const THREAD_DIRECT_IMPERSONATION   As Long = (&H200)

Private Const THREAD_TERMINATE              As Long = (&H1)
Private Const THREAD_SUSPEND_RESUME         As Long = (&H2)
Private Const THREAD_SET_CONTEXT            As Long = (&H10)
Private Const THREAD_SET_INFORMATION        As Long = (&H20)
Private Const THREAD_QUERY_INFORMATION      As Long = (&H40)
Private Const THREAD_SET_THREAD_TOKEN       As Long = (&H80)

'Private Const THREAD_ALL_ACCESS             As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)


'https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-getcurrentthread
Private Declare Function GetCurrentThread Lib "kernel32" () As LongPtr

'https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-getcurrentthreadid
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

'https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-getexitcodethread
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As LongPtr, lpExitCode_out As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-getthreadpriority
Private Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As LongPtr) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-setthreadpriority
Private Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As LongPtr, ByVal nPriority As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-getprocessaffinitymask
'BOOL GetProcessAffinityMask(
'  [in]  HANDLE     hProcess,
'  [out] PDWORD_PTR lpProcessAffinityMask,
'  [out] PDWORD_PTR lpSystemAffinityMask
');
Private Declare Function GetProcessAffinityMask Lib "kernel32" (ByVal hProcess As LongPtr, ByVal dwThreadAffinityMask As Long) As Long

'https://learn.microsoft.com/en-us/windows/win32/api/winbase/nf-winbase-setthreadaffinitymask
Private Declare Function SetThreadAffinityMask Lib "kernel32" (ByVal hThread As LongPtr, ByVal dwThreadAffinityMask As Long) As Long

Private m_hThread As LongPtr
Private m_ThrId   As Long
Private m_AffinityMask As Long

Private Sub Class_Initialize()
    m_hThread = GetCurrentThread
    m_ThrId = GetCurrentThreadId
    m_AffinityMask = SetThreadAffinityMask(m_hThread, 4095)
    Dim tam As Long: tam = SetThreadAffinityMask(m_hThread, m_AffinityMask)
    If tam <> m_AffinityMask Then m_AffinityMask = tam
End Sub

Friend Sub New_(aPriority As EThreadPriority, aAffinityMask As Long)
    Me.Priority = aPriority: Me.AffinityMask = aAffinityMask
End Sub

'Copy to MNew:
'Public Function Thread(ByVal aPriority As EThreadPriority, ByVal aAffinityMask As Long) As Thread
'    Set Thread = New Thread: Thread.New_ aPriority, aAffinityMask
'End Function

Public Property Get ExitCode() As Long
    Dim hr As Long: hr = GetExitCodeThread(m_hThread, ExitCode)
End Property

Public Property Get Handle() As LongPtr
    Handle = m_hThread
End Property

Public Property Get Id() As Long
    Id = m_ThrId
End Property

Public Property Get AffinityMask() As Long
    AffinityMask = m_AffinityMask
End Property
Public Property Let AffinityMask(ByVal Value As Long)
    Dim hr As Long: hr = SetThreadAffinityMask(m_hThread, Value)
    m_AffinityMask = Value
End Property

Public Property Get Priority() As EThreadPriority
    Priority = GetThreadPriority(m_hThread)
End Property
Public Property Let Priority(ByVal Value As EThreadPriority)
    Dim hr As Long: hr = SetThreadPriority(m_hThread, Value)
End Property

Public Function Priority_ToStr() As String
    Priority_ToStr = EThreadPriority_ToStr(Me.Priority)
End Function

Private Function EThreadPriority_ToStr(e As EThreadPriority) As String
    Dim s As String
    Select Case e
    Case EThreadPriority.PRIORITY_IDLE:          s = "IDLE"
    Case EThreadPriority.PRIORITY_LOWEST:        s = "LOWEST"
    Case EThreadPriority.PRIORITY_BELOW_NORMAL:  s = "BELOW_NORMAL"
    Case EThreadPriority.PRIORITY_NORMAL:        s = "NORMAL"
    Case EThreadPriority.PRIORITY_ABOVE_NORMAL:  s = "ABOVE_NORMAL"
    Case EThreadPriority.PRIORITY_HIGHEST:       s = "HIGHEST"
    Case EThreadPriority.PRIORITY_TIME_CRITICAL: s = "TIME_CRITICAL"
    End Select
    EThreadPriority_ToStr = s
End Function
