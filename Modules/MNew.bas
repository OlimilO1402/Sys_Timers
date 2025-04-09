Attribute VB_Name = "MNew"
Option Explicit


Public Function XTimer(Listener As IListenXTimer, ByVal Interval_ms As Single) As XTimer
    Set XTimer = New XTimer: XTimer.New_ Listener, Interval_ms
End Function
Public Function XTimerL(Listener As IListenXTimer, ByVal Interval_ms As Single) As XTimerL
    Set XTimerL = New XTimerL: XTimerL.New_ Listener, Interval_ms
End Function

Public Function Thread(ByVal aPriority As EThreadPriority, ByVal aAffinityMask As Long) As Thread
    Set Thread = New Thread: Thread.New_ aPriority, aAffinityMask
End Function

