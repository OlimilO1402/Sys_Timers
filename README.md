# Sys_Timers  
## high accurate Timer for animations with a max resolution of about 1000 FPS ~ 1.0ms  
  
[![GitHub](https://img.shields.io/github/license/OlimilO1402/Sys_Timers?style=plastic)](https://github.com/OlimilO1402/Sys_Timers/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Sys_Timers?style=plastic)](https://github.com/OlimilO1402/Sys_Timers/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Sys_Timers/total.svg)](https://github.com/OlimilO1402/Sys_Timers/releases/download/v2025.4.9/Timers_v2025.4.9.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)


Project started around 2000.  
This example shows how to do animations without a Timer control. The VBs intrinsic Timer Control has some disadvantages.  
The Timer control is not very accurate, the minimum resolution is about 55ms, it uses a windows event, so it is not very stable and it does not even exist in VBA.  
Maybe you remember the class XTimer from ActiveVB. This repo contains the improved version of the XTimer class.   
In fact there are 2 classes both share the same Interface, both are interchangeable even during timer-runtime.   
* XTimerL uses the API function timeGetTime together with a locale variable of datatype Long  
* XTimer  uses the API function QueryPerformanceCounter with a locale variable of datatype Currency  
    
How does it work?    
The new XTimer does not make use of Windows Events, instead it works with the Listener pattern.  
This is well known in the Java-world, you can do this in VB the same way, just by using an interface.  
There is the Interface IListenXTimer with 2 function stubs "Sub Frames(FPS)" and "Sub XTimer()" every object who wants to "listen" to XTimer-messages has to implement this interface.  
```vba
Interface IListenXTimer
Public Sub XTimer()
Public Sub Frames(ByVal FPS As Long)
```

"Sub Frames" fires every second and is just for displaying the frames per second. "Sub XTimer" fires of course every interval.  
The property Interval is of datatype Single to get or set the timer-interval in milliseconds. But you could also use the property FPS.
```vba
class XTimer
Public Property Get Interval() As Single
    'get or set the interval in milliseconds
    Interval = m_Interval
End Property
Public Property Let Interval(ByVal Value_ms As Single)
    If Value_ms <= 0 Then Value_ms = 1
    m_Interval = Value_ms
End Property
Public Property Get FPS() As Single
    'get or set the interval in terms of frames per second
    FPS = 1 / (m_Interval / 1000)
End Property
Public Property Let FPS(ByVal Value As Single)
    If Value <= 0 Then Value = 1
    m_Interval = 1000 / Value
End Property
```
what is the difference between XTimer and XTimerL?  
With the class XTimerL, if you set 450 Frames per second, it does just about 333 FPS,   
because 450 fps are 2.222 ms so it rounds upt to 3 ms. Switch to the "Timer (Currency)" and it does the real 450 FPS  
Have a look at the repo [Sys_Stopwatch](https://github.com/OlimilO1402/Sys_StopWatch) 

![Timers Image](Resources/Timers.png "Timers Image")