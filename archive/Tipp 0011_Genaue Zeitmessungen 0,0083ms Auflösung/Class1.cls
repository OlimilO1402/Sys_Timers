VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

'Deklaration: Globale Klassen API-Typen
Private Type LARGE_INTEGER
    Lo As Long
    Hi As Long
End Type

'Deklaration: Globale Klassen API-Funktionen
Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
    lpPerformanceCount As LARGE_INTEGER) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
    lpFrequency As LARGE_INTEGER) As Long

'Deklaration: Globale Klassen-Variablen
Dim dblCalibrieren As Double

Dim udtStart As LARGE_INTEGER
Dim udtEnde As LARGE_INTEGER
Dim udtFreq As LARGE_INTEGER

Public Sub Calibrieren()
    Call QueryPerformanceCounter(udtStart)
    Call QueryPerformanceCounter(udtEnde)

    dblCalibrieren = (D(udtEnde) - D(udtStart)) / D(udtFreq) * 1000
End Sub

Private Sub Class_Initialize()
    Call QueryPerformanceFrequency(udtFreq)
End Sub

Public Sub Halt()
    Call QueryPerformanceCounter(udtEnde)
End Sub

Public Sub Start()
    Call QueryPerformanceCounter(udtStart)
End Sub

Public Property Get RunTime() As Double
    RunTime = (D(udtEnde) - D(udtStart)) / D(udtFreq) * 1000 - dblCalibrieren
End Property

Private Function D(udtX As LARGE_INTEGER) As Double
    'Deklaration: Lokale Prozedur-Variablen
    Dim dblHigh As Double
    Dim dblLow As Double

    dblLow = udtX.Lo
    dblHigh = udtX.Hi

    If dblLow < 0 Then
        dblLow = 4294967296# + dblLow + 1
    End If

    If dblHigh < 0 Then
        dblHigh = 4294967296# + dblHigh + 1
    End If

    D = dblLow + dblHigh * 4294967296#
End Function


