VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "www.ActiveVB.de"
   ClientHeight    =   390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   390
   ScaleWidth      =   2655
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.

'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source !

Option Explicit

'Deklaration: Globale Form-Variablen
Dim xTimer As New xTimer

Private Sub Command1_Click()
    xTimer.Calibrieren
    xTimer.Start
End Sub

Private Sub Command2_Click()
    xTimer.Halt
    ShowTime
End Sub

Private Sub Form_Load()
    'Control-Eigenschaften initialisieren
    Command1.Caption = "Start"
    Command2.Caption = "Stop"
End Sub

Private Sub ShowTime()
    MsgBox "Zeitmessung: " & Format(xTimer.RunTime, "0.00 ms")
End Sub
