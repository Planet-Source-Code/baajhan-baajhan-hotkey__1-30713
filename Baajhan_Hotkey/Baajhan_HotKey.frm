VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Baajhan_HotKey"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Press Ctrl and ""A"" Together - Ctrl+A - note Pad should pop Up. You can also place some other code in there...."
      Height          =   720
      Left            =   285
      TabIndex        =   1
      Top             =   915
      Width           =   3705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Baajhan_HotKey"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'getasynckeystate allows you to capture two or more key
  'strokes combination at a time using the simple "And" statement.
  'is this simple enough to understand? then vote fer me.
  
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'This is to execute an external application :)
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

'Baajhan@Yahoo.com; ICQ - 73726505
'Baajhan.S.Ramanathan
'Vote fer me in PlanetSourceCode. If this is useful
'Any API queries are welcome.


'I put the piece of code in keyDown - you may put it anywhere
'Not necessary here :)
'vote fer me okay...

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  
  'getasynckeystate allows you to capture two or more key
  'strokes combination at a time using simple and statement.
  'check it with the if condition and fire your event there.!
  'is this simple enough to understand? then vote fer me.
  '                                         - Baajhan
  If GetAsyncKeyState(vbKeyControl) And GetAsyncKeyState(vbKeyA) Then
        WinExec "notepad.exe", 10
        'msgbox "Why dont You try this Too?"
    End If

End Sub

