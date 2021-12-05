VERSION 5.00
Begin VB.Form mainForm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rewind5s"
   ClientHeight    =   975
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "mainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtText1 
      Enabled         =   0   'False
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "mainForm.frx":058A
      Top             =   0
      Width           =   4695
   End
   Begin VB.Timer shortcutListenerTimer 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   3360
      Top             =   960
   End
   Begin VB.Menu menuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu menuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu menuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'//Rewind5s
'//Author: s0ft
'//c0dew0rth.blogspot.com
'//github.com/globalpolicy
'//09:59 PM | 4th Dec. 2021
'//Compile in P-Code because it makes use of Multiple threads

Option Explicit

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetForegroundWindow Lib "USER32" () As Long

Private Declare Sub ExitProcess Lib "kernel32.dll" (ByVal uExitCode As Long)

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long


Private Const SW_MINIMIZE As Long = 6
Private Const SW_NORMAL As Long = 1
Private Const SW_HIDE As Long = 0
Private Const SW_SHOW As Long = 5


Dim hwndForeground As Long
Dim keyPressCounterJ As Integer
Dim keyPressCounterK As Integer
Dim keyPressCounterL As Integer
Dim firstPressTime As Long

Private Sub Form_Load()
    Call RunScreenSampler
    shortcutListenerTimer.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call StopScreenPainter
    ExitProcess (0)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call Me.PopupMenu(menuMain)
    End If
End Sub

Private Sub menuHide_Click()
    Me.Hide
End Sub

Private Sub menuAbout_Click()
    MsgBox "Author: s0ft" & vbNewLine & "c0dew0rth.blogspot.com" & vbNewLine & "github.com/globalpolicy", vbOKOnly, "About"
End Sub




Private Sub shortcutListenerTimer_Timer()
    Dim key_J As Long
    Dim key_K As Long
    Dim key_L As Long
    key_J = &H4A
    key_K = &H4B
    key_L = &H4C
    
    If GetAsyncKeyState(key_J) = -32767 Then
        keyPressCounterJ = keyPressCounterJ + 1
        If keyPressCounterJ = 1 Then firstPressTime = GetTickCount()
    End If
    
    If GetAsyncKeyState(key_K) = -32767 Then
        keyPressCounterK = keyPressCounterK + 1
        If keyPressCounterK = 1 Then firstPressTime = GetTickCount()
    End If
    
    If GetAsyncKeyState(key_L) = -32767 Then
        keyPressCounterL = keyPressCounterL + 1
        If keyPressCounterL = 1 Then firstPressTime = GetTickCount()
    End If
    
    If keyPressCounterJ = 2 Then
        keyPressCounterJ = 0
        If GetTickCount() - firstPressTime < 500 Then
            Call ShowHistoryPicture
        End If
    End If
    
    If keyPressCounterK = 2 Then
        keyPressCounterK = 0
        If GetTickCount() - firstPressTime < 500 Then
            Me.Show
        End If
    End If
    
    If keyPressCounterL = 2 Then
        keyPressCounterL = 0
        If GetTickCount() - firstPressTime < 500 Then
            Call ClearHistoryPicture
        End If
    End If
End Sub

Private Sub ShowHistoryPicture()
    hwndForeground = GetForegroundWindow() '//save the current foreground window
    ShowWindow hwndForeground, SW_HIDE '//hide the current foreground window. this is in case the foreground window is a video in which case repainting is glitchy
    Call RunScreenPainter
End Sub

Private Sub ClearHistoryPicture()
    Call StopScreenPainter
    ShowWindow hwndForeground, SW_SHOW '//restore the foreground window
End Sub
