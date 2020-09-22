VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Windows Functions"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   2400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find Dialog"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   2175
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run Dialog"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton cmdExplorer 
      Caption         =   "Windows Explorer"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2175
   End
   Begin VB.CommandButton cmdSysProps 
      Caption         =   "System Properties Panel"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CommandButton cmdUndoMinimize 
      Caption         =   "Undo Minimize"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton cmdMinimize 
      Caption         =   "Minimize all Windows"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

'   These are the 'keys' that we're going to use.
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B
Const VK_PAUSE = &H13
Const VK_SHIFT = &H10
Const VK_M = &H4D
Const VK_F = &H46
Const VK_R = &H52
Const VK_E = &H45

Private Sub cmdExplorer_Click()
'   Send the keystroke for the left Windows Key
    Call keybd_event(VK_LWIN, 0, 0, 0)
'   Send the keystroke for the E Key
    Call keybd_event(VK_E, 0, 0, 0)
'   Tell Windows to take its finger off the Windows key :)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub cmdFind_Click()
'   See cmdExplorer_Click for information
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(VK_F, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub cmdRun_Click()
'   See cmdExplorer_Click for information
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(VK_R, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub cmdMinimize_Click()
'   See cmdExplorer_Click for information
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(VK_M, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub cmdSysProps_Click()
'   See cmdExplorer_Click for information
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(VK_PAUSE, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub cmdUndoMinimize_Click()
'   See cmdExplorer_Click for information
    Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(VK_SHIFT, 0, 0, 0)
    Call keybd_event(VK_M, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
    Call keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0)
End Sub
