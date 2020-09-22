VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Window Name"
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnableWindow& Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long)
  Private Declare Function FindWindow& Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String)
Function DisWin(WindowName$, EnabOrDisab&) 'EX: Call DisWin("mIRC32", 0)
Dim lFndWnd As Long
Dim lDisEnWnd As Long

lFndWnd = FindWindow(vbNullString, WindowName$) 'Finds the Window Name
lDisEnWnd = EnableWindow(lFndWnd, ByVal EnabOrDisab&) 'Disables all mouse and keyboard input to the specified window.
                                                 'In ByVal EnabOrDisab& you either enter: 0 to Disable Window or 1 to Enable it.
End Function

Private Sub Command1_Click()
Call DisWin(Text1.Text, 0)
End Sub
