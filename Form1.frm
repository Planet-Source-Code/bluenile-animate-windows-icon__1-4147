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
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   525
      Index           =   2
      Left            =   720
      Picture         =   "FORM1.frx":0000
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   60
      Width           =   555
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Index           =   1
      Left            =   1320
      Picture         =   "FORM1.frx":0442
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   270
      Top             =   2010
   End
   Begin VB.PictureBox Picture1 
      Height          =   525
      Index           =   0
      Left            =   120
      Picture         =   "FORM1.frx":0884
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Const GCL_HICON = (-14)
' please note the pictures in Picture1 are icons

Private Sub Form_Load()
'Check for previous instance
If App.PrevInstance = True Then
    End
    Set Form1 = Nothing
    Unload Me
End If
'open notepad. please check the path of notepad.exe
Call Shell("c:\windows\notepad.exe", vbMaximizedFocus)
End Sub

Private Sub Timer1_Timer()
Static i As Integer, loaded As Boolean, hwnda As Long
'check if notepad is loaded
If hwnda = 0 Then
    hwnda = FindWindow("Notepad", "Untitled - Notepad")
    If hwnda <> 0 Then
        'change the caption
        SetWindowText hwnda, "NILESH"
        loaded = True
        Timer1.Interval = 500
    Else
        Exit Sub
    End If
End If
'keep checking if notepad is still loaded else end the app.
hwnda = FindWindow("Notepad", "NILESH")
If hwnda = 0 And loaded = True Then
        End
        Unload Me
        Set Form1 = Nothing
End If
'change icons
SetClassWord hwnda, GCL_HICON, Picture1(i).Picture.Handle
i = i + 1
If i = 3 Then i = 0
End Sub
