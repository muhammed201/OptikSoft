VERSION 5.00
Begin VB.Form AnaMenu 
   Caption         =   "Optik-Soft 1.0 "
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   Picture         =   "AnaMenu.frx":0000
   ScaleHeight     =   3675
   ScaleWidth      =   7395
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer trmLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer trmUnload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   480
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "AnaMenu.frx":36D7
      Height          =   495
      Index           =   5
      Left            =   5040
      Picture         =   "AnaMenu.frx":3CB8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "AnaMenu.frx":4480
      Height          =   495
      Index           =   4
      Left            =   2760
      Picture         =   "AnaMenu.frx":4A66
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "AnaMenu.frx":52BB
      Height          =   495
      Index           =   3
      Left            =   480
      Picture         =   "AnaMenu.frx":5838
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "AnaMenu.frx":6013
      Height          =   495
      Index           =   2
      Left            =   5040
      Picture         =   "AnaMenu.frx":65CB
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "AnaMenu.frx":6E87
      Height          =   495
      Index           =   1
      Left            =   2760
      Picture         =   "AnaMenu.frx":7464
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      DownPicture     =   "AnaMenu.frx":7BC8
      Height          =   495
      Index           =   0
      Left            =   480
      Picture         =   "AnaMenu.frx":818C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "AnaMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
galeri.Show
trmUnload = True
ElseIf Index = 1 Then
Similasyon.Show
trmUnload = True
ElseIf Index = 2 Then
KURLAR.Show
trmUnload = True
ElseIf Index = 3 Then
Ayarlar.Show
trmUnload = True
ElseIf Index = 4 Then
Yardim.Show
trmUnload = True
ElseIf Index = 5 Then
trmUnload = True
End If
End Sub

Private Sub Form_Load()
t = 0
SetTransparent hwnd, 0
trmLoad.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
trmUnload = True
trmLoad = False

End Sub

Private Sub trmLoad_Timer()
SetTransparent hwnd, t
  If t >= 228 Then
    trmLoad.Enabled = False
        Else
    t = t + 4
  End If
End Sub

Private Sub trmUnload_Timer()
SetTransparent hwnd, t
  If trmLoad.Enabled = True Then trmLoad.Enabled = False
  If t <= 0 Then
    trmUnload.Enabled = False
    Unload Me
        Else
    t = t - 4
  End If
End Sub

