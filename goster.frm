VERSION 5.00
Object = "{D0CD41D3-0FA6-4C72-8799-854FB0C7CC6F}#1.0#0"; "AResizeLite.ocx"
Begin VB.Form goster 
   BorderStyle     =   0  'None
   Caption         =   "Ön Ýzleme"
   ClientHeight    =   9750
   ClientLeft      =   1680
   ClientTop       =   555
   ClientWidth     =   11490
   LinkTopic       =   "Form2"
   Picture         =   "goster.frx":0000
   ScaleHeight     =   9750
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer trmUnload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer trmLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   480
   End
   Begin ActiveResizeLiteCtl.ActiveResizeLite ActiveResizeLite1 
      Left            =   11880
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   2
      ScreenHeight    =   768
      ScreenWidth     =   1024
      ScreenHeightDT  =   768
      ScreenWidthDT   =   1024
      FormHeightDT    =   9750
      FormWidthDT     =   11490
      FormScaleHeightDT=   9750
      FormScaleWidthDT=   11490
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "Enter tuþu=AnaMenü                                              Space(Boþluk) tuþu=Büyült/Küçült"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   9480
      Width           =   10335
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000013&
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   4455
      Left            =   5880
      Top             =   4920
      Width           =   5175
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000013&
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      Height          =   4455
      Left            =   360
      Top             =   4920
      Width           =   5175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000013&
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      FillColor       =   &H00FF0000&
      Height          =   4455
      Left            =   5880
      Top             =   240
      Width           =   5175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      FillColor       =   &H000000FF&
      Height          =   4455
      Left            =   360
      Top             =   240
      Width           =   5175
   End
   Begin VB.Image resim1 
      Height          =   3975
      Left            =   600
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4695
   End
   Begin VB.Image resim2 
      Height          =   3975
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   480
      Width           =   4695
   End
   Begin VB.Image resim3 
      Height          =   3975
      Left            =   600
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   4695
   End
   Begin VB.Image resim4 
      Height          =   3975
      Left            =   6120
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   4695
   End
End
Attribute VB_Name = "goster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
trmUnload.Enabled = True
ElseIf KeyCode = 32 Then
If Me.WindowState = 0 Then
Me.WindowState = 2
Else
Me.WindowState = 0
End If
End If
End Sub

Private Sub Form_Load()
On Local Error GoTo kamerberra
Unload AnaMenu
t = 0
SetTransparent hwnd, 0
trmLoad.Enabled = True
resim1.Picture = LoadPicture(App.Path & "/similasyon(1).jpg")
resim2.Picture = LoadPicture(App.Path & "/similasyon(2).jpg")
resim3.Picture = LoadPicture(App.Path & "/similasyon(3).jpg")
resim4.Picture = LoadPicture(App.Path & "/similasyon(4).jpg")
kamerberra:
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnaMenu.Show
End Sub

Private Sub trmLoad_Timer()
SetTransparent hwnd, t
  If t >= 240 Then
    trmLoad.Enabled = False
        Else
    t = t + 10
  End If
End Sub

Private Sub trmUnload_Timer()
SetTransparent hwnd, t
  If trmLoad.Enabled = True Then trmLoad.Enabled = False
  If t <= 0 Then
    trmUnload.Enabled = False
    Unload Me
        Else
    t = t - 10
  End If
  End Sub


