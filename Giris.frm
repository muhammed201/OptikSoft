VERSION 5.00
Begin VB.Form Giris 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Optik Soft 1.0  programcý: muhammed201@hotmail.com"
   ClientHeight    =   3105
   ClientLeft      =   5295
   ClientTop       =   4470
   ClientWidth     =   6090
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Giris.frx":0000
   ScaleHeight     =   3105
   ScaleWidth      =   6090
   Begin VB.Timer trmUnload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   360
   End
   Begin VB.Timer trmLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Çýkýþ"
      Height          =   360
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   2880
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Giriþ"
      Height          =   360
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "Giris"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    
Option Explicit

Rem Formu köþeli yapan komutlar burdan baþlýyor
Dim Rgn1 As Long
Dim a, b

Private Sub Form_Load()
On Local Error GoTo muhammed
a = Me.Height / 10
b = Me.Width / 10
Rgn1 = CreateRoundRectRgn(0, 0, a, b, 30, 30)
SetWindowRgn hwnd, Rgn1, True
SetTransparent hwnd, 0
trmLoad.Enabled = True
muhammed:
Resume Next
End Sub
Rem Formu köþeli yapan komutlar burda bitiyor.
  

Private Sub Command1_Click()
Dim onay As Boolean
Call veri_ac(False, False)
Call tablo_ac("select * from guvenlik where kul='" & Text1(0) & "' and sifre='" & Text1(1) & "'")
If tablo.EOF = True Then
MsgBox "Yanlýþ Kullanýcý Adý yada Parola", vbCritical
Exit Sub
End If
trmUnload.Enabled = True
AnaMenu.Show
tablo.Close
veri.Close
End Sub

Private Sub Command2_Click()
trmUnload.Enabled = True
End Sub





Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Dim i As Byte
For i = 0 To 1
If KeyAscii = 13 Then
Text1(i).SetFocus
End If
Next i
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



