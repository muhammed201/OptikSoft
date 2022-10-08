VERSION 5.00
Begin VB.Form Yardim 
   BackColor       =   &H00EEDCD5&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Yardým"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   9600
      Picture         =   "Yardim.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6600
      Width           =   855
   End
   Begin VB.Timer trmLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9480
      Top             =   6600
   End
   Begin VB.Timer trmUnload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9960
      Top             =   6600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEDCD5&
      Caption         =   "Ayarlar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Index           =   3
      Left            =   240
      TabIndex        =   6
      Top             =   5400
      Width           =   10335
      Begin VB.Label Label1 
         BackColor       =   &H00EEDCD5&
         Caption         =   $"Yardim.frx":046A
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEDCD5&
      Caption         =   "Döviz Kuru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   10335
      Begin VB.Label Label1 
         BackColor       =   &H00EEDCD5&
         Caption         =   $"Yardim.frx":0537
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   9735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEDCD5&
      Caption         =   "Similasyon"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   10335
      Begin VB.Label Label1 
         BackColor       =   &H00EEDCD5&
         Caption         =   $"Yardim.frx":05C7
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   9735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEDCD5&
      Caption         =   "Galeri"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   10335
      Begin VB.Label Label1 
         BackColor       =   &H00EEDCD5&
         Caption         =   $"Yardim.frx":071D
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9735
      End
   End
End
Attribute VB_Name = "Yardim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Local Error GoTo aksehir
t = 0
SetTransparent hwnd, 0
trmLoad.Enabled = True
aksehir:
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnaMenu.Show
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
