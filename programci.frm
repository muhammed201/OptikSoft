VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form programci 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3465
   ClientLeft      =   5025
   ClientTop       =   2805
   ClientWidth     =   4095
   LinkTopic       =   "Form2"
   ScaleHeight     =   3465
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000014&
      Caption         =   "Programcý"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
      Begin WMPLibCtl.WindowsMediaPlayer cal 
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   3375
         URL             =   "g"
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   5953
         _cy             =   873
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Muhammed Zengin"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
         Caption         =   "Akþehir-2007"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   18
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "muhammed201@hotmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "05432223973"
         BeginProperty Font 
            Name            =   "Monotype Corsiva"
            Size            =   15.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   3015
      End
   End
End
Attribute VB_Name = "programci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
Ayarlar.Show
End Sub

Private Sub Form_Load()
cal.URL = App.Path & "/hosgeldin.wav"
End Sub

Private Sub Frame1_Click()
Unload Me
Ayarlar.Show
End Sub

Private Sub Label1_Click()
Unload Me
Ayarlar.Show
End Sub

Private Sub Label2_Click()
Unload Me
Ayarlar.Show
End Sub

Private Sub Label3_Click()
Unload Me
Ayarlar.Show
End Sub

Private Sub Label4_Click()
Unload Me
Ayarlar.Show
End Sub
