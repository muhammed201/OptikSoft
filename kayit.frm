VERSION 5.00
Begin VB.Form kayit 
   Caption         =   "Galeri"
   ClientHeight    =   7410
   ClientLeft      =   2010
   ClientTop       =   1980
   ClientWidth     =   11025
   ControlBox      =   0   'False
   DrawMode        =   2  'Blackness
   LinkTopic       =   "Form1"
   Picture         =   "kayit.frx":0000
   ScaleHeight     =   7410
   ScaleWidth      =   11025
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   9720
      Picture         =   "kayit.frx":2DD3A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6120
      Width           =   855
   End
   Begin VB.Timer anim2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8880
      Top             =   6840
   End
   Begin VB.Timer anim 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8400
      Top             =   6840
   End
   Begin VB.CommandButton ileri 
      DownPicture     =   "kayit.frx":2E1A4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Picture         =   "kayit.frx":2E60F
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton geri 
      DownPicture     =   "kayit.frx":2EBE0
      Height          =   615
      Left            =   7440
      Picture         =   "kayit.frx":2F028
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000016&
      Caption         =   "Ana Menü"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000016&
      Caption         =   "Browserý Gizle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6240
      UseMaskColor    =   -1  'True
      Width           =   2055
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   2430
      Left            =   3360
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   2340
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000013&
      Caption         =   "Resim"
      Height          =   4215
      Left            =   6000
      TabIndex        =   0
      Top             =   720
      Width           =   4575
      Begin VB.CommandButton Command5 
         Caption         =   "Büyült"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3840
         TabIndex        =   10
         Top             =   120
         Width           =   570
      End
      Begin VB.Image Image1 
         Height          =   3855
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Timer trmUnload 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10080
      Top             =   6840
   End
   Begin VB.Timer trmLoad 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9600
      Top             =   6840
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000013&
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      DrawMode        =   10  'Mask Pen
      FillStyle       =   7  'Diagonal Cross
      Height          =   975
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   6000
      Width           =   10335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resim Formatýnda ki dosyalar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000013&
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      DrawMode        =   10  'Mask Pen
      FillStyle       =   7  'Diagonal Cross
      Height          =   5295
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   10695
   End
End
Attribute VB_Name = "kayit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub anim_Timer()

If Frame2.Left <= 3120 Then
anim.Enabled = False
Else
Frame2.Left = Frame2.Left - 100
ileri.Left = ileri.Left - 100
geri.Left = geri.Left - 100
End If

End Sub

Private Sub anim2_Timer()
If Frame2.Left >= 5880 Then
anim2.Enabled = False
Drive1.Visible = True
File1.Visible = True
Dir1.Visible = True
Label1.Visible = True
Else
Frame2.Left = Frame2.Left + 100
ileri.Left = ileri.Left + 100
geri.Left = geri.Left + 100
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Browserý Gizle" Then
Drive1.Visible = False
File1.Visible = False
Dir1.Visible = False
Label1.Visible = False
anim2.Enabled = False
anim.Enabled = True
Command1.Caption = "Browserý Göster"
Else
anim.Enabled = False
anim2.Enabled = True
Command1.Caption = "Browserý Gizle"
End If
End Sub

Private Sub Command2_Click()
trmUnload.Enabled = True
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command5_Click()
buyuk.Show
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
File1.Pattern = "*.jpg"
File1.Pattern = "*.bmp"
File1.Pattern = "*.BMP"
File1.Pattern = "*.JPG"
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Local Error GoTo hata
Image1.Picture = LoadPicture(File1.Path & "\" & File1)
badres = File1.Path & "\" & File1
Label2.Caption = File1.ListIndex + 1
hata:
Resume Next
End Sub

Private Sub Form_Load()
On Local Error GoTo hata
t = 0
SetTransparent hwnd, 0
trmLoad.Enabled = True

Call veri_ac(False, False)
Call tablo_ac("Select * from adres")
Dir1 = tablo!yol1
Drive1.Drive = tablo!yol2
tablo.Close
veri.Close
File1.Pattern = "*.jpg"
File1.Pattern = "*.bmp"
File1.Pattern = "*.BMP"
File1.Pattern = "*.JPG"
hata:
If Err = 76 Then
Dir1 = ""
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error GoTo hata
Call veri_ac(False, False)
Call tablo_ac("Select * from adres")
If tablo.RecordCount <= 0 Then
tablo.AddNew
tablo!yol = File1.Path
tablo!yol2 = Drive1.Drive
tablo.Update
tablo.Close
veri.Close
Else
tablo.Edit
tablo!yol1 = File1.Path
tablo!yol2 = Drive1.Drive
tablo.Update
tablo.Close
veri.Close
End If
AnaMenu.Show
badres = ""
hata:
Resume Next
End Sub

Private Sub geri_Click()
If File1.ListIndex <= 0 Then
Exit Sub
Else
File1.ListIndex = File1.ListIndex - 1
End If
Label2.Caption = File1.ListIndex + 1
End Sub

Private Sub ileri_Click()
If File1.ListIndex >= File1.ListCount - 1 Then
Exit Sub
Else
File1.ListIndex = File1.ListIndex + 1
End If
Label2.Caption = File1.ListIndex + 1
End Sub

Private Sub trmLoad_Timer()
SetTransparent hwnd, t
  If t >= 235 Then
    trmLoad.Enabled = False
        Else
    t = t + 8
  End If
End Sub

Private Sub trmUnload_Timer()
SetTransparent hwnd, t
  If trmLoad.Enabled = True Then trmLoad.Enabled = False
  If t <= 0 Then
    trmUnload.Enabled = False
    Unload Me
        Else
    t = t - 8
  End If
  End Sub
