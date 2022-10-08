VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form kullanicifrm 
   BackColor       =   &H80000013&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kullanýcýlar"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   7800
      Picture         =   "kullanici.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000013&
      Caption         =   "Güncelleme"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   4320
      TabIndex        =   1
      Top             =   1080
      Width           =   4335
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF00FF&
         Caption         =   "Kaydý Sil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2160
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   960
         TabIndex        =   4
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   960
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Güncelle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2160
         Width           =   1575
      End
   End
   Begin MSFlexGridLib.MSFlexGrid liste 
      Height          =   3735
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   6588
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   16711680
      BackColorFixed  =   255
      BackColorBkg    =   65535
   End
End
Attribute VB_Name = "kullanicifrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Local Error GoTo hata
If Text1 = "admin" Then
MsgBox "Bu kaydý deðiþtiremezsiniz yada silemezsiniz!", vbExclamation
Exit Sub
End If
If Text1 = Empty Or Text2 = Empty Then
MsgBox "Boþ alan býrakamazsýnýz", vbCritical
Exit Sub
End If
Call veri_ac(False, False)
Call tablo_ac("select * from guvenlik where kul='" & liste.TextMatrix(liste.Row, 0) & "'")
tablo.Edit
If tablo!kul = "admin" Then
MsgBox "Bu kaydý deðiþtiremezsiniz"
Exit Sub
End If
tablo!kul = Text1
tablo!sifre = Text2
tablo.Update
MsgBox "Kullanýcý Deðiþtirilmiþtir"
Call yukle
hata:
If Err = 3022 Then
MsgBox "Bu isim kullanýlmaktadýr", vbCritical
Exit Sub
End If
End Sub

Private Sub Command2_Click()
On Local Error GoTo hata
If Text1 = "admin" Then
MsgBox "Bu kaydý deðiþtiremezsiniz yada silemezsiniz!", vbExclamation
Exit Sub
End If

Call veri_ac(False, False)
Call tablo_ac("Select * from guvenlik where kul='" & Text1 & "'")
tablo.Delete
MsgBox "Kayýt silinmiþtir", vbApplicationModal
Text1 = Empty
Text2 = Empty
liste.Clear
Call yukle
hata:
If Err = 3021 Then
MsgBox "Böyle bir kayýt yok", vbCritical
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call yukle
End Sub

Private Sub Form_Unload(Cancel As Integer)
Ayarlar.Show
End Sub

Private Sub liste_Click()
Text1 = liste.TextMatrix(liste.Row, 0)
Text2 = liste.TextMatrix(liste.Row, 1)
If Text1.Text = Empty Or Text2.Text = Empty Then
Frame1.Enabled = False
Else
Frame1.Enabled = True
End If
End Sub
Sub yukle()
Dim i As Byte
If Text1.Text = Empty Or Text2.Text = Empty Then
Frame1.Enabled = False
End If
i = 0
liste.Rows = 2
liste.Cols = 3
liste.TextMatrix(0, 0) = "Kullanýcý"
liste.TextMatrix(0, 1) = "Þifre"
Call veri_ac(False, False)
Call tablo_ac("select * from guvenlik")
Do While Not tablo.EOF
liste.AddItem ""
i = i + 1
liste.TextMatrix(i, 0) = tablo!kul
liste.TextMatrix(i, 1) = tablo!sifre
tablo.MoveNext
Loop
tablo.Close
veri.Close
End Sub
