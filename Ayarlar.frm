VERSION 5.00
Begin VB.Form Ayarlar 
   BackColor       =   &H80000013&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayarlar"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000013&
      Height          =   735
      Left            =   7680
      Picture         =   "Ayarlar.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H009AEDA7&
      Caption         =   "Þifreyi Aktif\Pasif Yap"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EAE7C8&
      Caption         =   "Programcý"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   5
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00D67A84&
      Caption         =   "Yeni Kullanýcý Ekle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00EFA7EA&
      Caption         =   "Kullanýcý Bilgileri Düzenle/Sil"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "Ayarlar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
If Index = 1 Then
On Local Error GoTo aksehir
Dim kullanici, sifre As String
kullanici = InputBox("Kullanýcý adaný giriniz", "Kullanýcý")
sifre = InputBox("Güvenlik þifresini giriniz", "Þifre")
If kullanici = Empty Or sifre = Empty Then
MsgBox "Kullanýcý adýný yada þifresini boþ býrakamazsýnýz!", vbCritical
Exit Sub
End If
Call veri_ac(False, False)
Call tablo_ac("select * from guvenlik")
tablo.AddNew
tablo!kul = kullanici
tablo!sifre = sifre
tablo!admin = True
tablo.Update
ElseIf Index = 0 Then
Me.Hide
kullanicifrm.Show
ElseIf Index = 2 Then
Call aktiff
ElseIf Index = 5 Then
Me.Hide
programci.Show
End If
aksehir:
If Err = 3022 Then
MsgBox "Bu kullanýcý ismini kullanamazsýnýz" & vbCrLf & "Bu isim veri tabanýnda zaten kayýtlý!", vbCritical, "HATA!!!"
Exit Sub
End If



End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
AnaMenu.Show
End Sub

Sub aktiff()
On Local Error GoTo hata
Dim numara As Byte
Call veri_ac(False, False)
Call tablo_ac("Select * from izin")
numara = MsgBox("Programýn þifre girilmeden açýlmasýný istiyormusunuz?" & vbCrLf & "Eðer þifre girmeden açmak istiyorsanýz EVET'i aksi halde HAYIR'ý týklayýnýz", vbYesNo, "Þifre Ýzni")
tablo.Edit
If numara = 6 Then
tablo!aktif = True
Else
tablo!aktif = False
End If
tablo.Update
tablo.Close
veri.Close

hata:
If Err = 3021 Then
Call veri_ac(False, False)
Call tablo_ac("Select * from izin")
tablo.AddNew
tablo!aktif = True
tablo.Update
tablo.Close
veri.Close
End If

End Sub
