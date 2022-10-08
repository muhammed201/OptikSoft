VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RSEFileRead.ocx"
Begin VB.Form Hatýrlatma2 
   Caption         =   "Form2"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10440
   LinkTopic       =   "Form2"
   ScaleHeight     =   8070
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "------------>"
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   7080
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<--------------"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   7080
      Width           =   1935
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   11033
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Hatýrlatma2.frx":0000
   End
End
Attribute VB_Name = "Hatýrlatma2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
