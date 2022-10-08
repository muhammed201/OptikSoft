Attribute VB_Name = "Module2"

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, _
ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H3
Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED = &H80000


Public Enum eType
  A1 = 1
  Ax = 99
End Enum
Public Type tQ
  nQ As Long
  TypeQ As eType
  nRA As Integer
  nA(100) As Boolean
  nAC As Integer
End Type
Public Q As tQ
Public Enum eCodeType
  ctNone = 0
  ctASCII10 = 1
  ctASCII16 = 2
End Enum
Public Type tTest
  sTheme As String
  sAutor As String
  nNumOfQuestions As Long
  CodeType As eCodeType
  sPath As String
  DontHelp As Boolean
End Type
Public test As tTest
Public bTestLoaded As Boolean
Public LastPath As String
Public nCorrectly As Long
Public HelpEn As Boolean
Public Win2k As Boolean
Public bDontHelp As Boolean

Public Function ClearMemory()
  test.sPath = ""
  test.CodeType = ctNone
  test.nNumOfQuestions = 0
  test.sTheme = "None"
  test.sAutor = ""
  nCorrectly = 0
  frmBase.Org.Text = ""
  frmBase.Dbl.Text = ""
  bTestLoaded = False
End Function

Public Function LoadBase(sBasePath As String) As Integer
  Dim ssTempEEE8B43A4F4B As String

  If Dir$(sBasePath) = "" Then
    LoadBase = 3
    Exit Function
  End If
  Open sBasePath For Input As #1

  Line Input #1, ssTempEEE8B43A4F4B
  If Not ssTempEEE8B43A4F4B = "BaseDate For Test v.1.0" Then GoTo FormatError
  Line Input #1, ssTempEEE8B43A4F4B
  If Not ssTempEEE8B43A4F4B = "#Include TestRWModule.Read" Then GoTo FormatError
  Line Input #1, ssTempEEE8B43A4F4B
  If Not ssTempEEE8B43A4F4B = "Begin" Then GoTo FormatError
  While Not EOF(1)
    Line Input #1, ssTempEEE8B43A4F4B
    If EOF(1) = True Then If ssTempEEE8B43A4F4B <> "End." Then GoTo FormatError
  Wend
  Close #1
  
  bTestLoaded = True
  LoadBase = -1
  Exit Function
  
FormatError:
  Close #1
  LoadBase = 1
End Function



Public Function SetTransparent(hwnd As Long, Layered As Byte) As Boolean
On Error GoTo 1
Dim Ret As Long

Ret = GetWindowLong(hwnd, GWL_EXSTYLE)

Ret = Ret Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, Ret

SetLayeredWindowAttributes hwnd, 1, Layered, LWA_ALPHA
SetTransparent = True
1 Exit Function
End Function


