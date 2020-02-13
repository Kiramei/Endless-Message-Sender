VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Dingtalk Boomer v1.0"
   ClientHeight    =   5160
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   DrawMode        =   8  'Xor Pen
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Timerd 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   15
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2280
      TabIndex        =   7
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox Counter 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   15
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2280
      TabIndex        =   6
      Top             =   2400
      Width           =   3615
   End
   Begin VB.TextBox content 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   15
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "   ÍË³ö"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   21.75
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   7440
      TabIndex        =   12
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "   Ö´ÐÐ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   21.75
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   6360
      TabIndex        =   11
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Developed By Kiramei"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   10.5
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1440
      TabIndex        =   10
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Ö´ÐÐÖ®ºó»áÓÐ3ÃëµÄµÈ´ýÊ±¼ä£¬Çëµã»÷¶¤¶¤¶Ô»°¿ò"
      BeginProperty Font 
         Name            =   "¿¬Ìå"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "ÊÊÁ¿¿ØÖÆ£¬ÒÔÃâ»ú×Ó¿¨±¬£¡"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   4320
      Width           =   7575
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFF00&
      Caption         =   "·¢ËÍ´ÎÊý£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFF00&
      Caption         =   "(µ¥Î»£ººÁÃë)"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   9
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF00&
      Caption         =   "¼ä¸ôÊ±¼ä£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      Caption         =   "Ö¸¶¨ÄÚÈÝ£º"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   14.25
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFF00&
      Caption         =   "Dingtalk Boomer"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   24
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Dim xa As Single, ya As Single

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xa = X
ya = Y
End Sub

Private Sub Form_Load()
Me.BackColor = RGB(17, 207, 255)
Counter.BackColor = RGB(34, 170, 170)
Title.BackColor = RGB(17, 207, 255)
Label1.BackColor = RGB(17, 207, 255)
Label2.BackColor = RGB(17, 207, 255)
Label3.BackColor = RGB(17, 207, 255)
Label4.BackColor = RGB(17, 207, 255)
Label5.BackColor = RGB(17, 207, 255)
Label6.BackColor = RGB(17, 207, 255)
Label7.BackColor = RGB(17, 207, 255)
Label8.BackColor = RGB(8, 231, 231)
Label9.BackColor = RGB(8, 231, 231)
Timerd.BackColor = RGB(34, 170, 170)
content.BackColor = RGB(34, 170, 170)
Dim rtn As Long
Dim BorderStyler
BorderStyler = 0
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, RGB(17, 207, 255), 200, LWA_ALPHA
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = RGB(8, 231, 231)
Label9.BackColor = RGB(8, 231, 231)
If Button = 1 Then Me.Move Me.Left + X - xa, Me.Top + Y - ya
End Sub
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = RGB(43, 43, 213)
End Sub
Private Sub Label9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label9.BackColor = RGB(43, 43, 213)
End Sub

Private Function Main(contentd As String, countd As Integer, timed As Integer)
Set objShell = CreateObject("Wscript.Shell")

Dim a As Integer: a = 0
Clipboard.Clear
Clipboard.SetText contentd

Sleep 3000

Do
SendKeys "^(v)"
SendKeys "{ENTER}"
a = a + 1
Sleep timed
Loop Until a = countd

End Function


Private Sub Label8_Click()
Dim defcount As Integer: defcount = 3
Dim deftime As Integer: deftime = 1000

Dim contents As String
contents = content.Text
Dim counts As Integer
counts = Val(Counter.Text)
Dim times As Integer
times = Val(Timerd.Text)


If Not contents <> "" Then
MsgBox "ÇëÊäÈëÖ¸¶¨ÄÚÈÝ£¡", vbDefaultButton1, "´íÎó"
Else
If counts = 0 Or counts < 0 Then
counts = defcount
End If
If times = 0 Or times < 0 Then
times = deftime
End If
Main contents, counts, times
End If
End Sub

Private Sub Label9_Click()
Unload Me
End Sub



