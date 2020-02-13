VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dingtalk Boomer v1.0"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8475
   DrawMode        =   8  'Xor Pen
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8475
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÍË³ö"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   21.75
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   7440
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ö´ÐÐ"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   21.75
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   6360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1440
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.TextBox Timerd 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2760
      TabIndex        =   7
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Counter 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2760
      TabIndex        =   6
      Top             =   2400
      Width           =   3135
   End
   Begin VB.TextBox content 
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   12
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   3135
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Developed By Kiramei"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   10.5
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Title 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dingtalk Boom"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ Light"
         Size            =   24
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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

Private Sub Command1_Click()
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

Private Sub Command3_Click()
Unload Me
End Sub

Public Function Main(contentd As String, countd As Integer, timed As Integer)
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

Private Sub Label6_Click()

End Sub
