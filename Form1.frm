VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "登录"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox txtAuth 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtAuth 
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "登录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4320
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "请联系 021-54749110"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "广告位招租"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "密码"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "用户名"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLogin_Click()
    Dim authJson As String
    
    If txtAuth(0) = "" Then
        MsgBox "用户名为必填！", vbCritical, "错误"
    ElseIf txtAuth(1) = "" Then
        MsgBox "密码为必填！", vbCritical, "错误"
    Else
        '连接到服务器鉴权
        authJson = Service.makeLoginJson(txtAuth(0), txtAuth(1))
        
        login = Service.postLogin(authJson)
        If login = False Then
            MsgBox "用户名或密码错误。", vbCritical, "错误"
        Else
             '登录成功
            Unload frmLogin
            frmMain.token = login
            frmMain.Show
        End If
    End If
End Sub

Private Sub Form_Load()
    mskinner.Attach Me.hwnd
End Sub

Private Sub Form_Paint()
    txtAuth(0).SetFocus
End Sub

Private Sub txtAuth_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdLogin_Click
    End If
End Sub
