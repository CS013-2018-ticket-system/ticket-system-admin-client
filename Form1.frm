VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��¼"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CommandButton cmdExit 
      Caption         =   "�˳�"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtAuth 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox txtAuth 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "��¼"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "����"
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�û���"
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLogin_Click()
    Dim authJson As String
    
    If txtAuth(0) = "" Then
        MsgBox "�û���Ϊ���", vbCritical, "����"
    ElseIf txtAuth(1) = "" Then
        MsgBox "����Ϊ���", vbCritical, "����"
    Else
        '���ӵ���������Ȩ
        authJson = Service.makeLoginJson(txtAuth(0), txtAuth(1))
        
        login = Service.postLogin(authJson)
        If login = False Then
            MsgBox "�û������������", vbCritical, "����"
        Else
             '��¼�ɹ�
            Unload frmLogin
            frmMain.Token = login
            frmMain.Show
        End If
    End If
End Sub

Private Sub txtAuth_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdLogin_Click
    End If
End Sub
