VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户管理"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.ListBox List1 
      Height          =   3300
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label Label1 
      Caption         =   "双击查看详情"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public token As String
Dim users As Object
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub Form_Load()
    Set ret_obj = Service.getUsers(token)
    If ret_obj.Item("success") = "True" Then
        user_count = ret_obj.Item("count")
        Set users = ret_obj.Item("data")
        For I = 1 To user_count
            List1.AddItem users.Item(I).Item("name")
        Next
    End If
    
End Sub

Private Function getSelected()
    For I = 0 To List1.ListCount - 1
        If List1.Selected(I) = True Then
            getSelected = I
        End If
    Next
End Function

Private Sub List1_Click()
    frmCancelDetail.token = Me.token
    Load frmCancelDetail
    SetParent frmCancelDetail.hwnd, frmMain.picMain.hwnd
    frmCancelDetail.loadStudentData users.Item(getSelected() + 1)
    frmCancelDetail.Show
    
End Sub
