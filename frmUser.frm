VERSION 5.00
Begin VB.Form frmUser 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.ListBox List1 
      Height          =   2580
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public token As String

Private Sub Form_Load()
    Set ret_obj = Service.getUsers(token)
    If ret_obj.Item("success") = "True" Then
        user_count = ret_obj.Item("count")
        Set users = ret_obj.Item("data")
        For i = 1 To user_count
            List1.AddItem users.Item(i).Item("name")
        Next
    End If
End Sub

