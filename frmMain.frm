VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "��Ʊϵͳ �������"
   ClientHeight    =   3870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   3870
   ScaleWidth      =   6795
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Token As String

Private Sub Form_Load()
    MsgBox Token
End Sub
