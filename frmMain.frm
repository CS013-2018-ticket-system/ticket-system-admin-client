VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "订票系统 管理面板"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   7800
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   0
      ScaleHeight     =   3705
      ScaleWidth      =   7785
      TabIndex        =   1
      Top             =   720
      Width           =   7815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   600
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   36
      ImageHeight     =   36
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   1270
      ButtonWidth     =   1138
      ButtonHeight    =   1111
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "用户管理"
            Object.ToolTipText     =   "用户管理"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "所有订单"
            Object.ToolTipText     =   "所有订单"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "订单取消审核"
            Object.ToolTipText     =   "审核取消订单"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public token As String

Private Sub Form_Resize()
    With picMain
        .Top = Toolbar1.Height
        .Left = 0
        .Height = Me.Height - Toolbar1.Height
        .Width = Me.Width
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.index
        Case 1
            frmUser.token = Me.token
            Load frmUser
            SetParent frmUser.hWnd, picMain.hWnd
            frmUser.Show
        Case 2
            frmOrder.token = Me.token
            Load frmOrder
            SetParent frmOrder.hWnd, picMain.hWnd
            frmOrder.Show
        Case 3
            frmCancel.token = Me.token
            Load frmCancel
            SetParent frmCancel.hWnd, picMain.hWnd
            frmCancel.Show
            
    End Select
End Sub
