VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCancel 
   Caption         =   "取消订单审核"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   6075
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdRefund 
      Caption         =   "确认退款"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "查看详情"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3480
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   353
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "学生姓名"
         Object.Width           =   1524
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "退款金额"
         Object.Width           =   1524
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "发起时间"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "处理情况"
         Object.Width           =   2117
      EndProperty
   End
End
Attribute VB_Name = "frmCancel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public token As String
Public refunds As Object
Public select_order As Object
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Private Sub renderTable(count As Integer, page As Integer, gettype As String)
    ListView1.ListItems.Clear
    Set refunds = Service.getRefunds(token, count, page, gettype)
    
    For i = 1 To refunds.Item("count")
        Dim itm As ListItem
        Set itm = ListView1.ListItems.Add(, , refunds.Item("data").Item(i).Item("id"))
        itm.SubItems(1) = refunds.Item("data").Item(i).Item("order").Item("user").Item("name")
        itm.SubItems(2) = "￥" & refunds.Item("data").Item(i).Item("order").Item("price")
        itm.SubItems(3) = refunds.Item("data").Item(i).Item("created_at")
        itm.SubItems(4) = IIf(refunds.Item("data").Item(i).Item("has_confirmed"), "已处理", "未处理")
        itm.Tag = i
    Next
    
End Sub

Private Sub cmdDetail_Click()
    Load frmOrderDetail
    SetParent frmOrderDetail.hWnd, frmMain.picMain.hWnd
    frmOrderDetail.loadData select_order
End Sub

Private Sub cmdRefund_Click()
    MsgBox ListView1.SelectedItem.Text
End Sub

Private Sub Form_Load()
    renderTable 10, 0, "all"
End Sub

Private Sub Form_Resize()
    With ListView1
        .Width = Me.Width - 540
    End With
End Sub

Private Sub ListView1_DblClick()
    cmdDetail_Click
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.SubItems(4) = "未处理" Then
        cmdRefund.Enabled = True
    Else
        cmdRefund.Enabled = False
    End If
    
    Set select_order = refunds.Item("data").Item(Item.Tag).Item("order")
    cmdDetail.Enabled = True
    
End Sub
