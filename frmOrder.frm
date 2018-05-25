VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOrder 
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   10095
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5295
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   9255
      Begin MSComctlLib.ListView ListView1 
         Height          =   4215
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   7435
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public token As String

Private Sub Form_Load()
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    ListView1.GridLines = True
    ListView1.LabelEdit = lvwManual
    ListView1.FullRowSelect = True
    ListView1.ColumnHeaders.Add , , "#", 0
    ListView1.ColumnHeaders.Add , , "订单流水号", 1300
    ListView1.ColumnHeaders.Add , , "用户ID", 1000
    ListView1.ColumnHeaders.Add , , "车次号", 1000
    ListView1.ColumnHeaders.Add , , "用户出发站", 1300
    ListView1.ColumnHeaders.Add , , "用户到达站", 1300
    ListView1.ColumnHeaders.Add , , "出发日期", 1000
    ListView1.ColumnHeaders.Add , , "出发时间", 1000
    ListView1.ColumnHeaders.Add , , "座位类型", 1000
    ListView1.ColumnHeaders.Add , , "座位号", 800
    ListView1.ColumnHeaders.Add , , "总价", 600
    ListView1.ColumnHeaders.Add , , "是否已支付", 1300
    ListView1.ColumnHeaders.Add , , "是否取消", 1000
    '-------------------------------------------------------
    ListView1.View = lvwReport
    ListView1.AllowColumnReorder = True
    ListView1.Arrange = lvwAutoLeft
    ListView1.Arrange = lvwAutoTop
    ListView1.FlatScrollBar = False
    ListView1.FlatScrollBar = True
    ListView1.FullRowSelect = True
    ListView1.LabelEdit = lvwManual
    ListView1.GridLines = True
    ListView1.LabelWrap = True
    ListView1.MultiSelect = True
    ListView1.PictureAlignment = lvwTopLeft
    ListView1.Checkboxes = True

  
    Set ret_obj = Service.getOrders(token)
    If ret_obj.Item("success") = "True" Then
        ccount = ret_obj.Item("count")
        Print ccount
        Set orders = ret_obj.Item("data")

        For i = 1 To ccount
Dim itm As ListItem
Set itm = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)  '序号
itm.SubItems(1) = orders.Item(i).Item("id")
itm.SubItems(2) = orders.Item(i).Item("user_id")
itm.SubItems(3) = orders.Item(i).Item("train_code")
itm.SubItems(4) = orders.Item(i).Item("from_station")
itm.SubItems(5) = orders.Item(i).Item("to_station")
itm.SubItems(6) = orders.Item(i).Item("departure_date")
itm.SubItems(7) = orders.Item(i).Item("departure_time")
itm.SubItems(8) = orders.Item(i).Item("seat_type")
If orders.Item(i).Item("seat_no") <> Null Then
itm.SubItems(9) = orders.Item(i).Item("seat_no")
Else
itm.SubItems(9) = "/"
End If

itm.SubItems(10) = orders.Item(i).Item("price")
If orders.Item(i).Item("has_paid") = 1 Then
itm.SubItems(11) = "已支付"
Else
itm.SubItems(11) = "未支付"
End If
If orders.Item(i).Item("has_cancelled") = 1 Then
itm.SubItems(12) = "已取消"
Else
itm.SubItems(12) = "未取消"
End If
        Next i
    End If
    
End Sub

