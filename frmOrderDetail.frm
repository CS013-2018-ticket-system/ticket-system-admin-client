VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCancelDetail 
   Caption         =   "取消订单详情"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4575
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4683
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmCancelDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public token As String

Public Function loadData(order As Object)
    Set User = order.Item("user")
    Dim nodeindex As Node
    Set nodeindex = TreeView1.Nodes.Add(, , "name", User.Item("name"))
    'nodeindex.Sorted = True
    
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "studentinfo", "学生信息")
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "stationinfo", "列车信息")
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "ticketinfo", "车票信息")
    
    Set nodeindex = TreeView1.Nodes.Add("studentinfo", tvwChild, "student_id", "学号")
    Set nodeindex = TreeView1.Nodes.Add("student_id", tvwChild, "id_content", User.Item("student_id"))
    Set nodeindex = TreeView1.Nodes.Add("studentinfo", tvwChild, "balance", "余额")
    Set nodeindex = TreeView1.Nodes.Add("balance", tvwChild, "balance_content", User.Item("balance"))
    
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "from", "起始站")
    Set nodeindex = TreeView1.Nodes.Add("from", tvwChild, "from_station", order.Item("from_station"))
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "to", "终点站")
    Set nodeindex = TreeView1.Nodes.Add("to", tvwChild, "to_station", order.Item("to_station"))
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "date", "发车日期")
    Set nodeindex = TreeView1.Nodes.Add("date", tvwChild, "departure_date", order.Item("departure_date"))
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "time", "发车时间")
    Set nodeindex = TreeView1.Nodes.Add("time", tvwChild, "departure_time", order.Item("departure_time"))
    
    
    Set nodeindex = TreeView1.Nodes.Add("ticketinfo", tvwChild, "type", "座位类型")
    Set nodeindex = TreeView1.Nodes.Add("type", tvwChild, "seat_type", order.Item("seat_type"))
    Set nodeindex = TreeView1.Nodes.Add("ticketinfo", tvwChild, "no", "座位号")
    Set nodeindex = TreeView1.Nodes.Add("no", tvwChild, "seat_no", order.Item("seat_no"))
    Set nodeindex = TreeView1.Nodes.Add("ticketinfo", tvwChild, "price", "票价")
    Set nodeindex = TreeView1.Nodes.Add("price", tvwChild, "ticket_price", order.Item("price"))
    
End Function
