VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCancelDetail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "取消订单详情"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   6376
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
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

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

Public Function loadStudentData(student As Object)
    Me.Caption = "学生详情"
    Set User = student
    Dim nodeindex As Node
    Set nodeindex = TreeView1.Nodes.Add(, , "name", User.Item("name"))
    'nodeindex.Sorted = True
    
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "student_id", "学号")
    Set nodeindex = TreeView1.Nodes.Add("student_id", tvwChild, "id_content", User.Item("student_id"))
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "balance", "余额")
    Set nodeindex = TreeView1.Nodes.Add("balance", tvwChild, "balance_content", User.Item("balance"))
    
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "jaccount", "jAccount")
    Set nodeindex = TreeView1.Nodes.Add("jaccount", tvwChild, "jaccount_", User.Item("jaccount"))
    
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "college", "院系")
    Set nodeindex = TreeView1.Nodes.Add("college", tvwChild, "college_", User.Item("college"))
    
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "id", "身份证号")
    Set nodeindex = TreeView1.Nodes.Add("id", tvwChild, "id_", User.Item("id_number"))

    
End Function

