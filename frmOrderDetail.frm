VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCancelDetail 
   Caption         =   "ȡ����������"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4575
   StartUpPosition =   3  '����ȱʡ
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
    
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "studentinfo", "ѧ����Ϣ")
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "stationinfo", "�г���Ϣ")
    Set nodeindex = TreeView1.Nodes.Add("name", tvwChild, "ticketinfo", "��Ʊ��Ϣ")
    
    Set nodeindex = TreeView1.Nodes.Add("studentinfo", tvwChild, "student_id", "ѧ��")
    Set nodeindex = TreeView1.Nodes.Add("student_id", tvwChild, "id_content", User.Item("student_id"))
    Set nodeindex = TreeView1.Nodes.Add("studentinfo", tvwChild, "balance", "���")
    Set nodeindex = TreeView1.Nodes.Add("balance", tvwChild, "balance_content", User.Item("balance"))
    
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "from", "��ʼվ")
    Set nodeindex = TreeView1.Nodes.Add("from", tvwChild, "from_station", order.Item("from_station"))
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "to", "�յ�վ")
    Set nodeindex = TreeView1.Nodes.Add("to", tvwChild, "to_station", order.Item("to_station"))
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "date", "��������")
    Set nodeindex = TreeView1.Nodes.Add("date", tvwChild, "departure_date", order.Item("departure_date"))
    Set nodeindex = TreeView1.Nodes.Add("stationinfo", tvwChild, "time", "����ʱ��")
    Set nodeindex = TreeView1.Nodes.Add("time", tvwChild, "departure_time", order.Item("departure_time"))
    
    
    Set nodeindex = TreeView1.Nodes.Add("ticketinfo", tvwChild, "type", "��λ����")
    Set nodeindex = TreeView1.Nodes.Add("type", tvwChild, "seat_type", order.Item("seat_type"))
    Set nodeindex = TreeView1.Nodes.Add("ticketinfo", tvwChild, "no", "��λ��")
    Set nodeindex = TreeView1.Nodes.Add("no", tvwChild, "seat_no", order.Item("seat_no"))
    Set nodeindex = TreeView1.Nodes.Add("ticketinfo", tvwChild, "price", "Ʊ��")
    Set nodeindex = TreeView1.Nodes.Add("price", tvwChild, "ticket_price", order.Item("price"))
    
End Function
