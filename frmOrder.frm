VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOrder 
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17850
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   17850
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   7335
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   14055
      Begin VB.CommandButton Command4 
         Caption         =   "�����ڲ�ѯ"
         Height          =   615
         Left            =   6360
         TabIndex        =   5
         Top             =   6480
         Width           =   2175
      End
      Begin VB.CommandButton Command3 
         Caption         =   "δ֧������"
         Height          =   615
         Left            =   4320
         TabIndex        =   4
         Top             =   6480
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��֧������"
         Height          =   615
         Left            =   2400
         TabIndex        =   3
         Top             =   6480
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����ˮ�Ų��Ҷ���"
         Height          =   615
         Left            =   720
         TabIndex        =   2
         Top             =   6480
         Width           =   1695
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5535
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   9763
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

Private Sub Command1_Click()
ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    ListView1.GridLines = True
    ListView1.LabelEdit = lvwManual
    ListView1.FullRowSelect = True
    ListView1.ColumnHeaders.Add , , "#", 0
    ListView1.ColumnHeaders.Add , , "������ˮ��", 1300
    ListView1.ColumnHeaders.Add , , "�û�ID", 1000
    ListView1.ColumnHeaders.Add , , "���κ�", 1000
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "��������", 1000
    ListView1.ColumnHeaders.Add , , "����ʱ��", 1000
    ListView1.ColumnHeaders.Add , , "��λ����", 1000
    ListView1.ColumnHeaders.Add , , "��λ��", 800
    ListView1.ColumnHeaders.Add , , "�ܼ�", 600
    ListView1.ColumnHeaders.Add , , "�Ƿ���֧��", 1300
    ListView1.ColumnHeaders.Add , , "�Ƿ�ȡ��", 1000
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
    Dim t As Integer
    t = InputBox("�����ҵ���ˮ��")
        For i = 1 To ccount
If Val(orders.Item(i).Item("id")) = t Then
Dim itm As ListItem
Set itm = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)  '���
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
itm.SubItems(11) = "��֧��"
Else
itm.SubItems(11) = "δ֧��"
End If
If orders.Item(i).Item("has_cancelled") = 1 Then
itm.SubItems(12) = "��ȡ��"
Else
itm.SubItems(12) = "δȡ��"
End If
End If
        Next i
    End If
    


End Sub

Private Sub Command2_Click()
ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    ListView1.GridLines = True
    ListView1.LabelEdit = lvwManual
    ListView1.FullRowSelect = True
    ListView1.ColumnHeaders.Add , , "#", 0
    ListView1.ColumnHeaders.Add , , "������ˮ��", 1300
    ListView1.ColumnHeaders.Add , , "�û�ID", 1000
    ListView1.ColumnHeaders.Add , , "���κ�", 1000
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "��������", 1000
    ListView1.ColumnHeaders.Add , , "����ʱ��", 1000
    ListView1.ColumnHeaders.Add , , "��λ����", 1000
    ListView1.ColumnHeaders.Add , , "��λ��", 800
    ListView1.ColumnHeaders.Add , , "�ܼ�", 600
    ListView1.ColumnHeaders.Add , , "�Ƿ���֧��", 1300
    ListView1.ColumnHeaders.Add , , "�Ƿ�ȡ��", 1000
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
If orders.Item(i).Item("has_paid") = 1 Then
Dim itm As ListItem
Set itm = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)  '���
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
itm.SubItems(11) = "��֧��"
Else
itm.SubItems(11) = "δ֧��"
End If
If orders.Item(i).Item("has_cancelled") = 1 Then
itm.SubItems(12) = "��ȡ��"
Else
itm.SubItems(12) = "δȡ��"
End If
End If
        Next i
    End If
    

End Sub

Private Sub Command3_Click()
ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    ListView1.GridLines = True
    ListView1.LabelEdit = lvwManual
    ListView1.FullRowSelect = True
    ListView1.ColumnHeaders.Add , , "#", 0
    ListView1.ColumnHeaders.Add , , "������ˮ��", 1300
    ListView1.ColumnHeaders.Add , , "�û�ID", 1000
    ListView1.ColumnHeaders.Add , , "���κ�", 1000
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "��������", 1000
    ListView1.ColumnHeaders.Add , , "����ʱ��", 1000
    ListView1.ColumnHeaders.Add , , "��λ����", 1000
    ListView1.ColumnHeaders.Add , , "��λ��", 800
    ListView1.ColumnHeaders.Add , , "�ܼ�", 600
    ListView1.ColumnHeaders.Add , , "�Ƿ���֧��", 1300
    ListView1.ColumnHeaders.Add , , "�Ƿ�ȡ��", 1000
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
If orders.Item(i).Item("has_paid") = 0 Then
Dim itm As ListItem
Set itm = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)  '���
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
itm.SubItems(11) = "��֧��"
Else
itm.SubItems(11) = "δ֧��"
End If
If orders.Item(i).Item("has_cancelled") = 1 Then
itm.SubItems(12) = "��ȡ��"
Else
itm.SubItems(12) = "δȡ��"
End If
End If
        Next i
    End If
    

End Sub

Private Sub Command4_Click()
ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    ListView1.GridLines = True
    ListView1.LabelEdit = lvwManual
    ListView1.FullRowSelect = True
    ListView1.ColumnHeaders.Add , , "#", 0
    ListView1.ColumnHeaders.Add , , "������ˮ��", 1300
    ListView1.ColumnHeaders.Add , , "�û�ID", 1000
    ListView1.ColumnHeaders.Add , , "���κ�", 1000
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "��������", 1000
    ListView1.ColumnHeaders.Add , , "����ʱ��", 1000
    ListView1.ColumnHeaders.Add , , "��λ����", 1000
    ListView1.ColumnHeaders.Add , , "��λ��", 800
    ListView1.ColumnHeaders.Add , , "�ܼ�", 600
    ListView1.ColumnHeaders.Add , , "�Ƿ���֧��", 1300
    ListView1.ColumnHeaders.Add , , "�Ƿ�ȡ��", 1000
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
Dim a As String
a = InputBox("��ѯ����(��ʽXXXX - XX - XX)")
        For i = 1 To ccount
        If orders.Item(i).Item("departure_date") = a Then
Dim itm As ListItem
Set itm = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)  '���
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
itm.SubItems(11) = "��֧��"
Else
itm.SubItems(11) = "δ֧��"
End If
If orders.Item(i).Item("has_cancelled") = 1 Then
itm.SubItems(12) = "��ȡ��"
Else
itm.SubItems(12) = "δȡ��"
End If
End If
        Next i
    End If
End Sub

Private Sub Form_Load()
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.View = lvwReport
    ListView1.GridLines = True
    ListView1.LabelEdit = lvwManual
    ListView1.FullRowSelect = True
    ListView1.ColumnHeaders.Add , , "#", 0
    ListView1.ColumnHeaders.Add , , "������ˮ��", 1300
    ListView1.ColumnHeaders.Add , , "�û�ID", 1000
    ListView1.ColumnHeaders.Add , , "���κ�", 1000
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "�û�����վ", 1300
    ListView1.ColumnHeaders.Add , , "��������", 1000
    ListView1.ColumnHeaders.Add , , "����ʱ��", 1000
    ListView1.ColumnHeaders.Add , , "��λ����", 1000
    ListView1.ColumnHeaders.Add , , "��λ��", 800
    ListView1.ColumnHeaders.Add , , "�ܼ�", 600
    ListView1.ColumnHeaders.Add , , "�Ƿ���֧��", 1300
    ListView1.ColumnHeaders.Add , , "�Ƿ�ȡ��", 1000
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
Set itm = ListView1.ListItems.Add(, , ListView1.ListItems.count + 1)  '���
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
itm.SubItems(11) = "��֧��"
Else
itm.SubItems(11) = "δ֧��"
End If
If orders.Item(i).Item("has_cancelled") = 1 Then
itm.SubItems(12) = "��ȡ��"
Else
itm.SubItems(12) = "δȡ��"
End If
        Next i
    End If
    
End Sub

