VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmCancel 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6075
   StartUpPosition =   3  '����ȱʡ
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   353
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ѧ������"
         Object.Width           =   1524
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "�˿���"
         Object.Width           =   1524
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "����ʱ��"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�������"
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

Private Sub renderTable(count As Integer, page As Integer, gettype As String)
    ListView1.ListItems.Clear
    Set refunds = Service.getRefunds(token, count, page, gettype)
    
    For i = 1 To refunds.Item("count")
        Dim itm As ListItem
        Set itm = ListView1.ListItems.Add(, , refunds.Item("data").Item(i).Item("id"))
        itm.SubItems(1) = refunds.Item("data").Item(i).Item("user").Item("name")
        itm.SubItems(2) = refunds.Item("data").Item(i).Item("order").Item("price")
        itm.SubItems(3) = refunds.Item("data").Item(i).Item("created_at")
        itm.SubItems(4) = IIf(refunds.Item("data").Item(i).Item("has_confirmed"), "�Ѵ���", "δ����")
    Next
    
End Sub

Private Sub Form_Load()
    renderTable 10, 0, "all"
End Sub
