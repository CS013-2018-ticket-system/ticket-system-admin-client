VERSION 5.00
Begin VB.Form frmOrder 
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   6615
      Left            =   1680
      TabIndex        =   0
      Top             =   960
      Width           =   14055
      Begin VB.ListBox List12 
         Height          =   3840
         Left            =   12840
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List11 
         Height          =   3840
         Left            =   11760
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List10 
         Height          =   3840
         Left            =   10680
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List9 
         Height          =   3840
         Left            =   9600
         TabIndex        =   9
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List8 
         Height          =   3840
         Left            =   8520
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List7 
         Height          =   3840
         Left            =   7440
         TabIndex        =   7
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List6 
         Height          =   3840
         Left            =   6360
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List5 
         Height          =   3840
         Left            =   5280
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List4 
         Height          =   3840
         Left            =   4200
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List3 
         Height          =   3840
         Left            =   3120
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List2 
         Height          =   3840
         Left            =   2040
         TabIndex        =   2
         Top             =   1440
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   3840
         Left            =   960
         TabIndex        =   1
         Top             =   1440
         Width           =   1095
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
Set ret_obj = Service.getOrders(token)
If ret_obj.Item("success") = "True" Then
    ccount = ret_obj.Item("count")
    Print ccount
    Set orders = ret_obj.Item("data")
    For i = 1 To ccount
    List1.AddItem orders.Item(i).Item("id")
    List2.AddItem orders.Item(i).Item("user_id")
    List3.AddItem orders.Item(i).Item("train_code")
    List4.AddItem orders.Item(i).Item("from_station")
    List5.AddItem orders.Item(i).Item("to_station")
    List6.AddItem orders.Item(i).Item("departure_date")
    List7.AddItem orders.Item(i).Item("departure_time")
    List8.AddItem orders.Item(i).Item("seat_type")
    If orders.Item(i).Item("seat_no") <> Null Then
    List9.AddItem orders.Item(i).Item("seat_no")
    End If
    List10.AddItem orders.Item(i).Item("price")
    List11.AddItem orders.Item(i).Item("has_paid")
    List12.AddItem orders.Item(i).Item("has_cancelled")
    Next i
    
    End If
    
End Sub

