Attribute VB_Name = "Service"
Public Function makeLoginJson(Username As String, Password As String)
    Dim p As Object
    Dim sInputJson As String

    Set p = JSON.parse("{}")
    p.Add "username", Username
    p.Add "password", Password
   
    makeLoginJson = JSON.toString(p)
End Function

Public Function constructUrl(Path As String, token As String)
    constructUrl = LoadResString(101) & Path & "?access_token=" & token
End Function

Public Function postLogin(Data As String)
    Dim base_url As String
    Dim ret_obj As Object
    
    base_url = LoadResString(101)
    
    Set ret_obj = HTTP.GetJsonResponse(base_url & "api/admin/login", "POST", Data)
    
    postLogin = IIf(ret_obj.Item("success") = "True", ret_obj.Item("token"), False)
    
End Function

Public Function getUsers(token As String)
    Dim ret(2)
    Set ret_obj = HTTP.GetJsonResponse(constructUrl("api/admin/users/all", token), "GET")
    
    If ret_obj.Item("success") = "False" Then
        MsgBox ret_obj.Item("msg")
        ret_obj = Null
    Else
        Set getUsers = ret_obj
    End If
End Function

Public Function getOrders(token As String)
    Dim ret(2)
    Set ret_obj = HTTP.GetJsonResponse(constructUrl("api/admin/orders/all", token), "GET")
    If ret_obj.Item("success") = "False" Then
        MsgBox ret_obj.Item("msg")
        ret_obj = Null
    Else
        Set getOrders = ret_obj
    End If
End Function

Public Function getRefunds(token As String, count As Integer, page As Integer, gettype As String)
    Dim ret(2)
    Set ret_obj = HTTP.GetJsonResponse(constructUrl("api/admin/refund/get", token) & "&count=" & count & "&offset=" & count * page & "&type=" & gettype, "GET")
    If ret_obj.Item("success") = "False" Then
        MsgBox ret_obj.Item("msg")
        ret_obj = Null
    Else
        Set getRefunds = ret_obj
    End If
End Function
