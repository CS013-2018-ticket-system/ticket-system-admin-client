Attribute VB_Name = "Service"
Public Function makeLoginJson(Username As String, Password As String)
    Dim p As Object
    Dim sInputJson As String

    Set p = JSON.parse("{}")
    p.Add "username", Username
    p.Add "password", Password
   
    makeLoginJson = JSON.toString(p)
End Function

Public Function postLogin(Data As String)
    Dim base_url As String
    Dim ret_obj As Object
    
    base_url = LoadResString(101)
    
    Set ret_obj = HTTP.GetJsonResponse(base_url & "api/admin/login", "POST", Data)
    
    MsgBox ret_obj.Item("success")
    
End Function
