Attribute VB_Name = "HTTP"
Private aHttpRequest As WinHttp.WinHttpRequest

Public Function GetJsonResponse(sURL As String, Optional Method As String, Optional Data As String) As Object
    Dim sMethod             As String
    Dim sBody               As String
    Dim content As String
    
    sBody = Data
    If Method = "" Then
        sMethod = "POST"        '或者(GET)
    Else
        sMethod = Method
    End If
   
    ''创建WinHttp.WinHttpRequest
    Set aHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
   
    '' 同步接收数据
    aHttpRequest.Open sMethod, sURL, False
    '' 非常重要(忽略错误)
    aHttpRequest.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    '' 其它请求头设置
    aHttpRequest.SetRequestHeader "Content-Type", "application/json"
    'aHttpRequest.setRequestHeader "Content-Length", Len(sBody)
   
    '' 发送
    aHttpRequest.Send sBody
    

    '' 得到返回文本(或者是其它)
    content = aHttpRequest.ResponseText
    
    Dim ret As Object
    
    Set aHttpRequest = Nothing
    Set GetJsonResponse = JSON.parse(content)

End Function
