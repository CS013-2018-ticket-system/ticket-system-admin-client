Attribute VB_Name = "HTTP"
Private aHttpRequest As WinHttp.WinHttpRequest

Public Function GetJsonResponse(sURL As String, Optional Method As String, Optional Data As String) As Object
    Dim sMethod             As String
    Dim sBody               As String
    Dim content As String
    
    sBody = Data
    If Method = "" Then
        sMethod = "POST"        '����(GET)
    Else
        sMethod = Method
    End If
   
    ''����WinHttp.WinHttpRequest
    Set aHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
   
    '' ͬ����������
    aHttpRequest.Open sMethod, sURL, False
    '' �ǳ���Ҫ(���Դ���)
    aHttpRequest.Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300
    '' ��������ͷ����
    aHttpRequest.SetRequestHeader "Content-Type", "application/json"
    'aHttpRequest.setRequestHeader "Content-Length", Len(sBody)
   
    '' ����
    aHttpRequest.Send sBody
    

    '' �õ������ı�(����������)
    content = aHttpRequest.ResponseText
    
    Dim ret As Object
    
    Set aHttpRequest = Nothing
    Set GetJsonResponse = JSON.parse(content)

End Function
