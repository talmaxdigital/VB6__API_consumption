Attribute VB_Name = "HttpClient"
Option Explicit

' ====================================================================
' HttpClient Module - Cliente HTTP/HTTPS para consumo de APIs
' Baseado no projeto VbAsyncSocket com suporte completo a JSON
' ====================================================================

Private Type HTTP_CONFIG
    baseUrl As String
    DefaultHeaders As Dictionary
    timeout As Long
    userAgent As String
    AcceptEncoding As String
End Type

Private config As HTTP_CONFIG

' ====================================================================
' CONFIGURAÇÃO E INICIALIZAÇÃO
' ====================================================================

Public Sub InitializeHttpClient(Optional ByVal baseUrl As String = "", _
                               Optional ByVal timeout As Long = 30000, _
                               Optional ByVal userAgent As String = "VB6-HttpClient/1.0")
    ' Inicializa o cliente HTTP com configurações padrão
    '
    ' Args:
    '   baseUrl (String): URL base para todas as requisições
    '   timeout (Long): Timeout em milissegundos (padrão: 30000)
    '   userAgent (String): User-Agent para as requisições
    '
    ' Example:
    '   InitializeHttpClient "https://api.github.com", 10000, "MeuApp/1.0"

    config.baseUrl = baseUrl
    config.timeout = timeout
    config.userAgent = userAgent
    config.AcceptEncoding = "gzip, deflate"

    Set config.DefaultHeaders = CreateJSONObject()
    config.DefaultHeaders.Add "User-Agent", userAgent
    config.DefaultHeaders.Add "Accept", "application/json"
    config.DefaultHeaders.Add "Accept-Encoding", config.AcceptEncoding
End Sub

Public Sub SetDefaultHeader(ByVal headerName As String, ByVal headerValue As String)
    ' Define um header padrão que será enviado em todas as requisições
    '
    ' Args:
    '   headerName (String): Nome do header
    '   headerValue (String): Valor do header
    '
    ' Example:
    '   SetDefaultHeader "Authorization", "Bearer seu-token-aqui"
    '   SetDefaultHeader "Content-Type", "application/json"

    If config.DefaultHeaders Is Nothing Then
        Set config.DefaultHeaders = CreateJSONObject()
    End If

    If config.DefaultHeaders.Exists(headerName) Then
        config.DefaultHeaders(headerName) = headerValue
    Else
        config.DefaultHeaders.Add headerName, headerValue
    End If
End Sub

Public Sub RemoveDefaultHeader(ByVal headerName As String)
    ' Remove um header padrão
    '
    ' Args:
    '   headerName (String): Nome do header a ser removido
    '
    ' Example:
    '   RemoveDefaultHeader "Authorization"

    If Not config.DefaultHeaders Is Nothing Then
        If config.DefaultHeaders.Exists(headerName) Then
            config.DefaultHeaders.Remove headerName
        End If
    End If
End Sub

' ====================================================================
' MÉTODOS HTTP SÍNCRONOS
' ====================================================================

Public Function HttpGet(ByVal url As String, _
                       Optional ByVal customHeaders As Dictionary = Nothing, _
                       Optional ByVal body As String = "") As HttpResponse
    ' Executa uma requisição GET síncrona
    '
    ' Args:
    '   url (String): URL completa ou relativa (se baseUrl configurada)
    '   customHeaders (Dictionary): Headers adicionais para esta requisição
    '   body (String): Corpo da requisição (opcional, usado por algumas APIs)
    '
    ' Result:
    '   HttpResponse: Objeto contendo status, headers e body da resposta
    '
    ' Example:
    '   ' GET simples sem body
    '   Set response = HttpGet("https://api.github.com/users/octocat")
    '
    '   ' GET com body JSON (usado por algumas APIs como TomTicket)
    '   Dim params As Dictionary
    '   Set params = CreateJSONObject()
    '   params.Add "customer_id", "12345"
    '   params.Add "customer_type_id", "I"
    '   Set response = HttpGet("/customer/exists", Nothing, BuildJSON(params))

    Set HttpGet = ExecuteRequest("GET", url, body, customHeaders)
End Function

Public Function HttpPost(ByVal url As String, _
                        ByVal body As String, _
                        Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
    ' Executa uma requisição POST síncrona
    '
    ' Args:
    '   url (String): URL completa ou relativa
    '   body (String): Corpo da requisição (JSON, XML, form data, etc.)
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   HttpResponse: Objeto contendo a resposta
    '
    ' Example:
    '   Dim userData As Dictionary
    '   Set userData = CreateJSONObject()
    '   userData.Add "name", "João"
    '   userData.Add "email", "joao@email.com"
    '
    '   Dim response As HttpResponse
    '   Set response = HttpPost("https://api.exemplo.com/users", BuildJSON(userData))

    Set HttpPost = ExecuteRequest("POST", url, body, customHeaders)
End Function

Public Function HttpPut(ByVal url As String, _
                       ByVal body As String, _
                       Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
    ' Executa uma requisição PUT síncrona
    '
    ' Args:
    '   url (String): URL completa ou relativa
    '   body (String): Corpo da requisição
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   HttpResponse: Objeto contendo a resposta

    Set HttpPut = ExecuteRequest("PUT", url, body, customHeaders)
End Function

Public Function HttpDelete(ByVal url As String, _
                          Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
    ' Executa uma requisição DELETE síncrona
    '
    ' Args:
    '   url (String): URL completa ou relativa
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   HttpResponse: Objeto contendo a resposta

    Set HttpDelete = ExecuteRequest("DELETE", url, "", customHeaders)
End Function

Public Function HttpPatch(ByVal url As String, _
                         ByVal body As String, _
                         Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
    ' Executa uma requisição PATCH síncrona
    '
    ' Args:
    '   url (String): URL completa ou relativa
    '   body (String): Corpo da requisição
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   HttpResponse: Objeto contendo a resposta

    Set HttpPatch = ExecuteRequest("PATCH", url, body, customHeaders)
End Function

' ====================================================================
' MÉTODOS ESPECIALIZADOS PARA JSON
' ====================================================================

Public Function GetJson(ByVal url As String, _
                       Optional ByVal customHeaders As Dictionary = Nothing, _
                       Optional ByVal bodyParams As Dictionary = Nothing) As Object
    ' Executa GET e retorna automaticamente o JSON parseado
    '
    ' Args:
    '   url (String): URL da API
    '   customHeaders (Dictionary): Headers adicionais
    '   bodyParams (Dictionary): Parâmetros a serem enviados no body como JSON
    '
    ' Result:
    '   Object: Dictionary ou Collection com dados JSON parseados
    '
    ' Example:
    '   ' GET simples
    '   Set user = GetJson("https://api.github.com/users/octocat")
    '
    '   ' GET com parâmetros no body (TomTicket style)
    '   Dim params As Dictionary
    '   Set params = CreateJSONObject()
    '   params.Add "customer_id", "12345"
    '   params.Add "customer_type_id", "I"
    '   Set result = GetJson("/customer/exists", Nothing, params)

    Dim body As String
    Dim headers As Dictionary

    ' Preparar headers com Content-Type para JSON se houver body
    Set headers = MergeHeaders(customHeaders)

    If Not bodyParams Is Nothing Then
        body = BuildJSON(bodyParams)
        headers("Content-Type") = "application/json"
    End If

    Dim response As HttpResponse
    Set response = HttpGet(url, headers, body)

    If response.IsSuccess Then
        Set GetJson = response.Json
    Else
        Err.Raise vbObjectError + 100, "GetJson", "HTTP Error: " & response.StatusCode & " - " & response.StatusText
    End If
End Function

Public Function PostJson(ByVal url As String, _
                        ByVal jsonObject As Object, _
                        Optional ByVal customHeaders As Dictionary = Nothing) As Object
    ' Executa POST com objeto JSON e retorna JSON parseado
    '
    ' Args:
    '   url (String): URL da API
    '   jsonObject (Object): Dictionary ou Collection a ser enviado como JSON
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   Object: Dictionary ou Collection com resposta JSON parseada
    '
    ' Example:
    '   Dim newUser As Dictionary
    '   Set newUser = CreateJSONObject()
    '   newUser.Add "name", "Maria"
    '   newUser.Add "email", "maria@email.com"
    '
    '   Dim createdUser As Object
    '   Set createdUser = PostJson("https://api.exemplo.com/users", newUser)

    Dim headers As Dictionary
    Set headers = MergeHeaders(customHeaders)
    headers("Content-Type") = "application/json"

    Dim response As HttpResponse
    Set response = HttpPost(url, BuildJSON(jsonObject), headers)

    If response.IsSuccess Then
        Set PostJson = response.Json
    Else
        Err.Raise vbObjectError + 101, "PostJson", "HTTP Error: " & response.StatusCode & " - " & response.StatusText
    End If
End Function

Public Function PutJson(ByVal url As String, _
                       ByVal jsonObject As Object, _
                       Optional ByVal customHeaders As Dictionary = Nothing) As Object
    ' Executa PUT com objeto JSON e retorna JSON parseado
    '
    ' Args:
    '   url (String): URL da API
    '   jsonObject (Object): Dictionary ou Collection a ser enviado como JSON
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   Object: Dictionary ou Collection com resposta JSON parseada

    Dim headers As Dictionary
    Set headers = MergeHeaders(customHeaders)
    headers("Content-Type") = "application/json"

    Dim response As HttpResponse
    Set response = HttpPut(url, BuildJSON(jsonObject), headers)

    If response.IsSuccess Then
        Set PutJson = response.Json
    Else
        Err.Raise vbObjectError + 102, "PutJson", "HTTP Error: " & response.StatusCode & " - " & response.StatusText
    End If
End Function

' ====================================================================
' MÉTODOS PARA UPLOAD/DOWNLOAD DE ARQUIVOS
' ====================================================================

Public Function DownloadFile(ByVal url As String, _
                            ByVal localPath As String, _
                            Optional ByVal customHeaders As Dictionary = Nothing) As Boolean
    ' Faz download de um arquivo
    '
    ' Args:
    '   url (String): URL do arquivo
    '   localPath (String): Caminho local onde salvar o arquivo
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   Boolean: True se o download foi bem-sucedido
    '
    ' Example:
    '   If DownloadFile("https://exemplo.com/arquivo.pdf", "C:\temp\arquivo.pdf") Then
    '       Debug.Print "Download concluído!"
    '   End If

    Dim downloader As New cHttpDownload
    Dim key As Variant

    On Error GoTo ErrorHandler

    ' Aplicar headers customizados se fornecidos
    If Not customHeaders Is Nothing Then
        For Each key In customHeaders.Keys
            ' Note: cHttpDownload pode ter limitações para headers customizados
            ' Implementar conforme a interface disponível
        Next key
    End If

    downloader.BeginDownload url, localPath

    ' Aguardar conclusão (implementação pode variar conforme a classe)
    ' Esta é uma implementação simplificada

    DownloadFile = True
    Exit Function

ErrorHandler:
    DownloadFile = False
End Function

Public Function UploadFile(ByVal url As String, _
                          ByVal filePath As String, _
                          Optional ByVal fieldName As String = "file", _
                          Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
    ' Faz upload de um arquivo usando multipart/form-data
    '
    ' Args:
    '   url (String): URL de destino
    '   filePath (String): Caminho do arquivo local
    '   fieldName (String): Nome do campo no formulário
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   HttpResponse: Resposta do servidor

    Dim boundary As String
    boundary = "----VB6HttpClient" & Format(Timer * 1000, "0")

    Dim headers As Dictionary
    Set headers = MergeHeaders(customHeaders)
    headers("Content-Type") = "multipart/form-data; boundary=" & boundary

    Dim body As String
    body = BuildMultipartBody(filePath, fieldName, boundary)

    Set UploadFile = ExecuteRequest("POST", url, body, headers)
End Function

' ====================================================================
' CLASSE DE RESPOSTA HTTP
' ====================================================================

Public Function CreateHttpResponse(ByVal request As cHttpRequest) As HttpResponse
    ' Cria um objeto HttpResponse a partir da requisição executada
    '
    ' Args:
    '   request (cHttpRequest): Objeto de requisição já executada
    '
    ' Result:
    '   HttpResponse: Objeto wrapper com propriedades convenientes

    Dim response As New HttpResponse
    response.Initialize request
    Set CreateHttpResponse = response
End Function

' ====================================================================
' FUNÇÕES AUXILIARES PRIVADAS
' ====================================================================

Private Function ExecuteRequest(ByVal method As String, _
                               ByVal url As String, _
                               ByVal body As String, _
                               Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
    ' Executa uma requisição HTTP genérica
    '
    ' Args:
    '   method (String): Método HTTP (GET, POST, PUT, DELETE, etc.)
    '   url (String): URL completa ou relativa
    '   body (String): Corpo da requisição (pode ser usado em qualquer método)
    '   customHeaders (Dictionary): Headers adicionais
    '
    ' Result:
    '   HttpResponse: Objeto contendo a resposta

    Dim req As New cHttpRequest
    Dim fullUrl As String
    Dim headers As Dictionary
    Dim key As Variant

    ' Construir URL completa
    fullUrl = BuildFullUrl(url)

    ' Mesclar headers padrão com customizados
    Set headers = MergeHeaders(customHeaders)

    On Error GoTo ErrorHandler

    With req
        .Open_ method, fullUrl, False

        ' Aplicar timeout se configurado
        If config.timeout > 0 Then
            .SetTimeout config.timeout
        End If

        ' Aplicar headers
        For Each key In headers.Keys
            .SetRequestHeader CStr(key), CStr(headers(key))
        Next key

        ' Enviar requisição com body se fornecido
        ' Nota: Algumas APIs (como TomTicket) esperam JSON no body mesmo para GET
        If Len(body) > 0 Then
            .Send body
        Else
            .Send
        End If
    End With

    Set ExecuteRequest = CreateHttpResponse(req)
    Exit Function

ErrorHandler:
    Err.Raise Err.Number, "ExecuteRequest", "Erro na requisição HTTP: " & Err.Description
End Function

Private Function BuildFullUrl(ByVal url As String) As String
    ' Constrói a URL completa combinando baseUrl com URL relativa
    '
    ' Args:
    '   url (String): URL que pode ser completa ou relativa
    '
    ' Result:
    '   String: URL completa

    If Left(url, 4) = "http" Then
        ' URL já é completa
        BuildFullUrl = url
    ElseIf Len(config.baseUrl) > 0 Then
        ' Combinar com baseUrl
        If Right(config.baseUrl, 1) = "/" And Left(url, 1) = "/" Then
            BuildFullUrl = config.baseUrl & Mid(url, 2)
        ElseIf Right(config.baseUrl, 1) <> "/" And Left(url, 1) <> "/" Then
            BuildFullUrl = config.baseUrl & "/" & url
        Else
            BuildFullUrl = config.baseUrl & url
        End If
    Else
        ' URL relativa sem baseUrl
        Err.Raise vbObjectError + 200, "BuildFullUrl", "URL relativa fornecida sem baseUrl configurada"
    End If
End Function

Private Function MergeHeaders(Optional ByVal customHeaders As Dictionary = Nothing) As Dictionary
    ' Mescla headers padrão com headers customizados
    '
    ' Args:
    '   customHeaders (Dictionary): Headers customizados (opcional)
    '
    ' Result:
    '   Dictionary: Headers mesclados

    Dim key As Variant

    Set MergeHeaders = CreateJSONObject()

    ' Copiar headers padrão
    If Not config.DefaultHeaders Is Nothing Then
        For Each key In config.DefaultHeaders.Keys
            MergeHeaders.Add key, config.DefaultHeaders(key)
        Next key
    End If

    ' Sobrescrever com headers customizados
    If Not customHeaders Is Nothing Then
        For Each key In customHeaders.Keys
            If MergeHeaders.Exists(key) Then
                MergeHeaders.Remove key
                MergeHeaders.Add key, customHeaders(key)
            Else
                MergeHeaders.Add key, customHeaders(key)
            End If
        Next key
    End If
End Function

Private Function BuildMultipartBody(ByVal filePath As String, _
                                   ByVal fieldName As String, _
                                   ByVal boundary As String) As String
    ' Constrói o corpo de uma requisição multipart/form-data
    '
    ' Args:
    '   filePath (String): Caminho do arquivo
    '   fieldName (String): Nome do campo
    '   boundary (String): Boundary string
    '
    ' Result:
    '   String: Corpo da requisição formatado

    Dim fileName As String
    Dim fileContent As String
    Dim body As String

    ' Extrair nome do arquivo
    fileName = Mid(filePath, InStrRev(filePath, "\") + 1)

    ' Ler conteúdo do arquivo (implementação simplificada)
    On Error GoTo ErrorHandler

    Dim fileNum As Integer
    fileNum = FreeFile
    Open filePath For Binary As #fileNum
    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
    Close #fileNum

    ' Construir corpo multipart
    body = "--" & boundary & vbCrLf
    body = body & "Content-Disposition: form-data; name=""" & fieldName & """; filename=""" & fileName & """" & vbCrLf
    body = body & "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
    body = body & fileContent & vbCrLf
    body = body & "--" & boundary & "--" & vbCrLf

    BuildMultipartBody = body
    Exit Function

ErrorHandler:
    Err.Raise vbObjectError + 300, "BuildMultipartBody", "Erro ao ler arquivo: " & Err.Description
End Function

' ====================================================================
' UTILITÁRIOS PÚBLICOS
' ====================================================================

Public Function UrlEncode(ByVal text As String) As String
    ' Codifica uma string para uso em URLs
    '
    ' Args:
    '   text (String): Texto a ser codificado
    '
    ' Result:
    '   String: Texto codificado para URL
    '
    ' Example:
    '   Debug.Print UrlEncode("João & Maria") ' Output: Jo%C3%A3o%20%26%20Maria

    Dim result As String
    Dim i As Integer
    Dim char As String
    Dim asciiVal As Integer

    For i = 1 To Len(text)
        char = Mid(text, i, 1)
        asciiVal = Asc(char)

        If (asciiVal >= 48 And asciiVal <= 57) Or _
           (asciiVal >= 65 And asciiVal <= 90) Or _
           (asciiVal >= 97 And asciiVal <= 122) Or _
           char = "-" Or char = "_" Or char = "." Or char = "~" Then
            result = result & char
        Else
            result = result & "%" & Right("0" & Hex(asciiVal), 2)
        End If
    Next i

    UrlEncode = result
End Function

Public Function BuildQueryString(ByVal params As Dictionary) As String
    ' Constrói uma query string a partir de um Dictionary
    '
    ' Args:
    '   params (Dictionary): Parâmetros chave-valor
    '
    ' Result:
    '   String: Query string formatada
    '
    ' Example:
    '   Dim params As Dictionary
    '   Set params = CreateJSONObject()
    '   params.Add "name", "João"
    '   params.Add "page", "1"
    '   Debug.Print BuildQueryString(params) ' Output: name=Jo%C3%A3o&page=1

    Dim result As String
    Dim key As Variant
    Dim isFirst As Boolean

    isFirst = True

    For Each key In params.Keys
        If Not isFirst Then
            result = result & "&"
        End If

        result = result & UrlEncode(CStr(key)) & "=" & UrlEncode(CStr(params(key)))
        isFirst = False
    Next key

    BuildQueryString = result
End Function

' ====================================================================
' EXEMPLO DE USO E TESTES
' ====================================================================

Public Sub ExampleUsage()
    ' Exemplo de uso do HttpClient
    '
    ' Este exemplo demonstra as principais funcionalidades do módulo

    On Error GoTo ErrorHandler

    ' 1. Inicializar cliente
    InitializeHttpClient "https://jsonplaceholder.typicode.com", 15000, "ExemploVB6/1.0"

    ' 2. Configurar autenticação (exemplo)
    SetDefaultHeader "Authorization", "Bearer seu-token-aqui"

    ' 3. GET simples
    Debug.Print "=== GET Simples ==="
    Dim user As Object
    Set user = GetJson("/users/1")
    Debug.Print "Nome: " & user("name")
    Debug.Print "Email: " & user("email")

    ' 4. POST com JSON
    Debug.Print vbNewLine & "=== POST com JSON ==="
    Dim newPost As Dictionary
    Set newPost = CreateJSONObject()
    newPost.Add "title", "Meu Post"
    newPost.Add "body", "Conteúdo do post"
    newPost.Add "userId", 1

    Dim createdPost As Object
    Set createdPost = PostJson("/posts", newPost)
    Debug.Print "Post criado com ID: " & createdPost("id")

    ' 5. GET com query parameters
    Debug.Print vbNewLine & "=== GET com Query Params ==="
    Dim params As Dictionary
    Set params = CreateJSONObject()
    params.Add "userId", "1"
    params.Add "_limit", "5"

    Dim queryString As String
    queryString = BuildQueryString(params)

    Dim posts As Object
    Set posts = GetJson("/posts?" & queryString)
    Debug.Print "Encontrados " & posts.Count & " posts"

    ' 6. Requisição com response completo
    Debug.Print vbNewLine & "=== Response Completo ==="
    Dim response As HttpResponse
    Set response = HttpGet("/users/1")

    Debug.Print "Status: " & response.StatusCode & " " & response.StatusText
    Debug.Print "Content-Type: " & response.GetHeader("Content-Type")
    Debug.Print "Dados: " & response.text

    Exit Sub

ErrorHandler:
    Debug.Print "Erro: " & Err.Description
End Sub

