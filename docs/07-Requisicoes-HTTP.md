# Requisições HTTP - Guia Completo

Este guia cobre todos os aspectos das requisições HTTP no sistema, desde operações básicas até cenários avançados de uso.

## 📋 Índice

- [Métodos HTTP Fundamentais](#métodos-http-fundamentais)
- [Configuração de Requisições](#configuração-de-requisições)
- [Tratamento de Respostas](#tratamento-de-respostas)
- [Cenários Avançados](#cenários-avançados)
- [Padrões de Uso por Tipo de API](#padrões-de-uso-por-tipo-de-api)

## Métodos HTTP Fundamentais

### GET - Consulta de Dados

```vb
' GET básico - URL completa
Dim response As HttpResponse
Set response = HttpGet("https://jsonplaceholder.typicode.com/posts/1")

If response.IsSuccess Then
    Debug.Print "Post encontrado: " & response.Text
End If

' GET com URL relativa (requer InitializeHttpClient com baseUrl)
InitializeHttpClient "https://jsonplaceholder.typicode.com"
Set response = HttpGet("/posts/1")

' GET com headers customizados
Dim headers As Dictionary
Set headers = CreateJSONObject()
headers.Add "Accept", "application/json"
headers.Add "User-Agent", "MeuApp/1.0"
Set response = HttpGet("/posts/1", headers)

' GET com parâmetros no body (usado por algumas APIs)
Dim searchParams As String
searchParams = "{""category"":""electronics"",""minPrice"":100}"
Set response = HttpGet("/products/search", Nothing, searchParams)
```

### POST - Criação de Recursos

```vb
' POST com JSON
Dim novoPost As Dictionary
Set novoPost = CreateJSONObject()
novoPost.Add "title", "Meu novo post"
novoPost.Add "body", "Conteúdo do post criado via VB6"
novoPost.Add "userId", 1

Dim jsonData As String
jsonData = BuildJSON(novoPost)

Dim response As HttpResponse
Set response = HttpPost("https://jsonplaceholder.typicode.com/posts", jsonData)

If response.IsSuccess Then
    Debug.Print "Post criado com ID: " & ParseJSON(response.Text)("id")
End If

' POST com form data
Dim formData As String
formData = "name=João&email=joao@email.com&age=30"

Dim headers As Dictionary
Set headers = CreateJSONObject()
headers.Add "Content-Type", "application/x-www-form-urlencoded"

Set response = HttpPost("/users", formData, headers)

' POST multipart (upload de arquivo)
Dim boundary As String
boundary = "----VB6FormBoundary" & Format(Timer * 1000, "0")

Dim multipartData As String
multipartData = "--" & boundary & vbCrLf
multipartData = multipartData & "Content-Disposition: form-data; name=""file""; filename=""dados.txt""" & vbCrLf
multipartData = multipartData & "Content-Type: text/plain" & vbCrLf & vbCrLf
multipartData = multipartData & "Conteúdo do arquivo" & vbCrLf
multipartData = multipartData & "--" & boundary & "--" & vbCrLf

Set headers = CreateJSONObject()
headers.Add "Content-Type", "multipart/form-data; boundary=" & boundary

Set response = HttpPost("/upload", multipartData, headers)
```

### PUT - Atualização Completa

```vb
' PUT para atualizar recurso completo
Dim usuarioAtualizado As Dictionary
Set usuarioAtualizado = CreateJSONObject()
usuarioAtualizado.Add "id", 1
usuarioAtualizado.Add "name", "João Silva Atualizado"
usuarioAtualizado.Add "email", "joao.novo@email.com"
usuarioAtualizado.Add "phone", "(11) 99999-9999"

Dim response As HttpResponse
Set response = HttpPut("/users/1", BuildJSON(usuarioAtualizado))

If response.IsSuccess Then
    Debug.Print "Usuário atualizado com sucesso"
Else
    Debug.Print "Erro ao atualizar: " & response.StatusCode
End If

' PUT com validação de ETag (controle de concorrência)
Dim headers As Dictionary
Set headers = CreateJSONObject()
headers.Add "If-Match", """abc123def456"""  ' ETag obtido em GET anterior

Set response = HttpPut("/users/1", BuildJSON(usuarioAtualizado), headers)

If response.StatusCode = 412 Then
    Debug.Print "Conflito: recurso foi modificado por outro usuário"
ElseIf response.IsSuccess Then
    Debug.Print "Atualização bem-sucedida"
End If
```

### PATCH - Atualização Parcial

```vb
' PATCH para atualizar apenas campos específicos
Dim atualizacaoParcial As Dictionary
Set atualizacaoParcial = CreateJSONObject()
atualizacaoParcial.Add "email", "novo.email@exemplo.com"
atualizacaoParcial.Add "phone", "(11) 88888-8888"

Dim response As HttpResponse
Set response = HttpPatch("/users/1", BuildJSON(atualizacaoParcial))

' PATCH com JSON Patch (RFC 6902)
Dim jsonPatch As Collection
Set jsonPatch = CreateJSONArray()

Dim operacao1 As Dictionary
Set operacao1 = CreateJSONObject()
operacao1.Add "op", "replace"
operacao1.Add "path", "/email"
operacao1.Add "value", "email.atualizado@exemplo.com"
jsonPatch.Add operacao1

Dim operacao2 As Dictionary
Set operacao2 = CreateJSONObject()
operacao2.Add "op", "add"
operacao2.Add "path", "/telefone"
operacao2.Add "value", "(11) 77777-7777"
jsonPatch.Add operacao2

Dim headers As Dictionary
Set headers = CreateJSONObject()
headers.Add "Content-Type", "application/json-patch+json"

Set response = HttpPatch("/users/1", BuildJSON(jsonPatch), headers)
```

### DELETE - Remoção de Recursos

```vb
' DELETE simples
Dim response As HttpResponse
Set response = HttpDelete("/users/1")

If response.StatusCode = 204 Then
    Debug.Print "Usuário removido com sucesso"
ElseIf response.StatusCode = 404 Then
    Debug.Print "Usuário não encontrado"
Else
    Debug.Print "Erro ao remover usuário: " & response.StatusCode
End If

' DELETE com confirmação
Dim headers As Dictionary
Set headers = CreateJSONObject()
headers.Add "X-Confirm-Delete", "true"
headers.Add "X-Reason", "Usuário inativo há mais de 2 anos"

Set response = HttpDelete("/users/1", headers)

' DELETE em lote
Dim idsParaRemover As Collection
Set idsParaRemover = CreateJSONArray()
idsParaRemover.Add 1
idsParaRemover.Add 2
idsParaRemover.Add 3

Dim body As String
body = BuildJSON(idsParaRemover)

' Algumas APIs usam DELETE com body
Set response = HttpDelete("/users/batch", Nothing, body)
```

## Configuração de Requisições

### Headers Essenciais

```vb
' Content-Type para diferentes formatos
Sub ConfigurarContentTypes()
    Dim headers As Dictionary
    Set headers = CreateJSONObject()

    ' JSON (mais comum)
    headers.Add "Content-Type", "application/json"

    ' Form data
    headers.Add "Content-Type", "application/x-www-form-urlencoded"

    ' XML
    headers.Add "Content-Type", "application/xml"

    ' Texto simples
    headers.Add "Content-Type", "text/plain"

    ' Upload de arquivo
    headers.Add "Content-Type", "multipart/form-data; boundary=----FormBoundary"
End Sub

' Accept para especificar formato de resposta desejado
Sub ConfigurarAccept()
    Dim headers As Dictionary
    Set headers = CreateJSONObject()

    ' Preferir JSON
    headers.Add "Accept", "application/json"

    ' Múltiplos formatos com prioridade
    headers.Add "Accept", "application/json, application/xml;q=0.9, text/plain;q=0.8"

    ' Versão específica da API
    headers.Add "Accept", "application/vnd.api.v2+json"
End Sub
```

### Autenticação

```vb
' Bearer Token (JWT, OAuth2)
Sub ConfigurarBearerToken()
    Dim token As String
    token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9..."

    SetDefaultHeader "Authorization", "Bearer " & token
End Sub

' API Key
Sub ConfigurarAPIKey()
    ' No header
    SetDefaultHeader "X-API-Key", "sua-api-key-aqui"

    ' Como parâmetro na URL seria:
    ' InitializeHttpClient "https://api.exemplo.com?api_key=sua-chave"
End Sub

' Basic Authentication
Sub ConfigurarBasicAuth()
    Dim usuario As String
    Dim senha As String
    usuario = "meuusuario"
    senha = "minhasenha"

    Dim credentials As String
    credentials = usuario & ":" & senha

    ' No VB6, você precisaria implementar Base64 encoding
    ' Por simplicidade, assumindo função Base64Encode disponível
    SetDefaultHeader "Authorization", "Basic " & Base64Encode(credentials)
End Sub

' Custom Authentication
Sub ConfigurarCustomAuth()
    SetDefaultHeader "X-Auth-Token", "token-customizado"
    SetDefaultHeader "X-Client-ID", "meu-client-id"
    SetDefaultHeader "X-Timestamp", CStr(DateDiff("s", #1/1/1970#, Now))
End Sub
```

### Timeout e Retry

```vb
' Configurar timeout global
Sub ConfigurarTimeout()
    InitializeHttpClient "https://api.exemplo.com", 45000  ' 45 segundos
End Sub

' Implementar retry com backoff
Function RequisicaoComRetry(url As String, maxTentativas As Integer, Optional metodo As String = "GET", Optional body As String = "") As HttpResponse
    Dim tentativa As Integer
    Dim waitTime As Long
    Dim response As HttpResponse

    For tentativa = 1 To maxTentativas
        On Error GoTo ProximaTentativa

        Select Case UCase(metodo)
            Case "GET":
                Set response = HttpGet(url)
            Case "POST":
                Set response = HttpPost(url, body)
            Case "PUT":
                Set response = HttpPut(url, body)
            Case "DELETE":
                Set response = HttpDelete(url)
        End Select

        ' Se chegou aqui, a requisição foi executada
        If response.IsSuccess Or response.StatusCode < 500 Then
            ' Sucesso ou erro de cliente (não tentar novamente)
            Set RequisicaoComRetry = response
            Exit Function
        End If

ProximaTentativa:
        Debug.Print "Tentativa " & tentativa & " falhou. Status: " & response.StatusCode

        If tentativa < maxTentativas Then
            ' Backoff exponencial: 1s, 2s, 4s, 8s...
            waitTime = 1000 * (2 ^ (tentativa - 1))
            Debug.Print "Aguardando " & waitTime & "ms antes da próxima tentativa..."

            ' Implementar Sleep (pode usar API do Windows ou Timer)
            Sleep waitTime
        End If

        On Error GoTo 0
    Next tentativa

    ' Todas as tentativas falharam
    Set RequisicaoComRetry = response
End Function
```

## Tratamento de Respostas

### Verificação de Status

```vb
Sub ProcessarResponse(response As HttpResponse)
    ' Categorias de status HTTP
    Select Case response.StatusCode
        Case 200 To 299:  ' Sucesso
            Debug.Print "Operação bem-sucedida"
            ProcessarDadosResposta response

        Case 300 To 399:  ' Redirecionamento
            Debug.Print "Redirecionamento: " & response.StatusCode
            Dim novaUrl As String
            novaUrl = response.GetHeader("Location")
            If Len(novaUrl) > 0 Then
                Debug.Print "Nova URL: " & novaUrl
            End If

        Case 400 To 499:  ' Erro do cliente
            Debug.Print "Erro do cliente: " & response.StatusCode
            ProcessarErroCliente response

        Case 500 To 599:  ' Erro do servidor
            Debug.Print "Erro do servidor: " & response.StatusCode
            ProcessarErroServidor response

        Case Else:
            Debug.Print "Status desconhecido: " & response.StatusCode
    End Select
End Sub

Sub ProcessarDadosResposta(response As HttpResponse)
    Dim contentType As String
    contentType = response.GetHeader("Content-Type")

    If InStr(LCase(contentType), "json") > 0 Then
        ' Resposta JSON
        Dim jsonData As Object
        Set jsonData = response.Json

        If Not jsonData Is Nothing Then
            Debug.Print "JSON recebido com sucesso"
            ' Processar dados JSON...
        Else
            Debug.Print "Erro ao parsear JSON"
        End If

    ElseIf InStr(LCase(contentType), "xml") > 0 Then
        ' Resposta XML
        Debug.Print "XML recebido: " & Left(response.Text, 200)

    Else
        ' Texto simples ou outro formato
        Debug.Print "Resposta texto: " & Left(response.Text, 200)
    End If
End Sub

Sub ProcessarErroCliente(response As HttpResponse)
    Select Case response.StatusCode
        Case 400:  ' Bad Request
            Debug.Print "Requisição inválida - verificar dados enviados"

        Case 401:  ' Unauthorized
            Debug.Print "Não autorizado - verificar autenticação"
            RenovarToken()

        Case 403:  ' Forbidden
            Debug.Print "Proibido - sem permissão para este recurso"

        Case 404:  ' Not Found
            Debug.Print "Recurso não encontrado"

        Case 422:  ' Unprocessable Entity
            Debug.Print "Dados inválidos:"
            If InStr(response.GetHeader("Content-Type"), "json") > 0 Then
                Dim erros As Object
                Set erros = response.Json
                ' Processar erros de validação...
            End If

        Case 429:  ' Too Many Requests
            Debug.Print "Rate limit atingido"
            Dim retryAfter As String
            retryAfter = response.GetHeader("Retry-After")
            If Len(retryAfter) > 0 Then
                Debug.Print "Tentar novamente após: " & retryAfter & " segundos"
            End If
    End Select
End Sub

Sub ProcessarErroServidor(response As HttpResponse)
    Select Case response.StatusCode
        Case 500:  ' Internal Server Error
            Debug.Print "Erro interno do servidor - tentar novamente mais tarde"

        Case 502:  ' Bad Gateway
            Debug.Print "Gateway inválido - problema de infraestrutura"

        Case 503:  ' Service Unavailable
            Debug.Print "Serviço indisponível - manutenção ou sobrecarga"

        Case 504:  ' Gateway Timeout
            Debug.Print "Timeout do gateway - operação demorou muito"
    End Select

    ' Log detalhado para erros de servidor
    LogErroServidor response
End Sub
```

### Headers de Resposta Importantes

```vb
Sub AnalisarHeadersResposta(response As HttpResponse)
    ' Content-Type: tipo do conteúdo
    Dim contentType As String
    contentType = response.GetHeader("Content-Type")
    Debug.Print "Tipo de conteúdo: " & contentType

    ' Content-Length: tamanho da resposta
    Dim contentLength As String
    contentLength = response.GetHeader("Content-Length")
    If Len(contentLength) > 0 Then
        Debug.Print "Tamanho: " & contentLength & " bytes"
    End If

    ' ETag: versionamento para cache
    Dim etag As String
    etag = response.GetHeader("ETag")
    If Len(etag) > 0 Then
        Debug.Print "ETag: " & etag
        ' Salvar para futuras requisições condicionais
    End If

    ' Last-Modified: data da última modificação
    Dim lastModified As String
    lastModified = response.GetHeader("Last-Modified")
    If Len(lastModified) > 0 Then
        Debug.Print "Última modificação: " & lastModified
    End If

    ' Rate Limiting
    Dim rateLimit As String
    Dim rateLimitRemaining As String
    rateLimit = response.GetHeader("X-Rate-Limit-Limit")
    rateLimitRemaining = response.GetHeader("X-Rate-Limit-Remaining")

    If Len(rateLimit) > 0 And Len(rateLimitRemaining) > 0 Then
        Debug.Print "Rate Limit: " & rateLimitRemaining & "/" & rateLimit

        If CLng(rateLimitRemaining) < 10 Then
            Debug.Print "Poucas requisições restantes!"
        End If
    End If

    ' Cache-Control: política de cache
    Dim cacheControl As String
    cacheControl = response.GetHeader("Cache-Control")
    If Len(cacheControl) > 0 Then
        Debug.Print "Cache Control: " & cacheControl
    End If

    ' Location: redirecionamento ou localização de recurso criado
    Dim location As String
    location = response.GetHeader("Location")
    If Len(location) > 0 Then
        Debug.Print "Location: " & location
    End If
End Sub
```

## Cenários Avançados

### Paginação de Dados

```vb
' Paginação baseada em página/tamanho
Function ObterTodosDados(baseUrl As String) As Collection
    Set ObterTodosDados = CreateJSONArray()

    Dim pagina As Integer
    Dim tamanhoPagina As Integer
    Dim totalPaginas As Integer

    pagina = 1
    tamanhoPagina = 50
    totalPaginas = 1  ' Será atualizado na primeira requisição

    Do While pagina <= totalPaginas
        Dim url As String
        url = baseUrl & "?page=" & pagina & "&limit=" & tamanhoPagina

        Dim response As HttpResponse
        Set response = HttpGet(url)

        If response.IsSuccess Then
            Dim pageData As Object
            Set pageData = response.Json

            ' Processar dados da página
            If pageData.Exists("data") Then
                Dim items As Collection
                Set items = pageData("data")

                Dim i As Integer
                For i = 1 To items.Count
                    ObterTodosDados.Add items(i)
                Next i
            End If

            ' Atualizar informações de paginação
            If pageData.Exists("total_pages") Then
                totalPaginas = pageData("total_pages")
            End If

            Debug.Print "Página " & pagina & " de " & totalPaginas & " processada"
            pagina = pagina + 1
        Else
            Debug.Print "Erro ao obter página " & pagina & ": " & response.StatusCode
            Exit Do
        End If
    Loop
End Function

' Paginação baseada em cursor
Function ObterDadosCursor(baseUrl As String) As Collection
    Set ObterDadosCursor = CreateJSONArray()

    Dim cursor As String
    Dim temMaisDados As Boolean

    cursor = ""
    temMaisDados = True

    Do While temMaisDados
        Dim url As String
        If Len(cursor) > 0 Then
            url = baseUrl & "?cursor=" & cursor & "&limit=50"
        Else
            url = baseUrl & "?limit=50"
        End If

        Dim response As HttpResponse
        Set response = HttpGet(url)

        If response.IsSuccess Then
            Dim pageData As Object
            Set pageData = response.Json

            ' Processar dados
            If pageData.Exists("data") Then
                Dim items As Collection
                Set items = pageData("data")

                Dim i As Integer
                For i = 1 To items.Count
                    ObterDadosCursor.Add items(i)
                Next i
            End If

            ' Verificar próxima página
            If pageData.Exists("has_more") And pageData("has_more") = True Then
                If pageData.Exists("next_cursor") Then
                    cursor = pageData("next_cursor")
                Else
                    temMaisDados = False
                End If
            Else
                temMaisDados = False
            End If
        Else
            Debug.Print "Erro na paginação: " & response.StatusCode
            temMaisDados = False
        End If
    Loop
End Function
```

### Upload de Arquivos

```vb
Function UploadArquivo(caminhoArquivo As String, url As String) As Boolean
    On Error GoTo ErroUpload

    ' Verificar se arquivo existe
    If Dir(caminhoArquivo) = "" Then
        Debug.Print "Arquivo não encontrado: " & caminhoArquivo
        UploadArquivo = False
        Exit Function
    End If

    ' Ler arquivo
    Dim fileNum As Integer
    fileNum = FreeFile

    Dim fileContent As String
    Open caminhoArquivo For Binary As #fileNum
    fileContent = Space$(LOF(fileNum))
    Get #fileNum, , fileContent
    Close #fileNum

    ' Preparar multipart data
    Dim boundary As String
    boundary = "----VB6FileUpload" & Format(Timer * 1000, "0")

    Dim fileName As String
    fileName = Mid(caminhoArquivo, InStrRev(caminhoArquivo, "\") + 1)

    Dim multipartData As String
    multipartData = "--" & boundary & vbCrLf
    multipartData = multipartData & "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & vbCrLf
    multipartData = multipartData & "Content-Type: application/octet-stream" & vbCrLf & vbCrLf
    multipartData = multipartData & fileContent & vbCrLf
    multipartData = multipartData & "--" & boundary & "--" & vbCrLf

    ' Configurar headers
    Dim headers As Dictionary
    Set headers = CreateJSONObject()
    headers.Add "Content-Type", "multipart/form-data; boundary=" & boundary

    ' Fazer upload
    Dim response As HttpResponse
    Set response = HttpPost(url, multipartData, headers)

    If response.IsSuccess Then
        Debug.Print "Upload bem-sucedido: " & fileName
        UploadArquivo = True
    Else
        Debug.Print "Erro no upload: " & response.StatusCode & " - " & response.StatusText
        UploadArquivo = False
    End If

    Exit Function

ErroUpload:
    Debug.Print "Erro no upload: " & Err.Description
    UploadArquivo = False
End Function
```

## Padrões de Uso por Tipo de API

### REST APIs Padrão

```vb
' Padrão CRUD completo
Sub ExemploRESTCompleto()
    InitializeHttpClient "https://api.exemplo.com/v1"
    SetDefaultHeader "Authorization", "Bearer " & GetToken()

    ' CREATE (POST)
    Dim novoUsuario As Dictionary
    Set novoUsuario = CreateJSONObject()
    novoUsuario.Add "name", "João Silva"
    novoUsuario.Add "email", "joao@email.com"

    Dim usuarioCriado As Object
    Set usuarioCriado = PostJson("/users", novoUsuario)

    Dim userId As Long
    userId = usuarioCriado("id")

    ' READ (GET)
    Dim usuario As Object
    Set usuario = GetJson("/users/" & userId)

    ' UPDATE (PUT)
    usuario("name") = "João Santos Silva"
    Dim usuarioAtualizado As Object
    Set usuarioAtualizado = PutJson("/users/" & userId, usuario)

    ' DELETE
    Dim response As HttpResponse
    Set response = HttpDelete("/users/" & userId)

    If response.StatusCode = 204 Then
        Debug.Print "Usuário removido com sucesso"
    End If
End Sub
```

### GraphQL APIs

```vb
' Consulta GraphQL
Function ExecutarGraphQL(query As String, Optional variables As Dictionary = Nothing) As Object
    Dim requestBody As Dictionary
    Set requestBody = CreateJSONObject()
    requestBody.Add "query", query

    If Not variables Is Nothing Then
        requestBody.Add "variables", variables
    End If

    Dim headers As Dictionary
    Set headers = CreateJSONObject()
    headers.Add "Content-Type", "application/json"

    Dim response As HttpResponse
    Set response = HttpPost("/graphql", BuildJSON(requestBody), headers)

    If response.IsSuccess Then
        Dim result As Object
        Set result = response.Json

        If result.Exists("errors") Then
            Debug.Print "Erros GraphQL: " & BuildJSON(result("errors"))
            Set ExecutarGraphQL = Nothing
        Else
            Set ExecutarGraphQL = result("data")
        End If
    Else
        Set ExecutarGraphQL = Nothing
    End If
End Function

' Exemplo de uso GraphQL
Sub ExemploGraphQL()
    InitializeHttpClient "https://api.github.com"
    SetDefaultHeader "Authorization", "Bearer " & GetGitHubToken()

    Dim query As String
    query = "query { viewer { login name email } }"

    Dim userData As Object
    Set userData = ExecutarGraphQL(query)

    If Not userData Is Nothing Then
        Debug.Print "Usuário GitHub: " & userData("viewer")("login")
    End If
End Sub
```

---

**💡 Dica de Performance**: Para APIs com rate limiting, implemente um sistema de cache local para evitar requisições desnecessárias.

**🔒 Segurança**: Sempre valide certificados SSL em produção e nunca inclua tokens de acesso no código fonte.
