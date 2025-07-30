# HttpResponse Class - Documentação Técnica

A classe `HttpResponse.cls` encapsula todas as informações de uma resposta HTTP, fornecendo uma interface limpa e intuitiva para acessar status, headers, corpo da resposta e dados JSON parseados.

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Propriedades Principais](#propriedades-principais)
- [Métodos e Funcionalidades](#métodos-e-funcionalidades)
- [Sistema de Headers](#sistema-de-headers)
- [Integração com JSON](#integração-com-json)
- [Casos de Uso Práticos](#casos-de-uso-práticos)
- [Tratamento de Erros](#tratamento-de-erros)

## Visão Geral

### Responsabilidades da Classe

Principais funcionalidades:

1. Encapsulamento de dados de resposta HTTP
2. Parse automático de JSON quando aplicável
3. Acesso estruturado a headers de resposta
4. Verificação simplificada de sucesso/erro
5. Abstração da complexidade do XMLHTTP

### Estrutura da Classe

```shell
    HttpResponse
    ├── Propriedades de Status
    │   ├── StatusCode (Long)
    │   ├── StatusText (String)
    │   └── IsSuccess (Boolean)
    ├── Propriedades de Conteúdo
    │   ├── Text (String)
    │   └── Json (Object)
    ├── Sistema de Headers
    │   ├── m_ResponseHeaders (Dictionary)
    │   └── GetHeader(String) → String
    └── Inicialização
        └── Initialize(Object)
```

## Propriedades Principais

### StatusCode

```vb
Public Property Get StatusCode() As Long
```

**Funcionalidade**: Retorna o código de status HTTP da resposta.

**Códigos HTTP Comuns**:

| Código | Categoria | Significado |
|--------|-----------|-------------|
| 200 | Sucesso | OK - Requisição bem-sucedida |
| 201 | Sucesso | Created - Recurso criado |
| 204 | Sucesso | No Content - Sucesso sem conteúdo |
| 400 | Cliente | Bad Request - Requisição inválida |
| 401 | Cliente | Unauthorized - Não autorizado |
| 404 | Cliente | Not Found - Recurso não encontrado |
| 500 | Servidor | Internal Server Error - Erro interno |

**Exemplo de Uso**:

```vb
Dim response As HttpResponse
Set response = HttpGet("https://api.exemplo.com/users/123")

Select Case response.StatusCode
    Case 200:
        Debug.Print "Usuário encontrado!"
    Case 404:
        Debug.Print "Usuário não existe"
    Case 401:
        Debug.Print "Token expirado - renovar autenticação"
    Case 500:
        Debug.Print "Erro no servidor - tentar novamente mais tarde"
    Case Else:
        Debug.Print "Status inesperado: " & response.StatusCode
End Select
```

### StatusText

```vb
Public Property Get StatusText() As String
```

**Funcionalidade**: Retorna a descrição textual do status HTTP.

**Exemplos de Status Text**:

- `200` → `"OK"`
- `404` → `"Not Found"`
- `500` → `"Internal Server Error"`

### IsSuccess

```vb
Public Property Get IsSuccess() As Boolean
```

**Lógica**: Retorna `True` para códigos de status entre 200-299 (sucessos HTTP).

**Uso Simplificado**:

```vb
Dim response As HttpResponse
Set response = HttpPost("/api/users", userData)

If response.IsSuccess Then
    Debug.Print "Usuário criado com sucesso!"
    ProcessarUsuario response.Json
Else
    Debug.Print "Erro ao criar usuário: " & response.StatusCode & " - " & response.StatusText
    TreatarErro response
End If
```

### Text

```vb
Public Property Get Text() As String
```

**Funcionalidade**: Retorna o corpo da resposta como string bruta.

**Casos de Uso**:

- Debug de respostas
- APIs que retornam texto simples
- Análise de formatos não-JSON (XML, HTML, etc.)

```vb
' Debug completo da resposta
Sub DebugResponse(response As HttpResponse)
    Debug.Print "=== RESPOSTA HTTP ==="
    Debug.Print "Status: " & response.StatusCode & " " & response.StatusText
    Debug.Print "Content-Type: " & response.GetHeader("Content-Type")
    Debug.Print "Content-Length: " & response.GetHeader("Content-Length")
    Debug.Print "Corpo da resposta:"
    Debug.Print response.Text
    Debug.Print "===================="
End Sub
```

## Métodos e Funcionalidades

### Initialize

```vb
Public Sub Initialize(ByVal xmlHttp As Object)
```

**Responsabilidade**: Inicializa a instância com dados do objeto XMLHTTP.

**Processo de Inicialização**:

Pseudocódigo do processo:

1. Extrair status code e text do XMLHTTP
2. Capturar corpo da resposta
3. Criar Dictionary para headers
4. Processar string de headers
5. Definir m_JsonData como Nothing (lazy loading)

**Exemplo de Uso Interno**:

```vb
' Usado internamente pelo HttpClient
Private Function CreateHttpResponse(ByVal request As cHttpRequest) As HttpResponse
    Dim response As New HttpResponse
    response.Initialize request  ' request é um wrapper do XMLHTTP
    Set CreateHttpResponse = response
End Function
```

## Sistema de Headers

### Estrutura Interna

```vb
Private m_ResponseHeaders As Dictionary
```

**Funcionalidade**: Armazena todos os headers de resposta em um Dictionary para acesso eficiente.

### GetHeader

```vb
Public Function GetHeader(ByVal headerName As String) As String
```

**Funcionalidade**: Retorna o valor de um header específico ou string vazia se não existir.

**Headers Comuns e Seus Usos**:

```vb
' Content-Type: Tipo do conteúdo
Dim contentType As String
contentType = response.GetHeader("Content-Type")
If InStr(contentType, "application/json") > 0 Then
    Debug.Print "Resposta é JSON"
End If

' Content-Length: Tamanho do conteúdo
Dim size As String
size = response.GetHeader("Content-Length")
Debug.Print "Tamanho da resposta: " & size & " bytes"

' Cache-Control: Política de cache
Dim cacheControl As String
cacheControl = response.GetHeader("Cache-Control")
Debug.Print "Cache: " & cacheControl

' X-Rate-Limit-*: Controle de rate limiting
Dim rateLimitRemaining As String
rateLimitRemaining = response.GetHeader("X-Rate-Limit-Remaining")
If Len(rateLimitRemaining) > 0 Then
    Debug.Print "Requisições restantes: " & rateLimitRemaining
End If

' ETag: Versionamento de recursos
Dim etag As String
etag = response.GetHeader("ETag")
If Len(etag) > 0 Then
    Debug.Print "ETag do recurso: " & etag
End If
```

### ParseResponseHeaders (Método Privado)

```vb
Private Sub ParseResponseHeaders(ByVal headersText As String)
```

**Algoritmo de Parsing**:

Processo de análise dos headers:

1. Dividir string por vbCrLf (quebras de linha)
2. Para cada linha:
   a. Encontrar posição dos dois pontos ':'
   b. Extrair nome (antes dos ':')
   c. Extrair valor (depois dos ':')
   d. Adicionar ao Dictionary (se não existir)

**Exemplo de String de Headers**:

``` http
HTTP/1.1 200 OK
Content-Type: application/json; charset=utf-8
Content-Length: 1234
Cache-Control: no-cache
X-Rate-Limit-Limit: 1000
X-Rate-Limit-Remaining: 999
ETag: "abc123def456"
Date: Mon, 01 Jan 2024 12:00:00 GMT
```

## Integração com JSON

### Propriedade Json

```vb
Public Property Get Json() As Object
```

**Funcionalidade**: Lazy loading de parsing JSON - só processa quando acessado.

**Algoritmo de Lazy Loading**:

```vb
' Pseudocódigo:
If m_JsonData Is Nothing Then
    If Len(m_ResponseText) > 0 Then
        Set m_JsonData = ParseJSON(m_ResponseText)
    End If
End If
Set Json = m_JsonData
```

**Vantagens do Lazy Loading**:

- Performance: só processa JSON quando necessário
- Memória: evita parsing desnecessário
- Flexibilidade: permite acesso tanto ao texto bruto quanto ao JSON

**Exemplos de Uso**:

```vb
' Uso básico
Dim response As HttpResponse
Set response = HttpGet("/api/user/123")

If response.IsSuccess Then
    Dim user As Object
    Set user = response.Json
    Debug.Print "Nome: " & user("name")
    Debug.Print "Email: " & user("email")
End If

' Verificação defensiva
If response.IsSuccess Then
    ' Verificar se é JSON válido antes de usar
    Dim contentType As String
    contentType = response.GetHeader("Content-Type")

    If InStr(contentType, "application/json") > 0 Then
        Dim data As Object
        Set data = response.Json

        If Not data Is Nothing Then
            ProcessarDadosJSON data
        Else
            Debug.Print "JSON inválido na resposta"
        End If
    Else
        Debug.Print "Resposta não é JSON: " & contentType
    End If
End If
```

## Casos de Uso Práticos

### Análise de Resposta Completa

```vb
Sub AnalisarRespostaCompleta(response As HttpResponse)
    Debug.Print "=== ANÁLISE DE RESPOSTA ==="
    Debug.Print "Status: " & response.StatusCode & " (" & response.StatusText & ")"

    ' Verificar sucesso
    If response.IsSuccess Then
        Debug.Print "✓ Requisição bem-sucedida"
    Else
        Debug.Print "✗ Requisição falhou"
    End If

    ' Analisar tipo de conteúdo
    Dim contentType As String
    contentType = response.GetHeader("Content-Type")
    Debug.Print "Tipo de conteúdo: " & contentType

    ' Tamanho da resposta
    Dim contentLength As String
    contentLength = response.GetHeader("Content-Length")
    If Len(contentLength) > 0 Then
        Debug.Print "Tamanho: " & contentLength & " bytes"
    Else
        Debug.Print "Tamanho: " & Len(response.Text) & " caracteres"
    End If

    ' Processamento específico por tipo
    If InStr(contentType, "application/json") > 0 Then
        Debug.Print "Processando como JSON..."
        Dim jsonData As Object
        Set jsonData = response.Json

        If Not jsonData Is Nothing Then
            Debug.Print "JSON parseado com sucesso"
            If TypeName(jsonData) = "Dictionary" Then
                Debug.Print "Tipo: Objeto JSON (" & jsonData.Count & " propriedades)"
            ElseIf TypeName(jsonData) = "Collection" Then
                Debug.Print "Tipo: Array JSON (" & jsonData.Count & " elementos)"
            End If
        Else
            Debug.Print "Erro ao parsear JSON"
        End If
    Else
        Debug.Print "Conteúdo texto (primeiros 100 chars):"
        Debug.Print Left(response.Text, 100)
    End If

    Debug.Print "=========================="
End Sub
```

### Sistema de Cache Baseado em ETag

```vb
Private m_CacheEtag As Dictionary  ' Cache global de ETags

Sub InitializeCache()
    Set m_CacheEtag = CreateJSONObject()
End Sub

Function GetWithCache(url As String) As Object
    ' Verificar se temos ETag em cache
    Dim cachedEtag As String
    If m_CacheEtag.Exists(url) Then
        cachedEtag = m_CacheEtag(url)
    End If

    ' Fazer requisição com If-None-Match
    Dim headers As Dictionary
    Set headers = CreateJSONObject()
    If Len(cachedEtag) > 0 Then
        headers.Add "If-None-Match", cachedEtag
    End If

    Dim response As HttpResponse
    Set response = HttpGet(url, headers)

    If response.StatusCode = 304 Then
        ' Não modificado - usar cache
        Debug.Print "Usando dados em cache para: " & url
        Set GetWithCache = GetCachedData(url)
    ElseIf response.IsSuccess Then
        ' Dados novos - atualizar cache
        Dim newEtag As String
        newEtag = response.GetHeader("ETag")

        If Len(newEtag) > 0 Then
            m_CacheEtag(url) = newEtag
        End If

        Set GetWithCache = response.Json
        CacheData url, response.Json
    Else
        Set GetWithCache = Nothing
    End If
End Function
```

### Rate Limiting

```vb
Sub CheckRateLimit(response As HttpResponse)
    Dim rateLimit As String
    Dim rateLimitRemaining As String
    Dim rateLimitReset As String

    rateLimit = response.GetHeader("X-Rate-Limit-Limit")
    rateLimitRemaining = response.GetHeader("X-Rate-Limit-Remaining")
    rateLimitReset = response.GetHeader("X-Rate-Limit-Reset")

    If Len(rateLimit) > 0 Then
        Debug.Print "Rate Limit: " & rateLimitRemaining & "/" & rateLimit

        Dim remaining As Long
        remaining = CLng(rateLimitRemaining)

        If remaining < 10 Then
            Debug.Print "⚠️ AVISO: Poucas requisições restantes!"

            If Len(rateLimitReset) > 0 Then
                Dim resetTime As Date
                resetTime = DateAdd("s", CLng(rateLimitReset), #1/1/1970#)
                Debug.Print "Reset em: " & resetTime
            End If
        End If

        If remaining = 0 Then
            Debug.Print "🚫 Rate limit atingido!"
            ' Implementar pausa ou retry
        End If
    End If
End Sub
```

## Tratamento de Erros

### Validação de Resposta

```vb
Function ValidateResponse(response As HttpResponse) As Boolean
    ' Verificar se response não é Nothing
    If response Is Nothing Then
        Debug.Print "Erro: Response é Nothing"
        ValidateResponse = False
        Exit Function
    End If

    ' Verificar status code
    If response.StatusCode = 0 Then
        Debug.Print "Erro: Status code inválido (possível erro de rede)"
        ValidateResponse = False
        Exit Function
    End If

    ' Verificar se resposta foi bem-sucedida
    If Not response.IsSuccess Then
        Debug.Print "Erro HTTP: " & response.StatusCode & " - " & response.StatusText
        LogHttpError response
        ValidateResponse = False
        Exit Function
    End If

    ValidateResponse = True
End Function

Sub LogHttpError(response As HttpResponse)
    Debug.Print "=== ERRO HTTP ==="
    Debug.Print "Status: " & response.StatusCode & " " & response.StatusText
    Debug.Print "Content-Type: " & response.GetHeader("Content-Type")

    ' Log do corpo da resposta para debugging
    If Len(response.Text) > 0 Then
        Debug.Print "Resposta de erro:"
        Debug.Print Left(response.Text, 500)  ' Primeiros 500 caracteres
    End If

    Debug.Print "================="
End Sub
```

### Tratamento de JSON Inválido

```vb
Function SafeGetJson(response As HttpResponse) As Object
    On Error GoTo ErrorHandler

    If Not response.IsSuccess Then
        Set SafeGetJson = Nothing
        Exit Function
    End If

    ' Verificar content-type
    Dim contentType As String
    contentType = response.GetHeader("Content-Type")

    If InStr(LCase(contentType), "json") = 0 Then
        Debug.Print "Aviso: Content-Type não indica JSON: " & contentType
    End If

    ' Tentar parsing
    Set SafeGetJson = response.Json
    Exit Function

ErrorHandler:
    Debug.Print "Erro ao parsear JSON: " & Err.Description
    Debug.Print "Resposta bruta: " & Left(response.Text, 200)
    Set SafeGetJson = Nothing
End Function
```

---

**🎯 Dica de Performance**: A classe HttpResponse usa lazy loading para o parsing JSON. Isso significa que você pode acessar `.Text` sem custo de processamento, e só pagar pelo parsing quando acessar `.Json`.

**🔒 Segurança**: Sempre valide o Content-Type antes de processar como JSON, especialmente ao trabalhar com APIs externas que podem retornar formatos inesperados.
