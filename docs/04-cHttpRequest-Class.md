# cHttpRequest Class - Documentação Técnica

A classe `cHttpRequest.cls` é um wrapper para o objeto XMLHTTP do Windows, fornecendo uma interface simplificada e consistente para execução de requisições HTTP/HTTPS.

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Arquitetura da Classe](#arquitetura-da-classe)
- [Métodos Principais](#métodos-principais)
- [Ciclo de Vida da Requisição](#ciclo-de-vida-da-requisição)
- [Configurações e Propriedades](#configurações-e-propriedades)
- [Integração com Sistema](#integração-com-sistema)
- [Tratamento de Erros](#tratamento-de-erros)

## Visão Geral

### Responsabilidades da Classe

Principais funcionalidades:

1. Wrapper simplificado para XMLHTTP
2. Configuração padronizada de requisições HTTP
3. Gerenciamento de headers de requisição
4. Controle de timeout e estados de conexão
5. Abstração das complexidades do XMLHTTP nativo

### Vantagens do Wrapper

- **Simplificação**: Interface mais limpa que o XMLHTTP nativo
- **Consistência**: Comportamento padronizado entre diferentes versões do Windows
- **Timeout**: Controle de timeout integrado
- **Debug**: Facilita debugging e logging de requisições

## Arquitetura da Classe

### Estrutura Interna

``` shell
cHttpRequest
├── Variáveis Privadas
│   ├── m_XmlHttp (Object)         # Instância XMLHTTP
│   └── m_Timeout (Long)           # Timeout configurado
├── Métodos de Configuração
│   ├── Open_()                    # Configurar método e URL
│   ├── SetRequestHeader()         # Definir headers
│   └── SetTimeout()               # Configurar timeout
├── Execução
│   └── Send()                     # Enviar requisição
└── Propriedades de Resposta
    ├── status                     # Código HTTP
    ├── statusText                 # Texto do status
    ├── responseText               # Corpo da resposta
    ├── getAllResponseHeaders()    # Headers da resposta
    └── readyState                 # Estado da requisição
```

### Inicialização (Class_Initialize)

```vb
Private Sub Class_Initialize()
    ' Cria instância do XMLHTTP
    Set m_XmlHttp = CreateObject("MSXML2.XMLHTTP")
    m_Timeout = 30000 ' 30 segundos padrão
End Sub
```

**Versões do XMLHTTP Suportadas**:

- `MSXML2.XMLHTTP.6.0` (Windows Vista+, preferencial)
- `MSXML2.XMLHTTP.3.0` (Windows XP+, fallback)
- `MSXML2.XMLHTTP` (Versão genérica)

## Métodos Principais

### Open_

```vb
Public Sub Open_(ByVal method As String, ByVal url As String, Optional ByVal async As Boolean = False)
```

**Funcionalidade**: Configura o método HTTP, URL e modo de operação (síncrono/assíncrono).

**Parâmetros Detalhados**:

- `method`: Método HTTP (GET, POST, PUT, DELETE, PATCH, HEAD, OPTIONS)
- `url`: URL completa da requisição
- `async`: Modo assíncrono (padrão: False para simplicidade)

**Métodos HTTP Suportados**:

```vb
' Métodos padrão
request.Open_ "GET", "https://api.exemplo.com/users"
request.Open_ "POST", "https://api.exemplo.com/users"
request.Open_ "PUT", "https://api.exemplo.com/users/123"
request.Open_ "DELETE", "https://api.exemplo.com/users/123"
request.Open_ "PATCH", "https://api.exemplo.com/users/123"

' Métodos menos comuns
request.Open_ "HEAD", "https://api.exemplo.com/status"
request.Open_ "OPTIONS", "https://api.exemplo.com/users"
```

**Exemplo de Uso Completo**:

```vb
Dim request As New cHttpRequest
request.Open_ "POST", "https://api.github.com/user/repos"
request.SetRequestHeader "Authorization", "Bearer ghp_xxxxxxxxxxxx"
request.SetRequestHeader "Content-Type", "application/json"
request.SetRequestHeader "Accept", "application/vnd.github.v3+json"
request.Send "{""name"":""meu-novo-repo"",""private"":false}"

If request.status = 201 Then
    Debug.Print "Repositório criado: " & request.responseText
End If
```

### SetRequestHeader

```vb
Public Sub SetRequestHeader(ByVal headerName As String, ByVal headerValue As String)
```

**Funcionalidade**: Define headers HTTP que serão enviados com a requisição.

**Headers Comuns e Casos de Uso**:

```vb
' Autenticação
request.SetRequestHeader "Authorization", "Bearer " & token
request.SetRequestHeader "Authorization", "Basic " & Base64Encode(username & ":" & password)
request.SetRequestHeader "X-API-Key", apiKey

' Tipo de conteúdo
request.SetRequestHeader "Content-Type", "application/json"
request.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
request.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary

' Aceitação de resposta
request.SetRequestHeader "Accept", "application/json"
request.SetRequestHeader "Accept", "application/xml"
request.SetRequestHeader "Accept-Language", "pt-BR,pt;q=0.9,en;q=0.8"

' Headers customizados da API
request.SetRequestHeader "X-API-Version", "v2"
request.SetRequestHeader "X-Client-ID", "VB6-App"
request.SetRequestHeader "X-Request-ID", GenerateUUID()

' Cache e condicional
request.SetRequestHeader "Cache-Control", "no-cache"
request.SetRequestHeader "If-None-Match", etag
request.SetRequestHeader "If-Modified-Since", lastModified
```

### Send

```vb
Public Sub Send(Optional ByVal body As Variant)
```

**Funcionalidade**: Executa a requisição HTTP, opcionalmente com corpo de dados.

**Tipos de Body Suportados**:

```vb
' Sem body (GET, DELETE)
request.Send

' String JSON
request.Send "{""name"":""João"",""email"":""joao@email.com""}"

' Form data
request.Send "name=João&email=joao@email.com"

' XML
request.Send "<?xml version=""1.0""?><user><name>João</name></user>"

' Dados binários (limitado no VB6)
request.Send binaryData
```

**Comportamento Síncrono vs Assíncrono**:

```vb
' Modo síncrono (padrão) - bloqueia até completar
request.Open_ "GET", url, False
request.Send
' Código aqui só executa após resposta completa

' Modo assíncrono (avançado) - não bloqueia
request.Open_ "GET", url, True
request.Send
' Precisaria implementar verificação de readyState
Do While request.readyState <> 4
    DoEvents
    Sleep 10  ' Evitar consumo excessivo de CPU
Loop
```

### SetTimeout

```vb
Public Sub SetTimeout(ByVal timeoutMs As Long)
```

**Funcionalidade**: Define timeout para a requisição em milissegundos.

**Configurações Típicas**:

```vb
' Requisições rápidas (APIs locais)
request.SetTimeout 5000   ' 5 segundos

' Requisições normais
request.SetTimeout 30000  ' 30 segundos (padrão)

' Upload de arquivos ou operações longas
request.SetTimeout 300000 ' 5 minutos

' APIs lentas ou instáveis
request.SetTimeout 60000  ' 1 minuto
```

**Limitações no VB6/XMLHTTP**:

>*Nota*:
> Nem todas as versões do XMLHTTP suportam timeout
> A implementação pode variar entre versões do Windows
> Para timeout mais robusto, considere implementar usando Timer

## Ciclo de Vida da Requisição

### Estados do XMLHTTP (readyState)

```vb
' Estados possíveis:
Const XMLHTTP_UNINITIALIZED = 0  ' Não inicializado
Const XMLHTTP_LOADING = 1        ' Carregando
Const XMLHTTP_LOADED = 2         ' Carregado
Const XMLHTTP_INTERACTIVE = 3    ' Interativo
Const XMLHTTP_COMPLETE = 4       ' Completo
```

### Fluxo Completo de Requisição

```vb
Sub ExemploFluxoCompleto()
    Dim request As New cHttpRequest

    ' 1. Configuração inicial
    request.SetTimeout 15000

    ' 2. Abertura da conexão
    request.Open_ "POST", "https://api.exemplo.com/data"

    ' 3. Configuração de headers
    request.SetRequestHeader "Content-Type", "application/json"
    request.SetRequestHeader "Authorization", "Bearer " & GetToken()

    ' 4. Envio da requisição
    Dim jsonData As String
    jsonData = "{""action"":""create"",""data"":{""name"":""Teste""}}"
    request.Send jsonData

    ' 5. Verificação do resultado
    Debug.Print "ReadyState: " & request.readyState  ' Deve ser 4
    Debug.Print "Status: " & request.status          ' Ex: 200, 201, 404...
    Debug.Print "StatusText: " & request.statusText  ' Ex: "OK", "Created"...

    ' 6. Processamento da resposta
    If request.status >= 200 And request.status <= 299 Then
        Debug.Print "Sucesso: " & request.responseText
    Else
        Debug.Print "Erro: " & request.status & " - " & request.statusText
    End If
End Sub
```

## Configurações e Propriedades

### Propriedades de Resposta

#### status

```vb
Public Property Get status() As Long
```

**Códigos de Status por Categoria**:

``` h
//1xx: Informacionais (raros em APIs REST)
100 Continue
101 Switching Protocols

// 2xx: Sucesso
200 OK - Requisição bem-sucedida
201 Created - Recurso criado
202 Accepted - Aceito para processamento
204 No Content - Sucesso sem conteúdo

// 3xx: Redirecionamento (geralmente tratado automaticamente)
301 Moved Permanently
302 Found
304 Not Modified

// 4xx: Erro do cliente
400 Bad Request - Requisição inválida
401 Unauthorized - Não autorizado
403 Forbidden - Proibido
404 Not Found - Não encontrado
422 Unprocessable Entity - Dados inválidos
429 Too Many Requests - Rate limit

// 5xx: Erro do servidor
500 Internal Server Error - Erro interno
502 Bad Gateway - Gateway inválido
503 Service Unavailable - Serviço indisponível
```

#### responseText

```vb
Public Property Get responseText() As String
```

**Limitações e Considerações**:

Encoding de caracteres

- XMLHTTP geralmente lida bem com UTF-8
- Caracteres especiais são preservados
- Para encoding específico, verificar Content-Type header

Tamanho da resposta

- VB6 String pode lidar com ~2GB teoricamente
- Na prática, limitado pela memória disponível
- Para arquivos grandes, considere streaming

#### getAllResponseHeaders

```vb
Public Function getAllResponseHeaders() As String
```

**Formato da Resposta**:

``` http
HTTP/1.1 200 OK
Content-Type: application/json; charset=utf-8
Content-Length: 1234
Cache-Control: no-cache
Date: Mon, 01 Jan 2024 12:00:00 GMT
Server: nginx/1.18.0
X-Rate-Limit-Limit: 1000
X-Rate-Limit-Remaining: 999
```

**Parsing Manual**:

```vb
Function ParseHeader(allHeaders As String, headerName As String) As String
    Dim lines() As String
    Dim i As Integer
    Dim colonPos As Integer

    lines = Split(allHeaders, vbCrLf)

    For i = 0 To UBound(lines)
        colonPos = InStr(lines(i), ":")
        If colonPos > 0 Then
            If Trim(Left(lines(i), colonPos - 1)) = headerName Then
                ParseHeader = Trim(Mid(lines(i), colonPos + 1))
                Exit Function
            End If
        End If
    Next i

    ParseHeader = ""
End Function
```

## Integração com Sistema

### Uso pelo HttpClient

```vb
' Como o HttpClient usa internamente:
Private Function ExecuteRequest(method As String, url As String, body As String, headers As Dictionary) As HttpResponse
    Dim req As New cHttpRequest
    Dim key As Variant

    With req
        .Open_ method, url, False
        .SetTimeout config.timeout

        ' Aplicar headers
        For Each key In headers.Keys
            .SetRequestHeader CStr(key), CStr(headers(key))
        Next key

        ' Enviar com ou sem body
        If Len(body) > 0 Then
            .Send body
        Else
            .Send
        End If
    End With

    ' Criar HttpResponse
    Set ExecuteRequest = CreateHttpResponse(req)
End Function
```

### Logging e Debug

```vb
Sub LogRequest(request As cHttpRequest, method As String, url As String, body As String)
    Debug.Print "=== HTTP REQUEST ==="
    Debug.Print "Method: " & method
    Debug.Print "URL: " & url
    Debug.Print "Timeout: " & request.m_Timeout & "ms"

    If Len(body) > 0 Then
        Debug.Print "Body: " & Left(body, 200)  ' Primeiros 200 chars
        If Len(body) > 200 Then
            Debug.Print "... (truncated)"
        End If
    End If

    Debug.Print "==================="
End Sub

Sub LogResponse(request As cHttpRequest)
    Debug.Print "=== HTTP RESPONSE ==="
    Debug.Print "Status: " & request.status & " " & request.statusText
    Debug.Print "Headers:"
    Debug.Print request.getAllResponseHeaders()
    Debug.Print "Body: " & Left(request.responseText, 500)
    Debug.Print "====================="
End Sub
```

## Tratamento de Erros

### Erros Comuns e Soluções

```vb
Function SafeExecuteRequest() As Boolean
    On Error GoTo ErrorHandler

    Dim request As New cHttpRequest
    request.Open_ "GET", "https://api.exemplo.com/data"
    request.Send

    SafeExecuteRequest = True
    Exit Function

ErrorHandler:
    Select Case Err.Number
        Case -2147012894:  ' 0x800C0002 - Timeout
            Debug.Print "Erro: Timeout na requisição"

        Case -2147012867:  ' 0x800C001D - ERROR_INTERNET_CANNOT_CONNECT
            Debug.Print "Erro: Não foi possível conectar ao servidor"

        Case -2147012889:  ' 0x800C0007 - Name not resolved
            Debug.Print "Erro: Nome do servidor não encontrado"

        Case -2147012866:  ' 0x800C001E - Connection reset
            Debug.Print "Erro: Conexão resetada pelo servidor"

        Case Else:
            Debug.Print "Erro HTTP não tratado: " & Err.Number & " - " & Err.Description
    End Select

    SafeExecuteRequest = False
End Function
```

### Retry com Backoff

```vb
Function RequestWithRetry(url As String, maxRetries As Integer) As cHttpRequest
    Dim attempt As Integer
    Dim request As cHttpRequest
    Dim waitTime As Long

    For attempt = 1 To maxRetries
        Set request = New cHttpRequest

        On Error GoTo RetryHandler

        request.Open_ "GET", url
        request.SetTimeout 30000
        request.Send

        ' Se chegou aqui, sucesso
        Set RequestWithRetry = request
        Exit Function

RetryHandler:
        Debug.Print "Tentativa " & attempt & " falhou: " & Err.Description

        If attempt < maxRetries Then
            ' Backoff exponencial: 1s, 2s, 4s, 8s...
            waitTime = 1000 * (2 ^ (attempt - 1))
            Debug.Print "Aguardando " & waitTime & "ms antes da próxima tentativa..."
            Sleep waitTime
        End If

        On Error GoTo 0
    Next attempt

    ' Todas as tentativas falharam
    Set RequestWithRetry = Nothing
End Function
```

### Validação de URL

```vb
Function IsValidUrl(url As String) As Boolean
    IsValidUrl = False

    ' Verificações básicas
    If Len(url) = 0 Then Exit Function
    If InStr(url, " ") > 0 Then Exit Function

    ' Protocolo válido
    If Not (Left(LCase(url), 7) = "http://" Or Left(LCase(url), 8) = "https://") Then
        Exit Function
    End If

    ' Deve ter pelo menos um ponto (domínio)
    If InStr(Mid(url, 9), ".") = 0 Then Exit Function

    IsValidUrl = True
End Function
```

---

**🔧 Nota Técnica**: A classe cHttpRequest é uma abstração fina sobre o XMLHTTP nativo. Para casos avançados que requerem controle total sobre a requisição, você ainda pode acessar o objeto XMLHTTP interno através de `m_XmlHttp`.

**⚡ Performance**: O modo síncrono é adequado para a maioria dos casos de uso. O modo assíncrono requer gerenciamento manual de estado e pode complicar o código sem benefícios significativos em aplicações desktop VB6.
