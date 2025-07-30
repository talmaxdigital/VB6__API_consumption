# HttpClient Module - Documentação Técnica

O `HttpClient.bas` é o módulo principal do sistema, responsável por todas as operações HTTP e pela coordenação entre os demais componentes.

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Configuração e Inicialização](#configuração-e-inicialização)
- [Métodos HTTP Básicos](#métodos-http-básicos)
- [Métodos Especializados JSON](#métodos-especializados-json)
- [Sistema de Headers](#sistema-de-headers)
- [Utilitários](#utilitários)
- [Estruturas Internas](#estruturas-internas)

## Visão Geral

### Responsabilidades do Módulo

Principais responsabilidades:

1. Configuração global do cliente HTTP
2. Execução de requisições HTTP (GET, POST, PUT, DELETE, PATCH)
3. Gerenciamento de headers padrão e customizados
4. Integração transparente com JSON
5. Tratamento de URLs (absolutas e relativas)
6. Utilitários para encoding e query strings

### Arquitetura Interna

``` shell
HttpClient Module
├── HTTP_CONFIG (Type)           # Configurações globais
├── InitializeHttpClient()       # Configuração inicial
├── HTTP Methods                 # GET, POST, PUT, DELETE, PATCH
│   ├── HttpGet()
│   ├── HttpPost()
│   ├── HttpPut()
│   ├── HttpDelete()
│   └── HttpPatch()
├── JSON Methods                 # Métodos especializados
│   ├── GetJson()
│   ├── PostJson()
│   └── PutJson()
├── Header Management            # Gerenciamento de headers
│   ├── SetDefaultHeader()
│   └── RemoveDefaultHeader()
├── File Operations              # Upload/Download
│   ├── DownloadFile()
│   └── UploadFile()
└── Utilities                    # Utilitários
    ├── UrlEncode()
    └── BuildQueryString()
```

## Configuração e Inicialização

### InitializeHttpClient

```vb
Public Sub InitializeHttpClient(Optional ByVal baseUrl As String = "", _
                               Optional ByVal timeout As Long = 30000, _
                               Optional ByVal userAgent As String = "VB6-HttpClient/1.0")
```

**Propósito**: Inicializa o cliente HTTP com configurações globais que serão aplicadas a todas as requisições subsequentes.

**Parâmetros**:

- `baseUrl` (String, opcional): URL base para requisições relativas
- `timeout` (Long, opcional): Timeout em milissegundos (padrão: 30000)
- `userAgent` (String, opcional): User-Agent para identificação

**Funcionamento Interno**:

```vb
' Exemplo de uso básico
InitializeHttpClient "https://api.github.com", 15000, "MeuApp/1.0"

' Configuração para API local
InitializeHttpClient "http://localhost:3000/api", 5000

' Configuração mínima (apenas timeout)
InitializeHttpClient "", 60000
```

**Configurações Aplicadas**:

```vb
' Headers padrão configurados automaticamente:
config.DefaultHeaders.Add "User-Agent", userAgent
config.DefaultHeaders.Add "Accept", "application/json"
config.DefaultHeaders.Add "Accept-Encoding", "gzip, deflate"
```

### Estrutura HTTP_CONFIG

```vb
Private Type HTTP_CONFIG
    baseUrl As String           ' URL base para requisições relativas
    DefaultHeaders As Dictionary ' Headers aplicados a todas as requisições
    timeout As Long             ' Timeout padrão em milissegundos
    userAgent As String         ' User-Agent padrão
    AcceptEncoding As String    ' Encodings aceitos
End Type
```

**Exemplo de Configuração Completa**:

```vb
Sub ConfigurarClienteCompleto()
    ' Inicialização base
    InitializeHttpClient "https://api.exemplo.com/v1", 20000, "MinhaApp/2.1"

    ' Headers globais
    SetDefaultHeader "Authorization", "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9..."
    SetDefaultHeader "X-API-Version", "2.1"
    SetDefaultHeader "X-Client-ID", "MinhaEmpresa"

    ' Agora todas as requisições usarão essas configurações
End Sub
```

## Métodos HTTP Básicos

### HttpGet

```vb
Public Function HttpGet(ByVal url As String, _
                       Optional ByVal customHeaders As Dictionary = Nothing, _
                       Optional ByVal body As String = "") As HttpResponse
```

**Características Especiais**:

- Suporte a body em requisições GET (necessário para algumas APIs)
- Merge automático de headers padrão com customizados
- Tratamento de URLs relativas e absolutas

**Exemplos Avançados**:

```vb
' GET simples
Dim response As HttpResponse
Set response = HttpGet("https://api.github.com/users/octocat")

' GET com headers customizados
Dim headers As Dictionary
Set headers = CreateJSONObject()
headers.Add "X-Custom-Header", "valor"
Set response = HttpGet("/endpoint", headers)

' GET com body (usado por APIs como TomTicket)
Dim params As String
params = "{""customer_id"":""12345"",""type"":""premium""}"
Set response = HttpGet("/customer/check", Nothing, params)
```

### HttpPost

```vb
Public Function HttpPost(ByVal url As String, _
                        ByVal body As String, _
                        Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
```

**Casos de Uso Típicos**:

```vb
' POST com JSON
Dim jsonData As String
jsonData = "{""name"":""João"",""email"":""joao@email.com""}"
Set response = HttpPost("/users", jsonData)

' POST com form data
Dim formData As String
formData = "name=João&email=joao@email.com"
Dim headers As Dictionary
Set headers = CreateJSONObject()
headers.Add "Content-Type", "application/x-www-form-urlencoded"
Set response = HttpPost("/form-endpoint", formData, headers)
```

### HttpPut, HttpDelete, HttpPatch

Seguem o mesmo padrão do HttpPost, adaptados para seus respectivos métodos HTTP.

```vb
' PUT para atualização completa
Set response = HttpPut("/users/123", jsonData)

' PATCH para atualização parcial
Set response = HttpPatch("/users/123", "{""email"":""novo@email.com""}")

' DELETE
Set response = HttpDelete("/users/123")
```

## Métodos Especializados JSON

### GetJson

```vb
Public Function GetJson(ByVal url As String, _
                       Optional ByVal customHeaders As Dictionary = Nothing, _
                       Optional ByVal bodyParams As Dictionary = Nothing) As Object
```

**Funcionalidade Avançada**: Combina requisição GET com parsing automático de JSON.

**Vantagens**:

- Parsing automático da resposta
- Tratamento de erros integrado
- Suporte a parâmetros no body (Dictionary → JSON)

**Exemplos Práticos**:

```vb
' GET simples com parsing automático
Dim user As Object
Set user = GetJson("https://api.github.com/users/octocat")
Debug.Print "Nome: " & user("name")
Debug.Print "Empresa: " & user("company")

' GET com parâmetros complexos no body
Dim params As Dictionary
Set params = CreateJSONObject()
params.Add "filters", CreateJSONObject()
params("filters").Add "status", "active"
params("filters").Add "date_from", "2024-01-01"

Dim results As Object
Set results = GetJson("/api/reports", Nothing, params)
```

### PostJson

```vb
Public Function PostJson(ByVal url As String, _
                        ByVal jsonObject As Object, _
                        Optional ByVal customHeaders As Dictionary = Nothing) As Object
```

**Fluxo Completo**:

1. Converte objeto VB6 para JSON string
2. Configura Content-Type automaticamente
3. Executa POST
4. Faz parsing da resposta
5. Retorna objeto VB6

**Exemplo Avançado**:

```vb
' Criar estrutura complexa
Dim produto As Dictionary
Set produto = CreateJSONObject()
produto.Add "nome", "Smartphone"
produto.Add "preco", 899.99
produto.Add "especificacoes", CreateJSONObject()
produto("especificacoes").Add "memoria", "128GB"
produto("especificacoes").Add "cor", "Preto"

Dim categorias As Collection
Set categorias = CreateJSONArray()
categorias.Add "eletrônicos"
categorias.Add "smartphones"
produto.Add "categorias", categorias

' Enviar e receber resposta
Dim produtoCriado As Object
Set produtoCriado = PostJson("/produtos", produto)

If Not produtoCriado Is Nothing Then
    Debug.Print "Produto criado com ID: " & produtoCriado("id")
    Debug.Print "Status: " & produtoCriado("status")
End If
```

### PutJson

Similar ao PostJson, mas usando método PUT para atualizações.

```vb
' Atualização de dados
Dim dadosAtualizacao As Dictionary
Set dadosAtualizacao = CreateJSONObject()
dadosAtualizacao.Add "nome", "Nome Atualizado"
dadosAtualizacao.Add "email", "novo@email.com"

Dim usuarioAtualizado As Object
Set usuarioAtualizado = PutJson("/users/123", dadosAtualizacao)
```

## Sistema de Headers

### SetDefaultHeader

```vb
Public Sub SetDefaultHeader(ByVal headerName As String, ByVal headerValue As String)
```

**Funcionalidade**: Define headers que serão aplicados automaticamente a todas as requisições.

**Casos de Uso**:

```vb
' Autenticação global
SetDefaultHeader "Authorization", "Bearer " & GetCurrentToken()

' Versionamento da API
SetDefaultHeader "API-Version", "v2"

' Identificação do cliente
SetDefaultHeader "X-Client-ID", "VB6-App"
SetDefaultHeader "X-Client-Version", "1.0.0"

' Headers de segurança
SetDefaultHeader "X-Requested-With", "XMLHttpRequest"
```

### RemoveDefaultHeader

```vb
Public Sub RemoveDefaultHeader(ByVal headerName As String)
```

**Exemplo de Rotação de Token**:

```vb
Sub RenovarAutenticacao()
    ' Remove token antigo
    RemoveDefaultHeader "Authorization"

    ' Obtém novo token
    Dim novoToken As String
    novoToken = ObterNovoToken()

    ' Define novo token
    SetDefaultHeader "Authorization", "Bearer " & novoToken
End Sub
```

## Utilitários

### UrlEncode

```vb
Public Function UrlEncode(ByVal text As String) As String
```

**Implementação Completa**: Codifica caracteres especiais seguindo o padrão RFC 3986.

**Caracteres Preservados**: A-Z, a-z, 0-9, -, _, ., ~
**Caracteres Codificados**: Todos os demais são convertidos para %XX

**Exemplos**:

```vb
Debug.Print UrlEncode("João & Maria")     ' Output: Jo%C3%A3o%20%26%20Maria
Debug.Print UrlEncode("user@domain.com")  ' Output: user%40domain.com
Debug.Print UrlEncode("100% correto")     ' Output: 100%25%20correto
```

### BuildQueryString

```vb
Public Function BuildQueryString(ByVal params As Dictionary) As String
```

**Funcionalidade**: Constrói query strings a partir de Dictionary, com encoding automático.

**Exemplo Completo**:

```vb
Dim filtros As Dictionary
Set filtros = CreateJSONObject()
filtros.Add "nome", "João Silva"
filtros.Add "idade", "30"
filtros.Add "cidade", "São Paulo"
filtros.Add "ativo", "true"

Dim queryString As String
queryString = BuildQueryString(filtros)
' Output: nome=Jo%C3%A3o%20Silva&idade=30&cidade=S%C3%A3o%20Paulo&ativo=true

Dim urlCompleta As String
urlCompleta = "https://api.exemplo.com/users?" & queryString
```

## Estruturas Internas

### ExecuteRequest (Função Privada)

```vb
Private Function ExecuteRequest(ByVal method As String, _
                               ByVal url As String, _
                               ByVal body As String, _
                               Optional ByVal customHeaders As Dictionary = Nothing) As HttpResponse
```

**Responsabilidades**:

1. Construção de URL completa (baseUrl + URL relativa)
2. Merge de headers padrão com customizados
3. Criação e configuração do objeto cHttpRequest
4. Aplicação de timeout
5. Execução da requisição
6. Criação do objeto HttpResponse

### MergeHeaders (Função Privada)

```vb
Private Function MergeHeaders(Optional ByVal customHeaders As Dictionary = Nothing) As Dictionary
```

**Lógica de Merge**:

1. Copia todos os headers padrão
2. Sobrescreve com headers customizados (se existirem)
3. Headers customizados têm prioridade

### BuildFullUrl (Função Privada)

```vb
Private Function BuildFullUrl(ByVal url As String) As String
```

**Lógica de Construção**:

```vb
' URL absoluta: retorna como está
"https://api.exemplo.com/users" → "https://api.exemplo.com/users"

' URL relativa com baseUrl configurada:
baseUrl = "https://api.exemplo.com"
"/users" → "https://api.exemplo.com/users"
"users" → "https://api.exemplo.com/users"

' Tratamento de barras:
baseUrl = "https://api.exemplo.com/"
"/users" → "https://api.exemplo.com/users" (remove barra dupla)
```

## Padrões de Uso Recomendados

### Inicialização da Aplicação

```vb
Sub Application_Initialize()
    ' Configuração base
    InitializeHttpClient "https://minha-api.com/v1", 30000, "MinhaApp/1.0"

    ' Headers globais
    SetDefaultHeader "Accept-Language", "pt-BR,pt;q=0.9,en;q=0.8"
    SetDefaultHeader "X-Client-Platform", "VB6-Windows"

    ' Autenticação (se disponível)
    Dim token As String
    token = LoadStoredToken()
    If Len(token) > 0 Then
        SetDefaultHeader "Authorization", "Bearer " & token
    End If
End Sub
```

### Tratamento de Erros

```vb
Function RequisicaoSegura(url As String) As Object
    On Error GoTo ErrorHandler

    Dim response As HttpResponse
    Set response = HttpGet(url)

    If response.IsSuccess Then
        Set RequisicaoSegura = response.Json
    Else
        LogError "HTTP " & response.StatusCode & ": " & response.StatusText
        Set RequisicaoSegura = Nothing
    End If

    Exit Function
ErrorHandler:
    LogError "Erro na requisição: " & Err.Description
    Set RequisicaoSegura = Nothing
End Function
```

---

**📝 Nota**: Este módulo é o coração do sistema de consumo de APIs. Todos os outros componentes trabalham em conjunto com ele para fornecer uma experiência completa e robusta de integração com APIs REST.
