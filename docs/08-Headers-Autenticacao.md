# Headers e Autenticação - Guia Completo

Este guia aborda todos os aspectos de configuração de headers HTTP e implementação de diferentes métodos de autenticação em APIs.

## 📋 Índice

- [Sistema de Headers](#sistema-de-headers)
- [Métodos de Autenticação](#métodos-de-autenticação)
- [Gestão de Tokens](#gestão-de-tokens)
- [Headers Especializados](#headers-especializados)
- [Boas Práticas de Segurança](#boas-práticas-de-segurança)

## Sistema de Headers

### Headers Padrão vs Customizados

```vb
' Headers padrão configurados automaticamente
Sub ConfiguracaoAutomatica()
    InitializeHttpClient "https://api.exemplo.com", 30000, "MeuApp/1.0"

    ' Os seguintes headers são adicionados automaticamente:
    ' User-Agent: MeuApp/1.0
    ' Accept: application/json
    ' Accept-Encoding: gzip, deflate
End Sub

' Sobrescrever headers padrão
Sub SobrescreverPadrao()
    InitializeHttpClient "https://api.exemplo.com"

    ' Sobrescrever User-Agent padrão
    SetDefaultHeader "User-Agent", "VB6-CustomClient/2.0 (Windows)"

    ' Sobrescrever Accept padrão
    SetDefaultHeader "Accept", "application/vnd.api+json"
End Sub

' Headers por requisição (não afetam configuração global)
Sub HeadersPorRequisicao()
    Dim headers As Dictionary
    Set headers = CreateJSONObject()
    headers.Add "Accept", "application/xml"  ' Só para esta requisição
    headers.Add "X-Request-ID", GenerateRequestID()

    Dim response As HttpResponse
    Set response = HttpGet("/data", headers)
End Sub
```

### Gerenciamento de Headers Globais

```vb
' Configuração completa de headers globais
Sub ConfigurarHeadersGlobais()
    InitializeHttpClient "https://api.minhaempresa.com/v2"

    ' Identificação do cliente
    SetDefaultHeader "User-Agent", "ERP-VB6/3.1 (Windows NT 10.0)"
    SetDefaultHeader "X-Client-Version", "3.1.0"
    SetDefaultHeader "X-Client-Platform", "VB6-Windows"

    ' Formato de dados
    SetDefaultHeader "Accept", "application/json"
    SetDefaultHeader "Accept-Language", "pt-BR,pt;q=0.9,en;q=0.8"
    SetDefaultHeader "Accept-Encoding", "gzip, deflate"

    ' Configurações de cache
    SetDefaultHeader "Cache-Control", "no-cache"
    SetDefaultHeader "Pragma", "no-cache"

    ' Headers customizados da empresa
    SetDefaultHeader "X-Company-ID", "ACME001"
    SetDefaultHeader "X-Environment", "production"
End Sub

' Remover headers quando necessário
Sub GerenciarHeadersDinamicamente()
    ' Configurar para ambiente de desenvolvimento
    SetDefaultHeader "X-Environment", "development"
    SetDefaultHeader "X-Debug-Mode", "true"

    ' ... código da aplicação ...

    ' Remover headers de debug para produção
    RemoveDefaultHeader "X-Debug-Mode"
    SetDefaultHeader "X-Environment", "production"
End Sub
```

### Headers de Conteúdo

```vb
' Content-Type para diferentes tipos de dados
Sub ConfigurarContentType()
    Dim headers As Dictionary

    ' JSON (mais comum)
    Set headers = CreateJSONObject()
    headers.Add "Content-Type", "application/json; charset=utf-8"

    ' Form data URL encoded
    Set headers = CreateJSONObject()
    headers.Add "Content-Type", "application/x-www-form-urlencoded"

    ' Multipart para upload de arquivos
    Dim boundary As String
    boundary = "----VB6FormBoundary" & Format(Timer * 1000, "0")
    Set headers = CreateJSONObject()
    headers.Add "Content-Type", "multipart/form-data; boundary=" & boundary

    ' XML
    Set headers = CreateJSONObject()
    headers.Add "Content-Type", "application/xml; charset=utf-8"

    ' Texto simples
    Set headers = CreateJSONObject()
    headers.Add "Content-Type", "text/plain; charset=utf-8"
End Sub

' Accept para especificar formato de resposta
Sub ConfigurarAccept()
    Dim headers As Dictionary
    Set headers = CreateJSONObject()

    ' JSON apenas
    headers.Add "Accept", "application/json"

    ' Múltiplos formatos com prioridade
    headers.Add "Accept", "application/json, application/xml;q=0.9, text/plain;q=0.8"

    ' Versão específica da API
    headers.Add "Accept", "application/vnd.github.v3+json"

    ' Com charset específico
    headers.Add "Accept", "application/json; charset=utf-8"
End Sub
```

## Métodos de Autenticação

### Bearer Token (JWT, OAuth2)

```vb
' Configuração simples de Bearer Token
Sub ConfigurarBearerToken()
    Dim token As String
    token = "eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiIxMjM0NTY3ODkwIiwibmFtZSI6IkpvaG4gRG9lIiwiYWRtaW4iOnRydWV9.TJVA95OrM7E2cBab30RMHrHDcEfxjoYZgeFONFh7HgQ"

    SetDefaultHeader "Authorization", "Bearer " & token

    Debug.Print "Bearer token configurado"
End Sub

' Sistema completo de JWT com renovação
Public Class CJWTManager
    Private m_AccessToken As String
    Private m_RefreshToken As String
    Private m_TokenExpiry As Date

    Public Sub SetTokens(accessToken As String, refreshToken As String, expiresIn As Long)
        m_AccessToken = accessToken
        m_RefreshToken = refreshToken
        m_TokenExpiry = DateAdd("s", expiresIn - 60, Now)  ' Renovar 1 minuto antes

        SetDefaultHeader "Authorization", "Bearer " & m_AccessToken
    End Sub

    Public Function IsTokenExpired() As Boolean
        IsTokenExpired = (Now >= m_TokenExpiry)
    End Function

    Public Function RenovarToken() As Boolean
        If Len(m_RefreshToken) = 0 Then
            RenovarToken = False
            Exit Function
        End If

        Dim refreshData As Dictionary
        Set refreshData = CreateJSONObject()
        refreshData.Add "refresh_token", m_RefreshToken
        refreshData.Add "grant_type", "refresh_token"

        ' Remover header temporariamente para renovação
        RemoveDefaultHeader "Authorization"

        Dim response As Object
        Set response = PostJson("/auth/refresh", refreshData)

        If Not response Is Nothing Then
            If response.Exists("access_token") Then
                SetTokens response("access_token"), response("refresh_token"), response("expires_in")
                RenovarToken = True
            Else
                RenovarToken = False
            End If
        Else
            RenovarToken = False
        End If
    End Function

    Public Sub ValidarToken()
        If IsTokenExpired() Then
            Debug.Print "Token expirado, renovando..."
            If Not RenovarToken() Then
                Debug.Print "Falha ao renovar token - reautenticação necessária"
                ' Implementar re-login
            End If
        End If
    End Sub
End Class

' Uso do gerenciador JWT
Sub ExemploJWT()
    Dim jwtManager As New CJWTManager

    ' Login inicial
    Dim loginData As Dictionary
    Set loginData = CreateJSONObject()
    loginData.Add "username", "usuario"
    loginData.Add "password", "senha"

    Dim authResponse As Object
    Set authResponse = PostJson("/auth/login", loginData)

    If Not authResponse Is Nothing Then
        jwtManager.SetTokens authResponse("access_token"), authResponse("refresh_token"), authResponse("expires_in")

        ' Usar API normalmente
        jwtManager.ValidarToken()  ' Verificar/renovar antes de usar
        Dim userData As Object
        Set userData = GetJson("/user/profile")
    End If
End Sub
```

### API Key Authentication

```vb
' API Key no header
Sub ConfigurarAPIKeyHeader()
    SetDefaultHeader "X-API-Key", "sk_live_51H123456789abcdef"
    SetDefaultHeader "X-API-Secret", "whsec_abcdef123456789"
End Sub

' API Key com nome customizado
Sub ConfigurarAPIKeyCustom()
    ' Diferentes APIs usam nomes diferentes
    SetDefaultHeader "Authorization", "API-Key sk_test_123456"
    SetDefaultHeader "X-Auth-Token", "abc123def456"
    SetDefaultHeader "Apikey", "sua-chave-aqui"
    SetDefaultHeader "X-RapidAPI-Key", "sua-rapid-api-key"
End Sub

' API Key na URL (menos seguro, mas algumas APIs usam)
Sub ConfigurarAPIKeyURL()
    InitializeHttpClient "https://api.exemplo.com?api_key=sua-chave-aqui"
    ' Todas as requisições terão a chave na URL
End Sub

' Sistema de múltiplas chaves (para diferentes serviços)
Public Class CAPIKeyManager
    Private m_Keys As Dictionary

    Private Sub Class_Initialize()
        Set m_Keys = CreateJSONObject()
    End Sub

    Public Sub AddKey(service As String, keyName As String, keyValue As String)
        Dim keyInfo As Dictionary
        Set keyInfo = CreateJSONObject()
        keyInfo.Add "name", keyName
        keyInfo.Add "value", keyValue

        m_Keys.Add service, keyInfo
    End Sub

    Public Sub UseService(service As String)
        If m_Keys.Exists(service) Then
            Dim keyInfo As Dictionary
            Set keyInfo = m_Keys(service)

            SetDefaultHeader keyInfo("name"), keyInfo("value")
            Debug.Print "API Key configurada para serviço: " & service
        Else
            Debug.Print "Serviço não encontrado: " & service
        End If
    End Sub
End Class

' Uso do gerenciador de chaves
Sub ExemploAPIKeys()
    Dim keyManager As New CAPIKeyManager

    ' Registrar diferentes serviços
    keyManager.AddKey "github", "Authorization", "token ghp_xxxxxxxxxxxx"
    keyManager.AddKey "stripe", "Authorization", "Bearer sk_test_xxxxxxxxxxxx"
    keyManager.AddKey "openai", "Authorization", "Bearer sk-xxxxxxxxxxxx"

    ' Usar GitHub API
    keyManager.UseService "github"
    Dim repos As Object
    Set repos = GetJson("https://api.github.com/user/repos")

    ' Mudar para Stripe API
    keyManager.UseService "stripe"
    Dim customers As Object
    Set customers = GetJson("https://api.stripe.com/v1/customers")
End Sub
```

### Basic Authentication

```vb
' Basic Auth simples
Sub ConfigurarBasicAuth()
    Dim usuario As String
    Dim senha As String
    usuario = "admin"
    senha = "senha123"

    ' No VB6, precisamos implementar Base64 encoding
    Dim credentials As String
    credentials = usuario & ":" & senha

    ' Assumindo função Base64Encode disponível
    SetDefaultHeader "Authorization", "Basic " & Base64Encode(credentials)
End Sub

' Implementação simplificada de Base64 para VB6
Function Base64Encode(text As String) As String
    ' Esta é uma implementação muito simplificada
    ' Para uso em produção, use uma biblioteca Base64 completa
    Dim i As Integer
    Dim result As String

    ' Implementação básica - em produção use biblioteca completa
    ' Por ora, retornando placeholder
    Base64Encode = "Basic_" & text & "_Encoded"

    ' TODO: Implementar Base64 real ou usar biblioteca externa
End Function

' Classe para Basic Auth
Public Class CBasicAuth
    Private m_Username As String
    Private m_Password As String

    Public Sub SetCredentials(username As String, password As String)
        m_Username = username
        m_Password = password
        UpdateAuthHeader
    End Sub

    Private Sub UpdateAuthHeader()
        Dim credentials As String
        credentials = m_Username & ":" & m_Password
        SetDefaultHeader "Authorization", "Basic " & Base64Encode(credentials)
    End Sub

    Public Sub ClearCredentials()
        m_Username = ""
        m_Password = ""
        RemoveDefaultHeader "Authorization"
    End Sub
End Class
```

### Custom Authentication

```vb
' Assinatura HMAC (usado por AWS, muitas APIs)
Function GerarAssinaturaHMAC(message As String, secretKey As String) As String
    ' Esta é uma implementação placeholder
    ' Em produção, use biblioteca HMAC-SHA256 adequada

    Dim timestamp As String
    timestamp = CStr(DateDiff("s", #1/1/1970#, Now))

    ' Concatenar dados para assinatura
    Dim stringToSign As String
    stringToSign = timestamp & message & secretKey

    ' Simular hash (implementar HMAC real)
    GerarAssinaturaHMAC = "hmac_" & Len(stringToSign) & "_" & timestamp
End Function

' Autenticação com assinatura customizada
Sub ConfigurarCustomSignature()
    Dim apiKey As String
    Dim secretKey As String
    Dim timestamp As String

    apiKey = "minha-api-key"
    secretKey = "minha-secret-key"
    timestamp = CStr(DateDiff("s", #1/1/1970#, Now))

    ' Dados para assinatura
    Dim method As String
    Dim path As String
    method = "GET"
    path = "/api/users"

    Dim stringToSign As String
    stringToSign = method & vbLf & path & vbLf & timestamp

    Dim signature As String
    signature = GerarAssinaturaHMAC(stringToSign, secretKey)

    ' Configurar headers
    SetDefaultHeader "X-API-Key", apiKey
    SetDefaultHeader "X-Timestamp", timestamp
    SetDefaultHeader "X-Signature", signature
End Sub

' Sistema de autenticação por sessão
Public Class CSessionAuth
    Private m_SessionID As String
    Private m_CSRFToken As String

    Public Function Login(username As String, password As String) As Boolean
        Dim loginData As Dictionary
        Set loginData = CreateJSONObject()
        loginData.Add "username", username
        loginData.Add "password", password

        Dim response As HttpResponse
        Set response = HttpPost("/auth/login", BuildJSON(loginData))

        If response.IsSuccess Then
            Dim authData As Object
            Set authData = response.Json

            If authData.Exists("session_id") Then
                m_SessionID = authData("session_id")
                SetDefaultHeader "Cookie", "sessionid=" & m_SessionID

                If authData.Exists("csrf_token") Then
                    m_CSRFToken = authData("csrf_token")
                    SetDefaultHeader "X-CSRFToken", m_CSRFToken
                End If

                Login = True
            Else
                Login = False
            End If
        Else
            Login = False
        End If
    End Function

    Public Sub Logout()
        HttpPost "/auth/logout", ""

        m_SessionID = ""
        m_CSRFToken = ""
        RemoveDefaultHeader "Cookie"
        RemoveDefaultHeader "X-CSRFToken"
    End Sub
End Class
```

## Gestão de Tokens

### Armazenamento Seguro

```vb
' Classe para gerenciar armazenamento de tokens
Public Class CTokenStorage
    Private Const REGISTRY_KEY = "SOFTWARE\MinhaEmpresa\MeuApp\Auth"

    Public Sub SaveToken(tokenType As String, tokenValue As String)
        ' Salvar no Registry (criptografado)
        On Error GoTo ErrorHandler

        ' Implementar criptografia simples
        Dim encryptedToken As String
        encryptedToken = SimpleEncrypt(tokenValue)

        SaveSetting "MeuApp", "Auth", tokenType, encryptedToken
        Exit Sub

ErrorHandler:
        Debug.Print "Erro ao salvar token: " & Err.Description
    End Sub

    Public Function LoadToken(tokenType As String) As String
        On Error GoTo ErrorHandler

        Dim encryptedToken As String
        encryptedToken = GetSetting("MeuApp", "Auth", tokenType, "")

        If Len(encryptedToken) > 0 Then
            LoadToken = SimpleDecrypt(encryptedToken)
        Else
            LoadToken = ""
        End If
        Exit Function

ErrorHandler:
        Debug.Print "Erro ao carregar token: " & Err.Description
        LoadToken = ""
    End Function

    Public Sub ClearTokens()
        DeleteSetting "MeuApp", "Auth"
    End Sub

    Private Function SimpleEncrypt(text As String) As String
        ' Implementação muito simples - usar criptografia real em produção
        Dim i As Integer
        Dim result As String

        For i = 1 To Len(text)
            result = result & Chr(Asc(Mid(text, i, 1)) + 5)
        Next i

        SimpleEncrypt = result
    End Function

    Private Function SimpleDecrypt(text As String) As String
        ' Implementação muito simples - usar criptografia real em produção
        Dim i As Integer
        Dim result As String

        For i = 1 To Len(text)
            result = result & Chr(Asc(Mid(text, i, 1)) - 5)
        Next i

        SimpleDecrypt = result
    End Function
End Class
```

### Renovação Automática

```vb
' Sistema completo de renovação automática
Public Class CAutoRenewalManager
    Private m_TokenStorage As CTokenStorage
    Private m_AccessToken As String
    Private m_RefreshToken As String
    Private m_ExpiryTime As Date
    Private m_RenewalInProgress As Boolean

    Private Sub Class_Initialize()
        Set m_TokenStorage = New CTokenStorage
        m_RenewalInProgress = False
        LoadStoredTokens
    End Sub

    Private Sub LoadStoredTokens()
        m_AccessToken = m_TokenStorage.LoadToken("access_token")
        m_RefreshToken = m_TokenStorage.LoadToken("refresh_token")

        Dim expiryStr As String
        expiryStr = m_TokenStorage.LoadToken("expiry_time")
        If Len(expiryStr) > 0 Then
            m_ExpiryTime = CDate(expiryStr)
        End If

        If Len(m_AccessToken) > 0 Then
            SetDefaultHeader "Authorization", "Bearer " & m_AccessToken
        End If
    End Sub

    Public Function EnsureValidToken() As Boolean
        If m_RenewalInProgress Then
            EnsureValidToken = False
            Exit Function
        End If

        If IsTokenExpired() Or Len(m_AccessToken) = 0 Then
            EnsureValidToken = RenewToken()
        Else
            EnsureValidToken = True
        End If
    End Function

    Private Function IsTokenExpired() As Boolean
        IsTokenExpired = (Now >= DateAdd("n", -5, m_ExpiryTime))  ' Renovar 5 min antes
    End Function

    Private Function RenewToken() As Boolean
        If Len(m_RefreshToken) = 0 Then
            Debug.Print "Refresh token não disponível - reautenticação necessária"
            RenewToken = False
            Exit Function
        End If

        m_RenewalInProgress = True

        ' Remover header temporariamente
        RemoveDefaultHeader "Authorization"

        Dim renewalData As Dictionary
        Set renewalData = CreateJSONObject()
        renewalData.Add "refresh_token", m_RefreshToken
        renewalData.Add "grant_type", "refresh_token"

        Dim response As HttpResponse
        Set response = HttpPost("/auth/refresh", BuildJSON(renewalData))

        If response.IsSuccess Then
            Dim tokenData As Object
            Set tokenData = response.Json

            If tokenData.Exists("access_token") Then
                m_AccessToken = tokenData("access_token")
                m_RefreshToken = tokenData("refresh_token")

                Dim expiresIn As Long
                expiresIn = tokenData("expires_in")
                m_ExpiryTime = DateAdd("s", expiresIn, Now)

                ' Salvar tokens
                m_TokenStorage.SaveToken "access_token", m_AccessToken
                m_TokenStorage.SaveToken "refresh_token", m_RefreshToken
                m_TokenStorage.SaveToken "expiry_time", CStr(m_ExpiryTime)

                ' Reconfigurar header
                SetDefaultHeader "Authorization", "Bearer " & m_AccessToken

                Debug.Print "Token renovado com sucesso"
                RenewToken = True
            Else
                Debug.Print "Resposta de renovação inválida"
                RenewToken = False
            End If
        Else
            Debug.Print "Falha ao renovar token: " & response.StatusCode
            RenewToken = False
        End If

        m_RenewalInProgress = False
    End Function

    Public Sub ClearTokens()
        m_AccessToken = ""
        m_RefreshToken = ""
        m_ExpiryTime = Now

        m_TokenStorage.ClearTokens
        RemoveDefaultHeader "Authorization"
    End Sub
End Class

' Uso do sistema de renovação automática
Sub ExemploRenovacaoAutomatica()
    Dim authManager As New CAutoRenewalManager

    ' Tentar fazer requisição - renovará automaticamente se necessário
    If authManager.EnsureValidToken() Then
        Dim userData As Object
        Set userData = GetJson("/user/profile")

        If Not userData Is Nothing Then
            Debug.Print "Perfil obtido: " & userData("name")
        End If
    Else
        Debug.Print "Falha na autenticação - login necessário"
    End If
End Sub
```

## Headers Especializados

### Headers de Controle de Cache

```vb
' Configurar políticas de cache
Sub ConfigurarCache()
    ' Não usar cache
    SetDefaultHeader "Cache-Control", "no-cache, no-store, must-revalidate"
    SetDefaultHeader "Pragma", "no-cache"
    SetDefaultHeader "Expires", "0"

    ' Cache por 1 hora
    SetDefaultHeader "Cache-Control", "max-age=3600"

    ' Cache apenas se não modificado
    SetDefaultHeader "Cache-Control", "max-age=0, must-revalidate"
End Sub

' Headers condicionais para cache inteligente
Sub UsarCacheCondicional(etag As String, lastModified As String)
    Dim headers As Dictionary
    Set headers = CreateJSONObject()

    If Len(etag) > 0 Then
        headers.Add "If-None-Match", etag
    End If

    If Len(lastModified) > 0 Then
        headers.Add "If-Modified-Since", lastModified
    End If

    Dim response As HttpResponse
    Set response = HttpGet("/data", headers)

    If response.StatusCode = 304 Then
        Debug.Print "Dados não modificados - usar cache local"
    ElseIf response.IsSuccess Then
        Debug.Print "Dados atualizados - atualizar cache"
        ' Salvar novos ETag e Last-Modified
        Dim newEtag As String
        Dim newLastModified As String
        newEtag = response.GetHeader("ETag")
        newLastModified = response.GetHeader("Last-Modified")
    End If
End Sub
```

### Headers de Segurança

```vb
' Headers de segurança para APIs corporativas
Sub ConfigurarSeguranca()
    SetDefaultHeader "X-Requested-With", "XMLHttpRequest"
    SetDefaultHeader "X-Content-Type-Options", "nosniff"
    SetDefaultHeader "X-Frame-Options", "DENY"
    SetDefaultHeader "Strict-Transport-Security", "max-age=31536000"

    ' Request ID para rastreamento
    SetDefaultHeader "X-Request-ID", GenerateRequestID()

    ' Informações do cliente
    SetDefaultHeader "X-Client-IP", GetLocalIP()
    SetDefaultHeader "X-Forwarded-For", GetLocalIP()
End Sub

Function GenerateRequestID() As String
    ' Gerar ID único para cada requisição
    GenerateRequestID = "req_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Timer * 1000, "0")
End Function

Function GetLocalIP() As String
    ' Implementar obtenção do IP local
    GetLocalIP = "127.0.0.1"  ' Placeholder
End Function
```

### Headers de Versionamento

```vb
' Versionamento via header
Sub ConfigurarVersaoAPI()
    ' Versão específica
    SetDefaultHeader "API-Version", "2023-10-01"
    SetDefaultHeader "X-API-Version", "v2"

    ' Accept com versão
    SetDefaultHeader "Accept", "application/vnd.api.v2+json"

    ' Compatibilidade
    SetDefaultHeader "X-Client-Version", "1.5.0"
    SetDefaultHeader "X-Min-API-Version", "v1"
    SetDefaultHeader "X-Max-API-Version", "v3"
End Sub
```

## Boas Práticas de Segurança

### Validação de Certificados

```vb
' Configurar validação SSL (limitado no VB6)
Sub ConfigurarSSL()
    ' No VB6 com XMLHTTP, a validação SSL é automática
    ' Para controle avançado, seria necessário usar WinINet diretamente

    Debug.Print "AVISO: Sempre use HTTPS em produção"
    Debug.Print "AVISO: Valide certificados manualmente se necessário"
End Sub
```

### Sanitização de Headers

```vb
' Validar headers antes de enviar
Function ValidateHeaderValue(headerValue As String) As String
    ' Remover caracteres perigosos
    Dim cleanValue As String
    cleanValue = headerValue

    ' Remover quebras de linha (CRLF injection)
    cleanValue = Replace(cleanValue, vbCr, "")
    cleanValue = Replace(cleanValue, vbLf, "")
    cleanValue = Replace(cleanValue, Chr(13), "")
    cleanValue = Replace(cleanValue, Chr(10), "")

    ' Limitar tamanho
    If Len(cleanValue) > 8192 Then
        cleanValue = Left(cleanValue, 8192)
    End If

    ValidateHeaderValue = cleanValue
End Function

' Wrapper seguro para SetDefaultHeader
Sub SafeSetDefaultHeader(headerName As String, headerValue As String)
    Dim safeName As String
    Dim safeValue As String

    safeName = ValidateHeaderValue(headerName)
    safeValue = ValidateHeaderValue(headerValue)

    SetDefaultHeader safeName, safeValue
End Sub
```

### Logging de Segurança

```vb
' Log de tentativas de autenticação
Sub LogAuthAttempt(success As Boolean, method As String, Optional details As String = "")
    Dim logEntry As String
    logEntry = Format(Now, "yyyy-mm-dd hh:nn:ss") & " - "
    logEntry = logEntry & "AUTH " & IIf(success, "SUCCESS", "FAILURE") & " - "
    logEntry = logEntry & method

    If Len(details) > 0 Then
        logEntry = logEntry & " - " & details
    End If

    ' Log para arquivo ou sistema de log
    Debug.Print logEntry

    ' Em produção, escrever para arquivo de log seguro
    WriteToSecureLog logEntry
End Sub

Sub WriteToSecureLog(logEntry As String)
    ' Implementar escrita segura em log
    ' Por exemplo, arquivo com permissões restritas
    On Error Resume Next

    Dim fileNum As Integer
    fileNum = FreeFile
    Open App.Path & "\logs\security.log" For Append As #fileNum
    Print #fileNum, logEntry
    Close #fileNum
End Sub
```

---

**🔒 Importante**: Nunca inclua tokens ou senhas diretamente no código fonte. Use sempre variáveis de ambiente, arquivo de configuração ou armazenamento seguro.

**⚡ Performance**: Headers são enviados em todas as requisições. Mantenha-os concisos e remove headers desnecessários para otimizar performance.
