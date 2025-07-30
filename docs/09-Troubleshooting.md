## 🔧 Troubleshooting

### Problemas Comuns e Soluções

#### Erro: "Tipo definido pelo usuário não definido"

**Causa**: Referências não configuradas corretamente.

**Solução**:

```vb
' Verifique se estas referências estão marcadas:
' - Microsoft Scripting Runtime
' - Microsoft XML HTTP Services
```

#### Erro: "Objeto requerido" ao fazer parsing JSON

**Causa**: Resposta da API não é um JSON válido.

**Solução**:

```vb
' Sempre verifique a resposta antes do parsing
Dim response As HttpResponse
Set response = HttpGet("/endpoint")

If response.IsSuccess Then
    Debug.Print "Resposta bruta: " & response.Text

    ' Só faça parsing se for JSON válido
    If Left(Trim(response.Text), 1) = "{" Or Left(Trim(response.Text), 1) = "[" Then
        Dim jsonObj As Object
        Set jsonObj = ParseJSON(response.Text)
    End If
End If
```

#### Timeout de Conexão

**Causa**: API lenta ou problemas de rede.

**Solução**:

```vb
' Aumentar timeout na inicialização
InitializeHttpClient "https://api.lenta.com", 60000  ' 60 segundos

' Ou implementar retry
Sub RequisicaoComRetry()
    Dim tentativas As Integer
    Dim response As HttpResponse

    For tentativas = 1 To 3
        Set response = HttpGet("/endpoint")
        If response.IsSuccess Then Exit For

        Debug.Print "Tentativa " & tentativas & " falhou, tentando novamente..."
        Sleep 2000  ' Aguarda 2 segundos
    Next tentativas
End Sub
```

#### Erro 401: Unauthorized

**Causa**: Token de autenticação inválido ou expirado.

**Solução**:

```vb
' Implementar renovação automática de token
Sub RenovarToken()
    Dim tokenData As Dictionary
    Set tokenData = CreateJSONObject()
    tokenData.Add "refresh_token", GetStoredRefreshToken()

    Dim newToken As Object
    Set newToken = PostJson("/auth/refresh", tokenData)

    If Not newToken Is Nothing Then
        SetDefaultHeader "Authorization", "Bearer " & newToken("access_token")
        SaveToken newToken("access_token"), newToken("refresh_token")
    End If
End Sub
```

### Debugging e Logs

```vb
' Habilitar logs detalhados para debug
Sub HabilitarDebug()
    ' Adicione este código antes das requisições para debug
    Debug.Print "=== DEBUG REQUISIÇÃO ==="
    Debug.Print "URL: " & url
    Debug.Print "Method: " & method
    Debug.Print "Headers: " & headersString
    Debug.Print "Body: " & requestBody
    Debug.Print "========================"
End Sub
```

