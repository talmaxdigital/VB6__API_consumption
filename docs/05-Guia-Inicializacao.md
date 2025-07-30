# Guia de Inicialização - VB6 API Consumption

Este guia aborda todos os aspectos da configuração inicial e primeiros passos com o sistema de consumo de APIs em VB6.

## 📋 Índice

- [Pré-requisitos](#pré-requisitos)
- [Configuração do Ambiente](#configuração-do-ambiente)
- [Primeiro Projeto](#primeiro-projeto)
- [Configuração Básica](#configuração-básica)
- [Primeira Requisição](#primeira-requisição)
- [Verificação e Testes](#verificação-e-testes)

## Pré-requisitos

### Sistema Operacional

✅ Windows 7 ou superior
✅ Visual Basic 6.0 IDE instalado
✅ VB6 Runtime (para execução)
✅ Conexão com internet (para testes)

### Componentes do Sistema Necessários

Componentes Windows obrigatórios:

1. Microsoft Scripting Runtime (scrrun.dll)
    - Localização: C:\Windows\System32\scrrun.dll
    - Fornece: Dictionary e Collection

2. Microsoft XML HTTP Services
    - msxml6.dll (preferencial) ou msxml3.dll (fallback)
    - Localização: C:\Windows\System32\
    - Fornece: Objeto XMLHTTP para requisições

### Verificação dos Componentes

```vb
' Código para verificar se componentes estão disponíveis:
Sub VerificarComponentes()
    On Error GoTo ComponenteNaoEncontrado

    ' Testar Scripting Runtime
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Debug.Print "Scripting Runtime: OK"

    ' Testar XMLHTTP
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    Debug.Print "XMLHTTP: OK"

    Debug.Print "Todos os componentes estão disponíveis!"
    Exit Sub

ComponenteNaoEncontrado:
    Debug.Print "Erro: " & Err.Description
    Debug.Print "Componente necessário não encontrado"
End Sub
```

## Configuração do Ambiente

### Passo 1: Configurar Referências no VB6

1. Abra o Visual Basic 6.0
2. Vá em **Project → References**
3. Marque as seguintes referências:

``` md
☑️ Microsoft Scripting Runtime
   Arquivo: C:\Windows\System32\scrrun.dll

☑️ Microsoft XML, v6.0 (preferencial)
   Arquivo: C:\Windows\System32\msxml6.dll

   OU (se v6.0 não disponível):

☑️ Microsoft XML, v3.0
   Arquivo: C:\Windows\System32\msxml3.dll
```

## Primeiro Projeto

### Criando Projeto do Zero

1. Criar novo projeto Standard EXE
2. Salvar com o nome desejado
3. Configurar referências (ver seção anterior)
4. Adicionar módulos e classes

### Importando os Arquivos

**Método 1: Usar Projeto como Base**

1. Abra `ConsumoAPI.vbp`
2. Adicione formulários necessários
3. Use como base para seu projeto

**Método 2: Importar Módulos Individuais**

1. **Project → Add Module** → Selecione `HttpClient.bas`
2. **Project → Add Module** → Selecione `JsonHelper.bas`
3. **Project → Add Class Module** → Selecione `HttpResponse.cls`
4. **Project → Add Class Module** → Selecione `cHttpRequest.cls`

### Validação da Importação

```vb
' Adicione este código em um formulário para testar:
Private Sub Form_Load()
    ' Testar criação de objetos JSON
    Dim obj As Dictionary
    Set obj = CreateJSONObject()
    obj.Add "teste", "ok"

    Dim json As String
    json = BuildJSON(obj)
    Debug.Print "JSON criado: " & json

    ' Testar parsing
    Dim parsed As Object
    Set parsed = ParseJSON(json)
    Debug.Print "Valor parseado: " & parsed("teste")

    MsgBox "Sistema funcionando corretamente!", vbInformation
End Sub
```

## Configuração Básica

### Inicialização Simples

```vb
' No Form_Load ou Sub Main:
Sub InicializarSistema()
    ' Configuração mínima
    InitializeHttpClient

    Debug.Print "Sistema HTTP inicializado"
End Sub
```

### Configuração com URL Base

```vb
' Para APIs com URL base fixa:
Sub InicializarComURL()
    InitializeHttpClient "https://jsonplaceholder.typicode.com", 30000, "MeuApp/1.0"

    Debug.Print "Cliente HTTP configurado para JSONPlaceholder"
End Sub
```

### Configuração Completa com Headers

```vb
' Configuração mais robusta:
Sub InicializarCompleto()
    ' URL base, timeout, user-agent
    InitializeHttpClient "https://api.github.com", 20000, "VB6-GitHubClient/1.0"

    ' Headers globais
    SetDefaultHeader "Accept", "application/vnd.github.v3+json"
    SetDefaultHeader "X-Client-Platform", "VB6-Windows"

    ' Token de autenticação (se disponível)
    Dim token As String
    token = GetStoredToken()  ' Função personalizada
    If Len(token) > 0 Then
        SetDefaultHeader "Authorization", "Bearer " & token
    End If

    Debug.Print "Cliente GitHub inicializado com autenticação"
End Sub

' Função auxiliar para carregar token salvo
Private Function GetStoredToken() As String
    ' Implementar lógica de carregamento de token
    ' Pode ser de arquivo, registry, etc.
    GetStoredToken = ""  ' Retorna vazio se não houver token
End Function
```

## Primeira Requisição

### GET Simples

```vb
' Primeira requisição básica:
Sub PrimeiraRequisicao()
    ' Inicializar sistema
    InitializeHttpClient

    ' Fazer requisição GET
    Dim response As HttpResponse
    Set response = HttpGet("https://jsonplaceholder.typicode.com/posts/1")

    ' Verificar resultado
    If response.IsSuccess Then
        Debug.Print "Requisição bem-sucedida!"
        Debug.Print "Status: " & response.StatusCode
        Debug.Print "Resposta: " & Left(response.Text, 200)
    Else
        Debug.Print "Erro na requisição: " & response.StatusCode
    End If
End Sub
```

### GET com JSON Parsing

```vb
' Requisição com processamento automático de JSON:
Sub RequisicaoComJSON()
    InitializeHttpClient

    ' Usar método especializado para JSON
    Dim post As Object
    Set post = GetJson("https://jsonplaceholder.typicode.com/posts/1")

    If Not post Is Nothing Then
        Debug.Print Post encontrado:"
        Debug.Print "ID: " & post("id")
        Debug.Print "Título: " & post("title")
        Debug.Print "Conteúdo: " & Left(post("body"), 50) & "..."
    Else
        Debug.Print "Erro ao obter post"
    End If
End Sub
```

### POST Básico

```vb
' Criar novo recurso via POST:
Sub CriarPost()
    InitializeHttpClient

    ' Criar dados para enviar
    Dim novoPost As Dictionary
    Set novoPost = CreateJSONObject()
    novoPost.Add "title", "Meu primeiro post via VB6"
    novoPost.Add "body", "Este post foi criado usando VB6!"
    novoPost.Add "userId", 1

    ' Enviar via POST
    Dim postCriado As Object
    Set postCriado = PostJson("https://jsonplaceholder.typicode.com/posts", novoPost)

    If Not postCriado Is Nothing Then
        Debug.Print "Post criado com sucesso!"
        Debug.Print "ID: " & postCriado("id")
        Debug.Print "Título: " & postCriado("title")
    Else
        Debug.Print "Erro ao criar post"
    End If
End Sub
```

## Verificação e Testes

### Teste de Conectividade

```vb
' Verificar se consegue acessar internet:
Function TestarConectividade() As Boolean
    On Error GoTo ErroConexao

    InitializeHttpClient

    Dim response As HttpResponse
    Set response = HttpGet("https://httpbin.org/status/200")

    TestarConectividade = response.IsSuccess
    Exit Function

ErroConexao:
    Debug.Print "Erro de conectividade: " & Err.Description
    TestarConectividade = False
End Function
```

### Teste de Componentes

```vb
' Verificar se todos os componentes funcionam:
Sub TestarTodosComponentes()
    Debug.Print "=== TESTE DE COMPONENTES ==="

    ' Teste 1: Criação de objetos JSON
    On Error GoTo Erro1
    Dim obj As Dictionary
    Set obj = CreateJSONObject()
    obj.Add "teste", "ok"
    Debug.Print "Criação de objetos JSON: OK"
    GoTo Teste2
Erro1:
    Debug.Print "Criação de objetos JSON: FALHOU"

Teste2:
    ' Teste 2: Conversão para JSON
    On Error GoTo Erro2
    Dim json As String
    json = BuildJSON(obj)
    Debug.Print "Conversão para JSON: OK"
    GoTo Teste3
Erro2:
    Debug.Print "Conversão para JSON: FALHOU"

Teste3:
    ' Teste 3: Parse de JSON
    On Error GoTo Erro3
    Dim parsed As Object
    Set parsed = ParseJSON(json)
    Debug.Print "Parse de JSON: OK"
    GoTo Teste4
Erro3:
    Debug.Print "Parse de JSON: FALHOU"

Teste4:
    ' Teste 4: Requisição HTTP
    On Error GoTo Erro4
    InitializeHttpClient
    Dim response As HttpResponse
    Set response = HttpGet("https://httpbin.org/json")
    If response.IsSuccess Then
        Debug.Print "Requisições HTTP: OK"
    Else
        Debug.Print "Requisições HTTP: Status " & response.StatusCode
    End If
    GoTo FimTeste
Erro4:
    Debug.Print "Requisições HTTP: FALHOU - " & Err.Description

FimTeste:
    Debug.Print "=== FIM DOS TESTES ==="
    On Error GoTo 0
End Sub
```

### Exemplo Completo Funcional

```vb
' Exemplo completo que você pode usar como template:
Private Sub btnTeste_Click()
    ' Limpar debug
    Debug.Print String(50, "=")
    Debug.Print "INICIANDO TESTE COMPLETO"
    Debug.Print String(50, "=")

    ' 1. Inicializar
    Debug.Print "1. Inicializando sistema..."
    InitializeHttpClient "https://jsonplaceholder.typicode.com", 15000, "VB6-TestApp/1.0"

    ' 2. Teste GET simples
    Debug.Print "2. Testando GET simples..."
    Dim response As HttpResponse
    Set response = HttpGet("/posts/1")

    If response.IsSuccess Then
        Debug.Print "GET: " & response.StatusCode & " " & response.StatusText
    Else
        Debug.Print "GET falhou: " & response.StatusCode
        Exit Sub
    End If

    ' 3. Teste GET com JSON
    Debug.Print "3. Testando GET com JSON parsing..."
    Dim post As Object
    Set post = GetJson("/posts/1")

    If Not post Is Nothing Then
        Debug.Print "JSON parsing: Post ID " & post("id")
        Debug.Print "Título: " & post("title")
    Else
        Debug.Print "JSON parsing falhou"
        Exit Sub
    End If

    ' 4. Teste POST
    Debug.Print "4. Testando POST..."
    Dim novoPost As Dictionary
    Set novoPost = CreateJSONObject()
    novoPost.Add "title", "Teste VB6"
    novoPost.Add "body", "Post criado via VB6"
    novoPost.Add "userId", 1

    Dim postCriado As Object
    Set postCriado = PostJson("/posts", novoPost)

    If Not postCriado Is Nothing Then
        Debug.Print "POST: Criado com ID " & postCriado("id")
    Else
        Debug.Print "POST falhou"
    End If

    Debug.Print String(50, "=")
    Debug.Print "TESTE COMPLETO FINALIZADO"
    Debug.Print String(50, "=")

    MsgBox "Teste completo executado! Veja a janela Debug para detalhes.", vbInformation
End Sub
```

### Solução de Problemas Comuns

```vb
' Diagnóstico de problemas:
Sub DiagnosticarProblemas()
    Debug.Print "=== DIAGNÓSTICO ==="

    ' Verificar referências
    On Error Resume Next
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then
        Debug.Print "Scripting Runtime não disponível"
        Debug.Print "   Solução: Instalar/registrar scrrun.dll"
    Else
        Debug.Print "Scripting Runtime: OK"
    End If

    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    If Err.Number <> 0 Then
        Debug.Print "XMLHTTP não disponível"
        Debug.Print "   Solução: Instalar/registrar msxml6.dll ou msxml3.dll"
    Else
        Debug.Print "XMLHTTP: OK"
    End If

    ' Verificar conectividade
    Err.Clear
    xmlhttp.Open "GET", "https://httpbin.org/status/200", False
    xmlhttp.Send
    If Err.Number <> 0 Then
        Debug.Print "Problemas de conectividade"
        Debug.Print "   Erro: " & Err.Description
    Else
        Debug.Print "Conectividade: OK"
    End If

    On Error GoTo 0
    Debug.Print "====================="
End Sub
```

---

**🎯 Próximos Passos**: Após concluir a inicialização, recomenda-se estudar o [Guia de JSON](06-Trabalhando-JSON.md) e [Requisições HTTP](07-Requisicoes-HTTP.md) para aprofundar o conhecimento.

**🔧 Dica**: Mantenha o método `TestarTodosComponentes()` no seu projeto para diagnosticar problemas rapidamente durante o desenvolvimento.
