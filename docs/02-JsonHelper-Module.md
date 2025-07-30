# JsonHelper Module - Documentação Técnica

O `JsonHelper.bas` é um parser e gerador JSON completamente nativo para VB6, implementado sem dependências externas além do Scripting Runtime.

## 📋 Índice

- [Visão Geral](#visão-geral)
- [Arquitetura do Parser](#arquitetura-do-parser)
- [Funções Principais](#funções-principais)
- [Estruturas de Dados](#estruturas-de-dados)
- [Tipos Suportados](#tipos-suportados)
- [Tratamento de Erros](#tratamento-de-erros)
- [Casos de Uso Avançados](#casos-de-uso-avançados)

## Visão Geral

### Capacidades do Módulo

Funcionalidades principais:

1. Parse de JSON strings para objetos VB6 (Dictionary/Collection)
2. Geração de JSON strings a partir de objetos VB6
3. Suporte completo aos tipos JSON (string, number, boolean, null, object, array)
4. Tratamento de caracteres especiais e escape
5. Validação de sintaxe JSON
6. Criação de objetos JSON estruturados

### Arquitetura Técnica

``` shell
JsonHelper Module
├── JSONSTATE (Type)                # Estado do parser
├── Parse Functions                 # Análise de JSON
│   ├── ParseJSON()                 # Função principal
│   ├── ParseObject()               # Objetos JSON {}
│   ├── ParseArray()                # Arrays JSON []
│   ├── ParseValue()                # Valores genéricos
│   ├── ParseString()               # Strings com escape
│   ├── ParseNumber()               # Números (int/float)
│   ├── ParseTrue/False/Null()      # Literais
│   └── SkipWhitespace()            # Utilitário
├── Build Functions                 # Geração de JSON
│   ├── BuildJSON()                 # Função principal
│   ├── BuildValue()                # Valores genéricos
│   ├── BuildObject()               # Objetos Dictionary
│   ├── BuildArray()                # Arrays Collection
│   ├── BuildString()               # Strings com escape
│   ├── BuildNumber()               # Números formatados
│   └── BuildBoolean()              # true/false
└── Factory Functions               # Criação de objetos
    ├── CreateJSONObject()          # Dictionary vazio
    └── CreateJSONArray()           # Collection vazia
```

## Arquitetura do Parser

### Estado do Parser (JSONSTATE)

```vb
Private Type JSONSTATE
    Json As String      ' String JSON sendo analisada
    position As Long    ' Posição atual na string
End Type
```

**Funcionamento**: O parser mantém um estado global durante a análise, avançando caractere por caractere através da string JSON.

### Fluxo de Parsing

``` mermaid
flowchart TD
    A["String JSON"] --> B["ParseJSON()"]
    B --> C{"Identifica tipo raiz<br><b>{ ou [</b>"}
    C -->|"<b>{</b>"| D["ParseObject()<br><i>para objetos</i>"]
    C -->|"<b>[</b>"| E["ParseArray()<br><i>para arrays</i>"]
    D --> F["ParseValue()\npara cada elemento/propriedade"]
    E --> F
    F --> G["Dictionary/Collection\nresultante"]

    style A stroke:#333,stroke-width:1px
    style G stroke:#0066cc,stroke-width:1px
    style C stroke:#ff9900,stroke-width:1px
```

## Funções Principais

### ParseJSON

```vb
Public Function ParseJSON(ByVal jsonString As String) As Object
```

**Responsabilidade**: Função principal que analisa uma string JSON e retorna o objeto VB6 equivalente.

**Algoritmo Interno**:

Pseudocódigo do algoritmo:

1. Inicializar estado do parser
2. Pular espaços em branco iniciais
3. Identificar tipo raiz:
   - '{' → ParseObject()
   - '[' → ParseArray()
   - Erro se não for objeto ou array na raiz
4. Retornar objeto resultante

**Exemplos de Uso**:

```vb
' Objeto simples
Dim user As Object
Set user = ParseJSON("{""name"":""João"",""age"":30,""active"":true}")
Debug.Print user("name")    ' Output: João
Debug.Print user("age")     ' Output: 30
Debug.Print user("active")  ' Output: True

' Array simples
Dim colors As Object
Set colors = ParseJSON("[""red"",""green"",""blue""]")
Debug.Print colors(1)       ' Output: red (VB6 usa índice base 1)
Debug.Print colors(2)       ' Output: green
Debug.Print colors(3)       ' Output: blue

' Estrutura complexa
Dim complexData As Object
Set complexData = ParseJSON("{""users"":[{""id"":1,""name"":""Ana""},{""id"":2,""name"":""Carlos""}],""total"":2}")
Debug.Print complexData("users")(1)("name")  ' Output: Ana
Debug.Print complexData("total")             ' Output: 2
```

### BuildJSON

```vb
Public Function BuildJSON(ByVal obj As Variant) As String
```

**Responsabilidade**: Converte objetos VB6 (Dictionary, Collection, tipos primitivos) para string JSON válida.

**Algoritmo de Conversão**:

```vb
' Lógica de identificação de tipo:
If IsObject(obj) Then
    If TypeName(obj) = "Dictionary" Then → BuildObject()
    ElseIf TypeName(obj) = "Collection" Then → BuildArray()
    Else → Erro (tipo não suportado)
ElseIf IsNull(obj) Then → "null"
ElseIf VarType(obj) = vbBoolean Then → "true"/"false"
ElseIf VarType(obj) = vbString Then → BuildString() com escape
ElseIf IsNumeric(obj) Then → Formato numérico
Else → Conversão para string
```

**Exemplos Práticos**:

```vb
' Construir objeto complexo
Dim produto As Dictionary
Set produto = CreateJSONObject()
produto.Add "id", 123
produto.Add "nome", "Notebook Dell"
produto.Add "preco", 2599.99
produto.Add "disponivel", True
produto.Add "descricao", Null

Dim jsonString As String
jsonString = BuildJSON(produto)
' Output: {"id":123,"nome":"Notebook Dell","preco":2599.99,"disponivel":true,"descricao":null}

' Construir array aninhado
Dim pedido As Dictionary
Set pedido = CreateJSONObject()
pedido.Add "id", 456
pedido.Add "items", CreateJSONArray()

Dim item1 As Dictionary
Set item1 = CreateJSONObject()
item1.Add "produto_id", 123
item1.Add "quantidade", 2
pedido("items").Add item1

Dim item2 As Dictionary
Set item2 = CreateJSONObject()
item2.Add "produto_id", 124
item2.Add "quantidade", 1
pedido("items").Add item2

jsonString = BuildJSON(pedido)
' Output: {"id":456,"items":[{"produto_id":123,"quantidade":2},{"produto_id":124,"quantidade":1}]}
```

## Estruturas de Dados

### Mapeamento JSON ↔ VB6

| Tipo JSON | Tipo VB6 | Exemplo JSON | Exemplo VB6 |
|-----------|----------|--------------|-------------|
| `object` | `Dictionary` | `{"key":"value"}` | `dict("key") = "value"` |
| `array` | `Collection` | `[1,2,3]` | `coll.Add 1: coll.Add 2: coll.Add 3` |
| `string` | `String` | `"Hello"` | `"Hello"` |
| `number` | `Long/Double` | `42` ou `3.14` | `42` ou `3.14` |
| `boolean` | `Boolean` | `true/false` | `True/False` |
| `null` | `Null` | `null` | `Null` |

### CreateJSONObject

```vb
Public Function CreateJSONObject() As Dictionary
```

**Funcionalidade**: Cria um Dictionary configurado para uso como objeto JSON.

**Exemplo de Uso Avançado**:

```vb
' Construir estrutura hierárquica
Dim empresa As Dictionary
Set empresa = CreateJSONObject()
empresa.Add "nome", "TechCorp"
empresa.Add "fundacao", 2010

' Endereço aninhado
empresa.Add "endereco", CreateJSONObject()
empresa("endereco").Add "rua", "Rua das Flores, 123"
empresa("endereco").Add "cidade", "São Paulo"
empresa("endereco").Add "cep", "01234-567"

' Array de funcionários
empresa.Add "funcionarios", CreateJSONArray()

Dim funcionario1 As Dictionary
Set funcionario1 = CreateJSONObject()
funcionario1.Add "id", 1
funcionario1.Add "nome", "Maria Silva"
funcionario1.Add "cargo", "Desenvolvedora"
funcionario1.Add "salario", 8500.50
empresa("funcionarios").Add funcionario1

Dim funcionario2 As Dictionary
Set funcionario2 = CreateJSONObject()
funcionario2.Add "id", 2
funcionario2.Add "nome", "João Santos"
funcionario2.Add "cargo", "Analista"
funcionario2.Add "salario", 7200.00
empresa("funcionarios").Add funcionario2
```

### CreateJSONArray

```vb
Public Function CreateJSONArray() As Collection
```

**Funcionalidade**: Cria uma Collection configurada para uso como array JSON.

**Padrões de Uso**:

```vb
' Array de tipos mistos
Dim dadosMistos As Collection
Set dadosMistos = CreateJSONArray()
dadosMistos.Add "texto"
dadosMistos.Add 42
dadosMistos.Add True
dadosMistos.Add Null

' Array de objetos
Dim produtos As Collection
Set produtos = CreateJSONArray()

Dim i As Integer
For i = 1 To 3
    Dim produto As Dictionary
    Set produto = CreateJSONObject()
    produto.Add "id", i
    produto.Add "nome", "Produto " & i
    produto.Add "preco", i * 100
    produtos.Add produto
Next i

Dim jsonArray As String
jsonArray = BuildJSON(produtos)
' Output: [{"id":1,"nome":"Produto 1","preco":100},{"id":2,"nome":"Produto 2","preco":200},{"id":3,"nome":"Produto 3","preco":300}]
```

## Tipos Suportados

### Strings e Caracteres Especiais

**Caracteres de Escape Suportados**:

| Escape | Significado | Uso |
|--------|-------------|-----|
| `\"` | Aspas duplas | `"Ele disse: \"Olá\""` |
| `\\` | Barra invertida | `"C:\\Windows\\System32"` |
| `\/` | Barra normal | `"http:\/\/exemplo.com"` |
| `\b` | Backspace | Caractere de controle |
| `\f` | Form feed | Caractere de controle |
| `\n` | Nova linha | `"Linha 1\nLinha 2"` |
| `\r` | Carriage return | `"Texto\r\n"` |
| `\t` | Tab | `"Coluna1\tColuna2"` |
| `\uXXXX` | Unicode | `"Caf\u00e9"` (Café) |

**Exemplo de Processamento**:

```vb
' String com caracteres especiais
Dim textoComplexo As String
textoComplexo = "Ele disse: ""Olá!"" e foi para C:\Pasta\Arquivo.txt" & vbNewLine & "Nova linha aqui."

Dim obj As Dictionary
Set obj = CreateJSONObject()
obj.Add "mensagem", textoComplexo

Dim json As String
json = BuildJSON(obj)
' Output: {"mensagem":"Ele disse: \"Olá!\" e foi para C:\\Pasta\\Arquivo.txt\nNova linha aqui."}

' Parse de volta
Dim parsed As Object
Set parsed = ParseJSON(json)
Debug.Print parsed("mensagem")  ' Texto original restaurado
```

### Números

**Formatos Suportados**:

- Inteiros: `42`, `-17`, `0`
- Decimais: `3.14`, `-0.5`, `123.456`
- Científicos: `1.23e10`, `4.56E-7`, `-2.1e+3`

**Lógica de Conversão**:

```vb
' No parsing: determina Long ou Double
If InStr(numStr, ".") > 0 Or InStr(numStr, "e") > 0 Or InStr(numStr, "E") > 0 Then
    ParseNumber = CDbl(numStr)    ' Double para decimais/científicos
Else
    ParseNumber = CLng(numStr)    ' Long para inteiros
End If

' Na geração: formato americano
If VarType(num) = vbSingle Or VarType(num) = vbDouble Then
    BuildNumber = Replace(CStr(num), ",", ".")  ' Força ponto decimal
Else
    BuildNumber = CStr(num)
End If
```

### Valores Nulos e Booleanos

```vb
' Tratamento de valores especiais
Dim testValues As Dictionary
Set testValues = CreateJSONObject()
testValues.Add "ativo", True        ' → "ativo":true
testValues.Add "inativo", False     ' → "inativo":false
testValues.Add "indefinido", Null   ' → "indefinido":null
testValues.Add "vazio", ""          ' → "vazio":""
testValues.Add "zero", 0            ' → "zero":0

' JSON resultante:
' {"ativo":true,"inativo":false,"indefinido":null,"vazio":"","zero":0}
```

## Tratamento de Erros

### Códigos de Erro

```vb
' Códigos de erro do parser (vbObjectError + código):
Const JSON_INVALID_ROOT = 1         ' String deve começar com '{' ou '['
Const JSON_PROPERTY_EXPECTED = 2    ' Nome de propriedade esperado
Const JSON_COLON_EXPECTED = 3       ' ':' esperado após nome de propriedade
Const JSON_COMMA_OR_END_EXPECTED = 4 ' ',' ou '}' esperado
Const JSON_ARRAY_COMMA_EXPECTED = 5  ' ',' ou ']' esperado
Const JSON_INVALID_VALUE = 6         ' Valor JSON inválido
Const JSON_INVALID_ESCAPE = 7        ' Sequência de escape inválida
Const JSON_UNTERMINATED_STRING = 8   ' String não terminada
Const JSON_INVALID_LITERAL = 9       ' Literal inválido (true/false/null)
Const JSON_UNSUPPORTED_TYPE = 20     ' Tipo de objeto não suportado na geração
```

### Estratégias de Tratamento

```vb
' Função robusta de parsing
Function SafeParseJSON(jsonString As String) As Object
    On Error GoTo ErrorHandler

    Set SafeParseJSON = ParseJSON(jsonString)
    Exit Function

ErrorHandler:
    Dim errorCode As Long
    errorCode = Err.Number - vbObjectError

    Select Case errorCode
        Case 1 To 11:  ' Erros de parsing
            LogParseError "JSON inválido: " & Err.Description
            Set SafeParseJSON = Nothing
        Case 20:       ' Tipo não suportado
            LogBuildError "Tipo de objeto não suportado: " & Err.Description
            Set SafeParseJSON = Nothing
        Case Else:     ' Outros erros
            LogGeneralError "Erro inesperado: " & Err.Description
            Set SafeParseJSON = Nothing
    End Select
End Function

' Função robusta de geração
Function SafeBuildJSON(obj As Variant) As String
    On Error GoTo ErrorHandler

    SafeBuildJSON = BuildJSON(obj)
    Exit Function

ErrorHandler:
    LogBuildError "Erro ao gerar JSON: " & Err.Description
    SafeBuildJSON = "{""error"":""JSON generation failed""}"
End Function
```

## Casos de Uso Avançados

### Serialização de Classes VB6

```vb
' Classe Pessoa
Public Class CPessoa
    Public Nome As String
    Public Idade As Integer
    Public Email As String
    Public Ativo As Boolean

    Public Function ToJSON() As Dictionary
        Set ToJSON = CreateJSONObject()
        ToJSON.Add "nome", Me.Nome
        ToJSON.Add "idade", Me.Idade
        ToJSON.Add "email", Me.Email
        ToJSON.Add "ativo", Me.Ativo
    End Function

    Public Sub FromJSON(jsonObj As Dictionary)
        Me.Nome = jsonObj("nome")
        Me.Idade = jsonObj("idade")
        Me.Email = jsonObj("email")
        Me.Ativo = jsonObj("ativo")
    End Sub
End Class

' Uso da serialização
Dim pessoa As New CPessoa
pessoa.Nome = "Ana Silva"
pessoa.Idade = 28
pessoa.Email = "ana@email.com"
pessoa.Ativo = True

Dim jsonString As String
jsonString = BuildJSON(pessoa.ToJSON())

' Deserialização
Dim novaPessoa As New CPessoa
novaPessoa.FromJSON ParseJSON(jsonString)
```

### Validação de Schema

```vb
' Validador simples de estrutura JSON
Function ValidateUserJSON(jsonObj As Object) As Boolean
    On Error GoTo ErrorHandler

    ' Verificar se é um Dictionary
    If TypeName(jsonObj) <> "Dictionary" Then
        ValidateUserJSON = False
        Exit Function
    End If

    ' Verificar campos obrigatórios
    If Not jsonObj.Exists("nome") Or Not jsonObj.Exists("email") Then
        ValidateUserJSON = False
        Exit Function
    End If

    ' Verificar tipos
    If VarType(jsonObj("nome")) <> vbString Or VarType(jsonObj("email")) <> vbString Then
        ValidateUserJSON = False
        Exit Function
    End If

    ' Validação de email básica
    If InStr(jsonObj("email"), "@") = 0 Then
        ValidateUserJSON = False
        Exit Function
    End If

    ValidateUserJSON = True
    Exit Function

ErrorHandler:
    ValidateUserJSON = False
End Function
```

### Transformação de Dados

```vb
' Função para transformar estrutura de dados
Function TransformAPIResponse(apiResponse As Object) As Dictionary
    Dim transformed As Dictionary
    Set transformed = CreateJSONObject()

    ' Extrair dados relevantes
    If apiResponse.Exists("data") Then
        Dim dataArray As Collection
        Set dataArray = apiResponse("data")

        transformed.Add "items", CreateJSONArray()
        transformed.Add "total", dataArray.Count

        Dim i As Integer
        For i = 1 To dataArray.Count
            Dim item As Dictionary
            Set item = dataArray(i)

            Dim transformedItem As Dictionary
            Set transformedItem = CreateJSONObject()
            transformedItem.Add "id", item("id")
            transformedItem.Add "title", item("nome")
            transformedItem.Add "value", item("preco")

            transformed("items").Add transformedItem
        Next i
    End If

    Set TransformAPIResponse = transformed
End Function
```

---

**🔧 Dica Técnica**: O JsonHelper é completamente thread-safe quando usado corretamente, pois não mantém estado global além do parser temporário. Cada chamada para ParseJSON ou BuildJSON é independente.

**⚡ Performance**: Para grandes volumes de dados, considere processar em lotes menores para evitar timeouts ou problemas de memória no VB6.
