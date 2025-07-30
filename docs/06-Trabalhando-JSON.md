# Trabalhando com JSON - Guia Completo

Este guia aborda todos os aspectos do trabalho com JSON no VB6, desde conceitos básicos até técnicas avançadas de manipulação de dados.

## 📋 Índice

- [Fundamentos JSON](#fundamentos-json)
- [Criação de Estruturas JSON](#criação-de-estruturas-json)
- [Parsing e Conversão](#parsing-e-conversão)
- [Manipulação de Dados](#manipulação-de-dados)
- [Padrões Avançados](#padrões-avançados)
- [Performance e Otimização](#performance-e-otimização)

## Fundamentos JSON

### Mapeamento JSON ↔ VB6

```vb
' Correspondência entre tipos JSON e VB6:

' JSON Object {} → VB6 Dictionary
Dim usuario As Dictionary
Set usuario = CreateJSONObject()
usuario.Add "nome", "João"
usuario.Add "idade", 30

' JSON Array [] → VB6 Collection
Dim cores As Collection
Set cores = CreateJSONArray()
cores.Add "vermelho"
cores.Add "verde"
cores.Add "azul"

' JSON String → VB6 String
Dim texto As String
texto = "Olá mundo"

' JSON Number → VB6 Long/Double
Dim inteiro As Long
inteiro = 42
Dim decimal As Double
decimal = 3.14159

' JSON Boolean → VB6 Boolean
Dim ativo As Boolean
ativo = True

' JSON null → VB6 Null
Dim nulo As Variant
nulo = Null
```

### Sintaxe JSON Válida

```json
{
  "string": "texto",
  "number": 123,
  "decimal": 45.67,
  "boolean": true,
  "null": null,
  "array": [1, 2, 3],
  "object": {
    "nested": "valor"
  }
}
```

## Criação de Estruturas JSON

### Objetos Simples

```vb
' Criar objeto básico:
Function CriarUsuario(nome As String, idade As Integer, email As String) As Dictionary
    Dim usuario As Dictionary
    Set usuario = CreateJSONObject()

    usuario.Add "nome", nome
    usuario.Add "idade", idade
    usuario.Add "email", email
    usuario.Add "ativo", True
    usuario.Add "ultimo_login", Null

    Set CriarUsuario = usuario
End Function

' Uso:
Dim joao As Dictionary
Set joao = CriarUsuario("João Silva", 35, "joao@email.com")

Dim json As String
json = BuildJSON(joao)
' Resultado: {"nome":"João Silva","idade":35,"email":"joao@email.com","ativo":true,"ultimo_login":null}
```

### Arrays e Listas

```vb
' Criar array de strings:
Function CriarListaCores() As Collection
    Dim cores As Collection
    Set cores = CreateJSONArray()

    cores.Add "vermelho"
    cores.Add "verde"
    cores.Add "azul"
    cores.Add "amarelo"

    Set CriarListaCores = cores
End Function

' Criar array de números:
Function CriarListaNumeros() As Collection
    Dim numeros As Collection
    Set numeros = CreateJSONArray()

    Dim i As Integer
    For i = 1 To 10
        numeros.Add i * i  ' Quadrados perfeitos
    Next i

    Set CriarListaNumeros = numeros
End Function

' Array misto (diferentes tipos):
Function CriarArrayMisto() As Collection
    Dim misto As Collection
    Set misto = CreateJSONArray()

    misto.Add "texto"
    misto.Add 42
    misto.Add True
    misto.Add Null
    misto.Add CreateJSONObject()  ' Objeto vazio

    Set CriarArrayMisto = misto
End Function
```

### Estruturas Aninhadas

```vb
' Criar estrutura complexa com objetos aninhados:
Function CriarEmpresa() As Dictionary
    Dim empresa As Dictionary
    Set empresa = CreateJSONObject()

    ' Dados básicos
    empresa.Add "nome", "TechCorp Ltda"
    empresa.Add "cnpj", "12.345.678/0001-90"
    empresa.Add "fundacao", 2010

    ' Endereço (objeto aninhado)
    Dim endereco As Dictionary
    Set endereco = CreateJSONObject()
    endereco.Add "rua", "Av. Paulista, 1000"
    endereco.Add "cidade", "São Paulo"
    endereco.Add "estado", "SP"
    endereco.Add "cep", "01310-100"
    empresa.Add "endereco", endereco

    ' Contatos (array de objetos)
    Dim contatos As Collection
    Set contatos = CreateJSONArray()

    Dim email As Dictionary
    Set email = CreateJSONObject()
    email.Add "tipo", "email"
    email.Add "valor", "contato@techcorp.com"
    contatos.Add email

    Dim telefone As Dictionary
    Set telefone = CreateJSONObject()
    telefone.Add "tipo", "telefone"
    telefone.Add "valor", "(11) 3333-4444"
    contatos.Add telefone

    empresa.Add "contatos", contatos

    ' Funcionários (array de objetos complexos)
    empresa.Add "funcionarios", CriarFuncionarios()

    Set CriarEmpresa = empresa
End Function

Private Function CriarFuncionarios() As Collection
    Dim funcionarios As Collection
    Set funcionarios = CreateJSONArray()

    ' Funcionário 1
    Dim func1 As Dictionary
    Set func1 = CreateJSONObject()
    func1.Add "id", 1
    func1.Add "nome", "Maria Santos"
    func1.Add "cargo", "Desenvolvedora Senior"
    func1.Add "salario", 12000.50
    func1.Add "beneficios", CreateJSONArray()
    func1("beneficios").Add "Vale Refeição"
    func1("beneficios").Add "Plano de Saúde"
    func1("beneficios").Add "Gympass"
    funcionarios.Add func1

    ' Funcionário 2
    Dim func2 As Dictionary
    Set func2 = CreateJSONObject()
    func2.Add "id", 2
    func2.Add "nome", "Pedro Lima"
    func2.Add "cargo", "Product Manager"
    func2.Add "salario", 15000.00
    func2.Add "beneficios", CreateJSONArray()
    func2("beneficios").Add "Vale Refeição"
    func2("beneficios").Add "Plano de Saúde"
    funcionarios.Add func2

    Set CriarFuncionarios = funcionarios
End Function
```

## Parsing e Conversão

### Parse de JSON String

```vb
' Parse básico:
Sub ExemploParseBasico()
    Dim jsonString As String
    jsonString = "{""nome"":""Ana"",""idade"":28,""ativo"":true}"

    Dim usuario As Object
    Set usuario = ParseJSON(jsonString)

    Debug.Print "Nome: " & usuario("nome")
    Debug.Print "Idade: " & usuario("idade")
    Debug.Print "Ativo: " & usuario("ativo")
End Sub

' Parse de array:
Sub ExemploParseArray()
    Dim jsonArray As String
    jsonArray = "[""maçã"",""banana"",""laranja""]"

    Dim frutas As Object
    Set frutas = ParseJSON(jsonArray)

    Dim i As Integer
    For i = 1 To frutas.Count
        Debug.Print "Fruta " & i & ": " & frutas(i)
    Next i
End Sub

' Parse com tratamento de erro:
Function SafeParseJSON(jsonString As String) As Object
    On Error GoTo ParseError

    Set SafeParseJSON = ParseJSON(jsonString)
    Exit Function

ParseError:
    Debug.Print "Erro ao parsear JSON: " & Err.Description
    Debug.Print "JSON inválido: " & Left(jsonString, 100)
    Set SafeParseJSON = Nothing
End Function
```

### Conversão para JSON String

```vb
' Conversão básica:
Sub ExemploConversaoBasica()
    Dim dados As Dictionary
    Set dados = CreateJSONObject()
    dados.Add "produto", "Notebook"
    dados.Add "preco", 2500.99
    dados.Add "disponivel", True

    Dim json As String
    json = BuildJSON(dados)

    Debug.Print json
    ' Output: {"produto":"Notebook","preco":2500.99,"disponivel":true}
End Sub

' Conversão com formatação (função auxiliar):
Function BuildJSONFormatted(obj As Variant) As String
    ' Esta é uma versão simplificada - a biblioteca não inclui formatação
    ' mas você pode implementar uma função auxiliar para debug

    Dim json As String
    json = BuildJSON(obj)

    ' Adicionar quebras de linha após vírgulas (simplificado)
    json = Replace(json, ",", "," & vbCrLf & "  ")
    json = Replace(json, "{", "{" & vbCrLf & "  ")
    json = Replace(json, "}", vbCrLf & "}")

    BuildJSONFormatted = json
End Function
```

## Manipulação de Dados

### Acessando Propriedades

```vb
' Acesso direto a propriedades:
Sub AcessarPropriedades()
    Dim jsonData As String
    jsonData = "{""usuario"":{""nome"":""Carlos"",""perfil"":{""nivel"":""admin"",""permissoes"":[""ler"",""escrever"",""deletar""]}}}"

    Dim data As Object
    Set data = ParseJSON(jsonData)

    ' Acesso aninhado
    Debug.Print "Nome: " & data("usuario")("nome")
    Debug.Print "Nível: " & data("usuario")("perfil")("nivel")

    ' Acesso a array
    Dim permissoes As Collection
    Set permissoes = data("usuario")("perfil")("permissoes")

    Dim i As Integer
    For i = 1 To permissoes.Count
        Debug.Print "Permissão " & i & ": " & permissoes(i)
    Next i
End Sub

' Acesso seguro com verificação:
Function GetNestedValue(obj As Object, path As String, Optional defaultValue As Variant = "") As Variant
    On Error GoTo NotFound

    Dim keys() As String
    keys = Split(path, ".")

    Dim current As Object
    Set current = obj

    Dim i As Integer
    For i = 0 To UBound(keys)
        If TypeName(current) = "Dictionary" Then
            If current.Exists(keys(i)) Then
                If IsObject(current(keys(i))) Then
                    Set current = current(keys(i))
                Else
                    GetNestedValue = current(keys(i))
                    Exit Function
                End If
            Else
                GoTo NotFound
            End If
        Else
            GoTo NotFound
        End If
    Next i

    Set GetNestedValue = current
    Exit Function

NotFound:
    GetNestedValue = defaultValue
End Function

' Uso do acesso seguro:
Sub ExemploAcessoSeguro()
    Dim data As Object
    Set data = ParseJSON("{""usuario"":{""nome"":""João""}}")

    ' Estes acessos não geram erro mesmo se o path não existir:
    Debug.Print GetNestedValue(data, "usuario.nome", "N/A")           ' Output: João
    Debug.Print GetNestedValue(data, "usuario.idade", "N/A")          ' Output: N/A
    Debug.Print GetNestedValue(data, "empresa.nome", "N/A")           ' Output: N/A
End Sub
```

### Modificando Estruturas JSON

```vb
' Adicionar propriedades:
Sub AdicionarPropriedades()
    Dim usuario As Dictionary
    Set usuario = CreateJSONObject()
    usuario.Add "nome", "Ana"
    usuario.Add "idade", 25

    ' Adicionar nova propriedade
    usuario.Add "email", "ana@email.com"

    ' Adicionar objeto aninhado
    usuario.Add "endereco", CreateJSONObject()
    usuario("endereco").Add "cidade", "Rio de Janeiro"
    usuario("endereco").Add "estado", "RJ"

    Debug.Print BuildJSON(usuario)
End Sub

' Remover propriedades:
Sub RemoverPropriedades()
    Dim data As Dictionary
    Set data = ParseJSON("{""nome"":""João"",""idade"":30,""temp"":""remover""}")

    ' Remover propriedade
    If data.Exists("temp") Then
        data.Remove "temp"
    End If

    Debug.Print BuildJSON(data)
    ' Output: {"nome":"João","idade":30}
End Sub

' Modificar valores:
Sub ModificarValores()
    Dim produto As Dictionary
    Set produto = CreateJSONObject()
    produto.Add "nome", "Notebook"
    produto.Add "preco", 2000.00
    produto.Add "desconto", 0

    ' Atualizar valores
    produto("preco") = 1800.00  ' Aplicar desconto
    produto("desconto") = 10    ' 10% de desconto
    produto.Add "promocao", True ' Nova propriedade

    Debug.Print BuildJSON(produto)
End Sub
```

### Trabalhando com Arrays

```vb
' Adicionar itens a array:
Sub AdicionarItensArray()
    Dim lista As Collection
    Set lista = CreateJSONArray()

    ' Adicionar um por vez
    lista.Add "Item 1"
    lista.Add "Item 2"
    lista.Add "Item 3"

    ' Adicionar múltiplos de uma vez
    Dim novosItens As Variant
    novosItens = Array("Item 4", "Item 5", "Item 6")

    Dim i As Integer
    For i = 0 To UBound(novosItens)
        lista.Add novosItens(i)
    Next i

    Debug.Print BuildJSON(lista)
End Sub

' Filtrar array:
Function FiltrarArray(arr As Collection, criterio As String) As Collection
    Dim filtrado As Collection
    Set filtrado = CreateJSONArray()

    Dim i As Integer
    For i = 1 To arr.Count
        ' Exemplo: filtrar strings que contêm o critério
        If TypeName(arr(i)) = "String" Then
            If InStr(LCase(arr(i)), LCase(criterio)) > 0 Then
                filtrado.Add arr(i)
            End If
        End If
    Next i

    Set FiltrarArray = filtrado
End Function

' Transformar array:
Function TransformarArray(arr As Collection, transformacao As String) As Collection
    Dim transformado As Collection
    Set transformado = CreateJSONArray()

    Dim i As Integer
    For i = 1 To arr.Count
        Select Case transformacao
            Case "uppercase":
                If TypeName(arr(i)) = "String" Then
                    transformado.Add UCase(arr(i))
                Else
                    transformado.Add arr(i)
                End If
            Case "dobrar":
                If IsNumeric(arr(i)) Then
                    transformado.Add arr(i) * 2
                Else
                    transformado.Add arr(i)
                End If
            Case Else:
                transformado.Add arr(i)
        End Select
    Next i

    Set TransformarArray = transformado
End Function
```

## Padrões Avançados

### Serialização de Classes

```vb
' Classe exemplo:
' (Criar como arquivo .cls separado)
Public Class CProduto
    Public ID As Long
    Public Nome As String
    Public Preco As Double
    Public Categoria As String
    Public Ativo As Boolean

    Public Function ToJSON() As Dictionary
        Dim json As Dictionary
        Set json = CreateJSONObject()

        json.Add "id", Me.ID
        json.Add "nome", Me.Nome
        json.Add "preco", Me.Preco
        json.Add "categoria", Me.Categoria
        json.Add "ativo", Me.Ativo

        Set ToJSON = json
    End Function

    Public Sub FromJSON(jsonObj As Dictionary)
        If jsonObj.Exists("id") Then Me.ID = jsonObj("id")
        If jsonObj.Exists("nome") Then Me.Nome = jsonObj("nome")
        If jsonObj.Exists("preco") Then Me.Preco = jsonObj("preco")
        If jsonObj.Exists("categoria") Then Me.Categoria = jsonObj("categoria")
        If jsonObj.Exists("ativo") Then Me.Ativo = jsonObj("ativo")
    End Sub
End Class

' Uso da serialização:
Sub ExemploSerializacao()
    ' Criar objeto
    Dim produto As New CProduto
    produto.ID = 123
    produto.Nome = "Smartphone"
    produto.Preco = 899.99
    produto.Categoria = "Eletrônicos"
    produto.Ativo = True

    ' Serializar para JSON
    Dim json As String
    json = BuildJSON(produto.ToJSON())
    Debug.Print "Serializado: " & json

    ' Deserializar de JSON
    Dim novoProduto As New CProduto
    novoProduto.FromJSON ParseJSON(json)

    Debug.Print "Deserializado: " & novoProduto.Nome & " - R$ " & novoProduto.Preco
End Sub
```

### Validação de Schema

```vb
' Validador de estrutura JSON:
Function ValidarEstrutura(jsonObj As Object, schema As Dictionary) As Boolean
    On Error GoTo ValidacaoError

    If TypeName(jsonObj) <> "Dictionary" Then
        ValidarEstrutura = False
        Exit Function
    End If

    Dim campo As Variant
    For Each campo In schema.Keys
        Dim regra As Dictionary
        Set regra = schema(campo)

        ' Verificar se campo existe
        If regra("obrigatorio") And Not jsonObj.Exists(campo) Then
            Debug.Print "Campo obrigatório ausente: " & campo
            ValidarEstrutura = False
            Exit Function
        End If

        ' Verificar tipo se campo existe
        If jsonObj.Exists(campo) Then
            Dim tipoEsperado As String
            tipoEsperado = regra("tipo")

            Select Case tipoEsperado
                Case "string":
                    If VarType(jsonObj(campo)) <> vbString Then
                        Debug.Print "Tipo inválido para " & campo & ": esperado string"
                        ValidarEstrutura = False
                        Exit Function
                    End If
                Case "number":
                    If Not IsNumeric(jsonObj(campo)) Then
                        Debug.Print "Tipo inválido para " & campo & ": esperado número"
                        ValidarEstrutura = False
                        Exit Function
                    End If
                Case "boolean":
                    If VarType(jsonObj(campo)) <> vbBoolean Then
                        Debug.Print "Tipo inválido para " & campo & ": esperado boolean"
                        ValidarEstrutura = False
                        Exit Function
                    End If
            End Select
        End If
    Next campo

    ValidarEstrutura = True
    Exit Function

ValidacaoError:
    Debug.Print "Erro na validação: " & Err.Description
    ValidarEstrutura = False
End Function

' Criar schema de validação:
Function CriarSchemaUsuario() As Dictionary
    Dim schema As Dictionary
    Set schema = CreateJSONObject()

    ' Campo nome
    Dim campoNome As Dictionary
    Set campoNome = CreateJSONObject()
    campoNome.Add "tipo", "string"
    campoNome.Add "obrigatorio", True
    schema.Add "nome", campoNome

    ' Campo idade
    Dim campoIdade As Dictionary
    Set campoIdade = CreateJSONObject()
    campoIdade.Add "tipo", "number"
    campoIdade.Add "obrigatorio", True
    schema.Add "idade", campoIdade

    ' Campo email
    Dim campoEmail As Dictionary
    Set campoEmail = CreateJSONObject()
    campoEmail.Add "tipo", "string"
    campoEmail.Add "obrigatorio", False
    schema.Add "email", campoEmail

    Set CriarSchemaUsuario = schema
End Function

' Uso da validação:
Sub ExemploValidacao()
    Dim schema As Dictionary
    Set schema = CriarSchemaUsuario()

    ' Dados válidos
    Dim usuarioValido As Dictionary
    Set usuarioValido = CreateJSONObject()
    usuarioValido.Add "nome", "João"
    usuarioValido.Add "idade", 30
    usuarioValido.Add "email", "joao@email.com"

    If ValidarEstrutura(usuarioValido, schema) Then
        Debug.Print "Usuário válido"
    Else
        Debug.Print "Usuário inválido"
    End If

    ' Dados inválidos
    Dim usuarioInvalido As Dictionary
    Set usuarioInvalido = CreateJSONObject()
    usuarioInvalido.Add "nome", 123  ' Tipo errado
    ' usuarioInvalido.Add "idade", 30  ' Campo obrigatório ausente

    If ValidarEstrutura(usuarioInvalido, schema) Then
        Debug.Print "Usuário válido"
    Else
        Debug.Print "Usuário inválido"
    End If
End Sub
```

## Performance e Otimização

### Lazy Loading para Estruturas Grandes

```vb
' Classe para lazy loading de dados JSON grandes:
Public Class CJSONLazyLoader
    Private m_JsonString As String
    Private m_ParsedData As Object
    Private m_IsParsed As Boolean

    Public Sub Initialize(jsonString As String)
        m_JsonString = jsonString
        m_IsParsed = False
        Set m_ParsedData = Nothing
    End Sub

    Public Property Get Data() As Object
        If Not m_IsParsed Then
            Set m_ParsedData = ParseJSON(m_JsonString)
            m_IsParsed = True
        End If
        Set Data = m_ParsedData
    End Property

    Public Property Get Size() As Long
        Size = Len(m_JsonString)
    End Property
End Class
```

### Cache de Conversões

```vb
' Sistema de cache para conversões frequentes:
Private m_CacheJSON As Dictionary

Sub InicializarCache()
    Set m_CacheJSON = CreateJSONObject()
End Sub

Function BuildJSONWithCache(obj As Object, cacheKey As String) As String
    ' Verificar cache
    If Not m_CacheJSON Is Nothing Then
        If m_CacheJSON.Exists(cacheKey) Then
            BuildJSONWithCache = m_CacheJSON(cacheKey)
            Exit Function
        End If
    End If

    ' Gerar JSON
    Dim json As String
    json = BuildJSON(obj)

    ' Salvar no cache
    If Not m_CacheJSON Is Nothing Then
        m_CacheJSON(cacheKey) = json
    End If

    BuildJSONWithCache = json
End Function

Sub LimparCache()
    If Not m_CacheJSON Is Nothing Then
        m_CacheJSON.RemoveAll
    End If
End Sub
```

---

**💡 Dica de Performance**: Para estruturas JSON muito grandes (>1MB), considere processar em partes menores ou usar lazy loading para evitar problemas de memória no VB6.

**🔧 Debugging**: Use a função `BuildJSONFormatted` durante desenvolvimento para visualizar melhor a estrutura dos objetos JSON complexos.
