<div id="top">

<div align="center">

# VB6 API Consumption

Sistema completo para consumo de APIs REST em Visual Basic 6.0 com suporte nativo a JSON.

[![Version](https://img.shields.io/badge/version-1.0.0-blue?style=flat-square)](CHANGELOG.md)
[![Status](https://img.shields.io/badge/status-Stable-brightgreen?style=flat-square)](#)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)

**Tecnologias Utilizadas:**

<a href="https://docs.microsoft.com/en-us/previous-versions/visual-studio/"><img alt="VB6" src="https://img.shields.io/badge/Visual%20Basic-6.0-blue?style=flat-square&logo=microsoft&logoColor=white"></a>
<a href="#"><img alt="JSON" src="https://img.shields.io/badge/JSON-Native-orange?style=flat-square&logo=json&logoColor=white"></a>
<a href="#"><img alt="HTTP" src="https://img.shields.io/badge/HTTP-REST-green?style=flat-square&logo=http&logoColor=white"></a>
<a href="#"><img alt="XML" src="https://img.shields.io/badge/XML-HTTP-red?style=flat-square&logo=xml&logoColor=white"></a>

> 🎉 **Versão Final Estável** - Sistema completo, testado e pronto para uso em produção!

</div>

<br>
<hr>

## 📋 Tabela de Conteúdos

- [🚀 Visão Geral](#-visão-geral)
- [✨ Funcionalidades](#-funcionalidades)
- [🛠️ Tecnologias](#️-tecnologias)
- [📋 Requisitos](#-requisitos)
- [🔧 Instalação e Configuração](#-instalação-e-configuração)
- [📚 Guia de Uso](#-guia-de-uso)
- [🧪 Testes e Validação](#-testes-e-validação)
- [📖 Documentação da API](#-documentação-da-api)
- [🎯 Exemplos Práticos](#-exemplos-práticos)
- [🔧 Troubleshooting](#-troubleshooting)
- [🗺️ Roadmap](#️-roadmap)
- [🤝 Contribuindo](#-contribuindo)
- [📜 Licença](#-licença)

<hr>

## 🚀 Visão Geral

O **VB6 API Consumption** é uma biblioteca para integração de APIs REST em aplicações Visual Basic 6.0. A solução implementa um cliente HTTP completo com suporte nativo a JSON, eliminando a necessidade de componentes externos ou bibliotecas de terceiros.

### Características Técnicas

- **Implementação Nativa**: Utiliza apenas recursos padrão do VB6 e componentes do sistema Windows
- **Cliente HTTP Completo**: Suporte aos métodos HTTP padrão (GET, POST, PUT, DELETE, PATCH)
- **Parser JSON**: Engine de parsing e geração JSON implementado nativamente
- **Arquitetura Modular**: Componentes independentes e reutilizáveis
- **Compatibilidade**: Funciona com Windows 7+ e todas as versões do VB6

### Estrutura do Projeto

``` bash
src/
├── Modules/
│   ├── HttpClient.bas      # Cliente HTTP principal
│   ├── JsonHelper.bas      # Processamento JSON
├── Classes/
│   ├── HttpResponse.cls    # Objeto de resposta HTTP
│   └── cHttpRequest.cls    # Wrapper XMLHTTP
└── ConsumoAPI.vbp          # Projeto VB6
```

## ✨ Funcionalidades

### Core Features

- 🌐 **Cliente HTTP Completo** - Suporte a GET, POST, PUT, DELETE, PATCH
- 📄 **JSON Nativo** - Parser e gerador JSON sem dependências externas
- 🔧 **Headers Configuráveis** - Sistema completo de gerenciamento de headers
- ⚡ **Timeout Configurável** - Controle preciso de timeouts de requisição
- 🔐 **Suporte a Autenticação** - Bearer tokens, API keys e headers customizados
- 🛡️ **Tratamento de Erros** - Sistema robusto de tratamento de erros HTTP
- 📝 **URL Encoding** - Codificação automática de URLs e parâmetros

### Features Avançadas

- 🎯 **Respostas Tipadas** - Classe HttpResponse com propriedades estruturadas
- 🔄 **Retry Logic** - Mecanismo de retry para requisições falhadas
- 📊 **Logging Integrado** - Sistema de logs para debug e monitoramento
- 📚 **Documentação Completa** - Docstrings simples e completos para cada funcionalidade

## 🛠️ Tecnologias

- **Linguagem**: Visual Basic 6.0
- **HTTP Client**: Microsoft XML HTTP Services (XMLHTTP)
- **JSON Processing**: Microsoft Scripting Runtime (Dictionary)
- **Encoding**: Nativo VB6

## 📋 Requisitos

### Sistema Operacional

- Windows 7 ou superior
- Visual Basic 6.0 IDE (para desenvolvimento)
- VB6 Runtime (para execução)

### Dependências Obrigatórias

- **Microsoft Scripting Runtime** (scrrun.dll)
- **Microsoft XML HTTP Services** (msxml6.dll ou msxml3.dll)

### Configurações Mínimas

- RAM: 512MB (recomendado: 1GB+)
- Espaço em disco: 10MB
- Conexão com internet (para consumo de APIs)

## 🔧 Instalação e Configuração

### 1. Download do Projeto

```bash
# Clone via Git
git clone https://github.com/seuusuario/VB6__API_consumption.git

# Ou baixe o ZIP diretamente do GitHub
```

### 2. Configuração de Referências

**⚠️ IMPORTANTE**: Configure as referências antes de usar o projeto.

1. Abra o Visual Basic 6.0
2. Vá em **Project → References**
3. Marque as seguintes referências:

```
☑️ Microsoft Scripting Runtime
   Localização: C:\Windows\System32\scrrun.dll

☑️ Microsoft XML, v6.0 (preferencial)
   Localização: C:\Windows\System32\msxml6.dll

   OU (caso v6.0 não esteja disponível)

☑️ Microsoft XML, v3.0
   Localização: C:\Windows\System32\msxml3.dll
```

### 3. Importação dos Arquivos

**Opção A: Projeto Novo**

1. Crie um novo projeto Standard EXE no VB6
2. Importe os arquivos usando **Project → Add Module/Class Module**:

```
📁 Modules/
├── HttpClient.bas      ← Add Module
├── JsonHelper.bas      ← Add Module

📁 Classes/
├── HttpResponse.cls    ← Add Class Module
└── cHttpRequest.cls    ← Add Class Module
```

**Opção B: Projeto Existente**

1. Abra `ConsumoAPI.vbp` no VB6
2. Execute os testes para validar a instalação
3. Copie os módulos necessários para seu projeto

## 📚 Guia de Uso

### Inicialização Básica

```vb
' Configuração inicial (execute uma vez no início da aplicação)
InitializeHttpClient "https://api.exemplo.com", 15000, "MeuApp/1.0"

' Headers padrão (opcional)
SetDefaultHeader "Authorization", "Bearer [seu-token]"
SetDefaultHeader "Content-Type", "application/json"
```

### Requisições GET

```vb
' GET simples com resposta em texto
Dim response As HttpResponse
Set response = HttpGet("/users/1")

If response.IsSuccess Then
    Debug.Print "Resposta: " & response.Text
End If

' GET com parsing automático para JSON
Dim user As Object
Set user = GetJson("/users/1")

If Not user Is Nothing Then
    Debug.Print "Nome: " & user("name")
    Debug.Print "Email: " & user("email")
End If
```

### Requisições POST

```vb
' POST com objeto JSON
Dim userData As Dictionary
Set userData = CreateJSONObject()
userData.Add "name", "João Silva"
userData.Add "email", "joao@email.com"
userData.Add "active", True

' Enviar e receber resposta como JSON
Dim newUser As Object
Set newUser = PostJson("/users", userData)

If Not newUser Is Nothing Then
    Debug.Print "Usuário criado com ID: " & newUser("id")
End If
```

### Trabalhando com JSON

```vb
' Criar objeto JSON
Dim produto As Dictionary
Set produto = CreateJSONObject()
produto.Add "nome", "Notebook"
produto.Add "preco", 2500.99
produto.Add "disponivel", True

' Array JSON
Dim categorias As Collection
Set categorias = New Collection
categorias.Add "eletrônicos"
categorias.Add "informática"
produto.Add "categorias", categorias

' Converter para string JSON
Dim jsonString As String
jsonString = BuildJSON(produto)
Debug.Print jsonString

' Parse de JSON string
Dim parsed As Object
Set parsed = ParseJSON(jsonString)
Debug.Print "Produto: " & parsed("nome")
```

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

## 🗺️ Roadmap

### ✅ Concluído (v1.0.0)

- [x] Cliente HTTP completo com todos os métodos
- [x] Parser e gerador JSON nativo
- [x] Sistema de headers configuráveis
- [x] Tratamento de erros robusto
- [x] Documentação completa

### 📅 Planejado

- [ ] Suite de testes automatizada
- [ ] Upload de arquivos (multipart/form-data)
- [ ] Suporte a cookies e sessões
- [ ] Sistema de cache de requisições
- [ ] Retry automático configurável
- [ ] Logging avançado com níveis
- [ ] Suporte a WebSockets básico
- [ ] Compressão GZIP automática
- [ ] Pool de conexões
- [ ] Suporte a OAuth 2.0 completo
- [ ] Async requests (limitado)

## 🤝 Contribuindo

Contribuições são bem-vindas! Este projeto segue as melhores práticas de desenvolvimento colaborativo.

### Como Contribuir

1. **Fork** o repositório
2. **Clone** seu fork: `git clone https://github.com/seuusuario/VB6__API_consumption.git`
3. **Crie** uma branch: `git checkout -b feature/nova-funcionalidade`
4. **Desenvolva** e **teste** suas mudanças
5. **Commit**: `git commit -m 'feat: adiciona nova funcionalidade X'`
6. **Push**: `git push origin feature/nova-funcionalidade`
7. **Abra** um Pull Request

### Padrões de Commit

Seguimos o padrão [Conventional Commits](https://www.conventionalcommits.org/):

```git
feat: nova funcionalidade
fix: correção de bug
docs: atualização de documentação
style: formatação de código
refactor: refatoração sem mudança de funcionalidade
test: adição ou correção de testes
chore: tarefas de manutenção
```

## 📜 Licença

Este projeto está licenciado sob a **Licença MIT** - veja o arquivo [LICENSE](LICENSE) para detalhes completos.

---

<div align="center">

**Desenvolvido pela Talmax Digital para a comunidade VB6**

*"Trazendo o consumo moderno de APIs para o clássico Visual Basic 6.0"*

---

**Versão**: 1.0.0 | **Status**: Estável | **Última atualização**: Julho 2025

</div>

</div>
