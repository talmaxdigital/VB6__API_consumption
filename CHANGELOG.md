# Changelog

Todas as alterações notáveis neste projeto serão documentadas neste arquivo.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/),
e este projeto adere ao [Versionamento Semântico](https://semver.org/lang/pt-BR/).

## [1.1.0] - 2025-09-10

### Adicionado

- **HttpRequestBuilder**: Adicionada uma nova classe `HttpRequestBuilder` que implementa o padrão Builder para uma construção de requisições HTTP mais fluente e legível.
- **Detecção Automática de JSON**: A classe `HttpRequest` agora detecta automaticamente se o corpo da requisição é um JSON e define o header `Content-Type` como `application/json; charset=utf-8` para simplificar o envio de dados.

### Corrigido

- **Sincronização de Requisições**: Corrigido um bug na classe `HttpRequest` onde o código podia continuar a execução antes da resposta de uma requisição síncrona ser recebida.
- **Codificação de Caracteres**: Corrigido o tratamento de caracteres especiais (escape) na construção de strings JSON no `JsonHelper`.

### Modificado

- **Renomeação de Classe**: A classe `cHttpRequest` foi renomeada para `HttpRequest` para seguir um padrão de nomenclatura mais limpo.
- **Limpeza de Código**: Removidas referências a projetos específicos (TomTicket) dos comentários no módulo `HttpClient`.
- **Documentação**: A documentação antiga em formato Markdown foi removida do repositório para ser substituída por uma versão mais atualizada futuramente.
- **Estrutura do Projeto**: Atualizados os arquivos de projeto (`.vbp` e `.vbw`) para refletir as renomeações de classes e outras mudanças.

## [1.0.3] - 2025-08-01

### Adicionado

- Módulo `RateLimiter` para controle de taxa de requisições

### Corrigido

- Suporte para especificar o encoding de arquivos via `settings.json` ("windows1252")

## [1.0.2] - 2025-07-30

### Adicionado

- Documentação técnica completa com 9 guias especializados

### Melhorado

- README: Reorganização com foco na documentação técnica

## [1.0.1] - 2025-07-28

### Modificado

- Suporte para enviar parâmetros no body das requisições GET

## [1.0.0] - 2025-07-28

### Adicionado

- HttpClient: Cliente HTTP completo com suporte a GET, POST, PUT, DELETE, PATCH
- HttpResponse: Classe para encapsular respostas HTTP com parsing automático de JSON
- HttpRequest: Wrapper para XMLHTTP com timeout configurável
- Métodos especializados para JSON (GetJson, PostJson, PutJson)
- Sistema de headers configuráveis (padrão e customizados)
- Utilitários para URL encoding e query strings
- Suporte para upload/download de arquivos
- Documentação completa com exemplos práticos

### Melhorado

- JsonHelper: Parser JSON otimizado com melhor tratamento de erros
- Suporte completo a tipos VB6 (Dictionary, Collection, primitivos)
- Sistema robusto de tratamento de erros HTTP

## [0.1.0] - 2025-07-22

### Adicionado

- Estrutura inicial do projeto
- Configuração básica do ambiente VB6
- JsonHelper: Parser e conversor JSON para VB6
