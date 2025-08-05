# Changelog

Todas as alterações notáveis neste projeto serão documentadas neste arquivo.

O formato é baseado em [Keep a Changelog](https://keepachangelog.com/pt-BR/1.0.0/),
e este projeto adere ao [Versionamento Semântico](https://semver.org/lang/pt-BR/).

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
- cHttpRequest: Wrapper para XMLHTTP com timeout configurável
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
