<div id="top">

<div align="center">

# VB6 API Consumption

Sistema completo para consumo de APIs REST em Visual Basic 6.0 com suporte nativo a JSON.

[![Version](https://img.shields.io/badge/version-1.0.2-blue?style=flat-square)](CHANGELOG.md)
[![Status](https://img.shields.io/badge/status-Stable-brightgreen?style=flat-square)](#)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)
[![Documentation](https://img.shields.io/badge/docs-blueviolet?style=flat-square)](docs\README.md)

**Tecnologias Utilizadas:**

<a href="https://docs.microsoft.com/en-us/previous-versions/visual-studio/"><img alt="VB6" src="https://img.shields.io/badge/Visual%20Basic-6.0-blue?style=flat-square&logo=microsoft&logoColor=white"></a>
<a href="#"><img alt="JSON" src="https://img.shields.io/badge/JSON-Native-orange?style=flat-square&logo=json&logoColor=white"></a>
<a href="#"><img alt="HTTP" src="https://img.shields.io/badge/HTTP-REST-green?style=flat-square&logo=http&logoColor=white"></a>
<a href="#"><img alt="XML" src="https://img.shields.io/badge/XML-HTTP-red?style=flat-square&logo=xml&logoColor=white"></a>

> **Versão Estável** - Sistema testado e pronto para uso em produção!

</div>

<br>
<hr>

## 📋 Tabela de Conteúdos

- [🚀 Visão Geral](#-visão-geral)
- [✨ Funcionalidades](#-funcionalidades)
- [🛠️ Tecnologias](#️-tecnologias)
- [📋 Requisitos](#-requisitos)
- [📖 Documentação](#-documentação)
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

## 📖 Documentação

A documentação completa do projeto está disponível em [docs](docs\README.md).

## 🗺️ Roadmap

### ✅ Concluído (v1.0.0)

- [x] Cliente HTTP completo com todos os métodos
- [x] Parser e gerador JSON nativo
- [x] Sistema de headers configuráveis
- [x] Tratamento de erros robusto
- [x] Documentação completa

### 📅 Planejado

- [x] Documentação completa
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

Para instruções detalhadas sobre como contribuir, incluindo o fluxo de trabalho, padrões de código e processo de revisão, consulte nosso [guia de contribuição](docs/contributing.md).

## 📜 Licença

Este projeto está licenciado sob a **Licença MIT** - veja o arquivo [LICENSE](LICENSE) para detalhes completos.

---

<div align="center">

**Desenvolvido pela Talmax Digital para a comunidade VB6**

*"Trazendo o consumo moderno de APIs para o clássico Visual Basic 6.0"*

---

**Versão**: 1.0.2 | **Status**: Estável | **Última atualização**: Julho 2025

</div>

</div>
