<div id="top">

<div align="center">

# 🎯 VB6 API Consumption

**Uma biblioteca moderna para consumo de APIs REST em Visual Basic 6.0, com suporte nativo a JSON e sem dependências externas complexas.**

[![Versão](https://img.shields.io/badge/version-1.0.3-blue?style=flat-square)](CHANGELOG.md)
[![Status](https://img.shields.io/badge/status-Estável-green?style=flat-square)](#)
[![Licença](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)
[![Documentação](https://img.shields.io/badge/docs-disponível-blueviolet?style=flat-square)](docs/README.md)

</div>

---

## Tabela de Conteúdos

- [🚀 Visão Geral](#-visão-geral)
- [✨ Funcionalidades](#-funcionalidades)
- [📦 Instalação e Configuração](#-instalação-e-configuração)
- [⚡ Guia Rápido](#-guia-rápido)
- [📚 Documentação Completa](#-documentação-completa)
- [🛠️ Tecnologias Utilizadas](#️-tecnologias-utilizadas)
- [🤝 Como Contribuir](#-como-contribuir)
- [📜 Licença](#-licença)

---

## 🚀 Visão Geral

A biblioteca **VB6 API Consumption** fornece um cliente HTTP completo e robusto para aplicações desenvolvidas em Visual Basic 6.0. Seu principal diferencial é a capacidade de manipular JSON de forma nativa, utilizando `Dictionary` e `Collection`, sem a necessidade de instalar DLLs ou OCXs de terceiros.

**Objetivos principais:**

- **Modernizar** o consumo de APIs em projetos VB6 legados.
- **Simplificar** a integração com serviços RESTful modernos.
- **Eliminar dependências** externas complexas para manipulação de JSON.

---

## ✨ Funcionalidades

- **Cliente HTTP Completo**: Suporte para métodos `GET`, `POST`, `PUT`, `DELETE` e `PATCH`.
- **Manipulação Nativa de JSON**:
  - **Parser**: Converte strings JSON em `Dictionary` (para objetos) e `Collection` (para arrays).
  - **Builder**: Gera strings JSON a partir de `Dictionary` e `Collection`.
- **Classe `HttpResponse`**: Encapsula a resposta HTTP, com acesso fácil a:
  - `StatusCode` e `StatusText`.
  - `Headers` da resposta.
  - `Text` (corpo da resposta como string).
  - `Json` (corpo da resposta já convertido para `Dictionary` ou `Collection`).
- **Classe `HttpRequest`**: Wrapper sobre `MSXML2.XMLHTTP` com configuração de timeout.
- **Gerenciamento de Headers**: Suporte para headers padrão (enviados em todas as requisições) e customizados.
- **Funções Auxiliares**: Utilitários para `UrlEncode`, `BuildQueryString` e construção de `multipart/form-data`.
- **Controle de Taxa**: Módulo `RateLimiter` para limitar o número de requisições por segundo.
- **Upload e Download**: Funções básicas para envio e recebimento de arquivos.

---

## 📦 Instalação e Configuração

Para utilizar a biblioteca em seu projeto VB6, siga os passos abaixo:

1. **Adicione os Arquivos**:
    - No menu do VB6, vá em `Project` > `Add Module` e adicione os seguintes arquivos:
      - `src/Modules/HttpClient.bas`
      - `src/Modules/JsonHelper.bas`
      - `src/Modules/RateLimiter.bas`
    - Em `Project` > `Add Class Module`, adicione:
      - `src/Classes/HttpRequest.cls`
      - `src/Classes/HttpResponse.cls`

2. **Adicione as Referências**:
    - Vá em `Project` > `References...` e marque a seguinte referência:
      - `Microsoft Scripting Runtime` (para `Scripting.Dictionary`).

3. **Dependências do Sistema**:
    - A biblioteca utiliza o `MSXML2.XMLHTTP`, que já vem instalado na maioria das versões do Windows. Nenhuma instalação adicional é necessária.

---

## 📚 Documentação Completa

Para exemplos detalhados sobre cada funcionalidade, consulte a **[Documentação Técnica](docs/README.md)**.

Lá você encontrará guias sobre:

- Requisições `POST`, `PUT` e `DELETE`.
- Manipulação avançada de JSON.
- Autenticação (Bearer Token, Basic Auth).
- Upload de arquivos.
- E muito mais.

---

## 🛠️ Tecnologias Utilizadas

- **Visual Basic 6.0**: Linguagem principal.
- **Microsoft Scripting Runtime**: Para uso do objeto `Scripting.Dictionary`.
- **Microsoft XML (MSXML2.XMLHTTP)**: Para realizar as requisições HTTP.

---

## 🤝 Como Contribuir

Contribuições são muito bem-vindas! Para colaborar:

- **Abra Issues**: Descreva problemas, bugs ou sugestões de melhorias.
- **Submeta Pull Requests**: Envie suas alterações com descrições claras. Lembre-se de seguir as convenções do projeto e atualizar o `CHANGELOG.md` se necessário.

---

## 📜 Licença

Este projeto é distribuído sob a Licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.

---

<div align="center">

Desenvolvido pela equipe Talmax

</div>

</div>
