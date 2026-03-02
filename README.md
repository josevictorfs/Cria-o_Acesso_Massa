# Gerador de Acesso (Angular Web)

Aplicação web moderna em Angular para geração de usuários em massa a partir de:

- PDF com lista de alunos (nome e RM)
- Entrada manual (um nome por linha)

## Funcionalidades

- Seleção de unidade
- Configuração de domínio de e-mail
- Extração de nomes/RM de PDF
- Geração de CSV (`Nome;Usuario;Email;Turma;Unidade`)
- Geração de DOCX com credenciais e imagem opcional
- Pré-visualização dos dados antes de exportar

## Requisitos

- Node.js 18+
- npm 9+

## Como executar

```bash
npm install
npm start
```

A aplicação será iniciada em `http://localhost:4200`.

## Build de produção

```bash
npm run build
```
