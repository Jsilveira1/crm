# Propostas SKA — sistema web

Gerador de propostas comerciais + cadastro de clientes, contatos, licenças e produtos.
Página estática (HTML/JS), dados no Supabase com login obrigatório (RLS).

- `index.html`, `estilo.css`, `app.js` — o sistema
- `config.js` — URL do Supabase e chave *publishable* (pública por natureza; a proteção dos dados vem do login + RLS)
- `libs/` — supabase-js, pizzip, docxtemplater

Este repositório **não contém dados de clientes** nem o modelo Word da proposta (ficam no banco, acessíveis só com login).
