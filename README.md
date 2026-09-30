# Validação noturna — Conferência de rotas

App web (HTML + jQuery + Supabase) para conferir pacotes por rota: importa as rotas colando o HTML da página de origem, bipa os pacotes, aponta fora de rota e duplicados, vincula placa → rotas e exporta CSV/XLSX. Funciona em vários aparelhos ao mesmo tempo e continua funcionando sem internet (sincroniza depois).

## Estrutura

| Arquivo | O que faz |
|---|---|
| `index.html` | Tela principal |
| `search.html` | Página só de busca de IDs no banco |
| `js/core.js` | Constantes, estado do app e utilitários (datas, IDs, Supabase, PIN) |
| `js/regras.js` | Regras da conferência: ok / fora de rota / duplicado, recálculo a partir das bipagens |
| `js/sync.js` | Sincronização com o Supabase, realtime e cache local |
| `js/banco.js` | Admin de operações, acompanhamento geral e busca de IDs |
| `js/importacao.js` | Importação de rotas pelo HTML e leitura de QR de placa/rota |
| `js/exportacao.js` | Exportações CSV / XLSX |
| `js/ui.js` | Renderização da interface |
| `js/eventos.js` | Cliques, leitor de código de barras e inicialização |
| `supabase_setup.sql` | Tabelas, permissões, realtime e limpeza automática do banco |
| `tests/` | Testes automatizados |

A ordem dos `<script>` no HTML importa: `core.js` primeiro, `eventos.js` por último.

## Como os dados ficam no banco

- **`scan_events`**: uma linha pequena por bipagem (~220 bytes). É a fonte da verdade das bipagens. O estado (conferidos, faltantes, fora de rota, duplicados) é recalculado no aparelho a partir delas, em ordem de horário.
- **`routes_state`**: uma linha por operação e dia com as definições das rotas (IDs importados, cluster, placas, exclusões).
- **`operations`**: operações (ERD1...) e o hash do PIN de cada uma.

Dados com mais de 90 dias são apagados automaticamente (pg_cron).

## Configurar o Supabase

1. No painel do Supabase, abra o **SQL Editor**, cole o conteúdo de `supabase_setup.sql` e clique em **Run** (pode rodar de novo sempre que o arquivo mudar).
2. A URL e a chave pública do projeto ficam no final de `index.html` e em `search.html`.

## PIN da operação

No **Admin**, dá para definir um PIN por operação. Com PIN, **excluir rota**, **limpar o dia** e **excluir bipagem de placa** pedem o PIN, que vale por 10 minutos no aparelho.

> O PIN evita acidentes e curiosos. Como o app não tem login, ele não protege contra alguém que saiba mexer no código ou no banco.

## Testes

Precisa só do [Node.js](https://nodejs.org) 20 ou mais novo (sem instalar pacotes):

```bash
npm test
```

Os testes rodam o código real do app com um Supabase falso em memória, simulando vários aparelhos, falta de internet e recarregamento da página.

## Publicar uma versão nova

Ao mudar arquivos em `js/`, atualize o `?v=...` das tags `<script>` em `index.html` e `search.html` para os aparelhos não usarem a versão antiga do cache.
