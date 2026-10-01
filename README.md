# Validação noturna — Conferência de rotas

App web (HTML + jQuery + Supabase) para conferir pacotes por rota: importa as rotas colando o HTML da página de origem, bipa os pacotes, aponta fora de rota e duplicados, vincula placa → rotas e exporta CSV/XLSX. Funciona em vários aparelhos ao mesmo tempo e continua funcionando sem internet (sincroniza depois).

## Estrutura

| Arquivo | O que faz |
|---|---|
| `index.html` | Tela principal |
| `search.html` | Página só de busca de IDs no banco |
| `js/core.js` | Constantes, estado do app e utilitários (datas, IDs, Supabase) |
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
- **`operations`**: operações (ERD1...).
- **`app_config`**: configurações protegidas (hash do PIN). A chave pública não tem acesso.

Dados com mais de 90 dias são apagados automaticamente (pg_cron).

## Configurar o Supabase

1. No painel do Supabase, abra o **SQL Editor**, cole o conteúdo de `supabase_setup.sql` e clique em **Run** (pode rodar de novo sempre que o arquivo mudar).
2. A URL e a chave pública do projeto ficam no final de `index.html` e em `search.html`.

## PIN

Um PIN único para todas as operações. Com PIN cadastrado, **excluir rota**, **limpar o dia**, **excluir bipagem de placa** e **salvar operação no Admin** pedem o PIN, que vale por 10 minutos no aparelho.

- O PIN é definido **só pelo banco** (SQL Editor). O hash fica na tabela `app_config`, que a chave pública não lê nem altera.
- O app só pergunta ao banco se existe PIN (`pin_enabled()`) e se o digitado está certo (`check_pin()`).
- Sem conexão, as ações que pedem PIN ficam bloqueadas.

Definir ou trocar o PIN (o hash é o SHA-256 em hexadecimal de `conferencia:SEU_PIN`):

```sql
insert into public.app_config (key, value) values ('pin_hash', '<hash>')
on conflict (key) do update set value = excluded.value;
```

Remover o PIN:

```sql
delete from public.app_config where key = 'pin_hash';
```

Para gerar o hash sem deixar o PIN no histórico do SQL Editor, rode no computador (com Node.js):

```bash
node -e "console.log(require('crypto').createHash('sha256').update('conferencia:' + process.argv[1]).digest('hex'))" SEU_PIN
```

## Testes

Precisa só do [Node.js](https://nodejs.org) 20 ou mais novo (sem instalar pacotes):

```bash
npm test
```

Os testes rodam o código real do app com um Supabase falso em memória, simulando vários aparelhos, falta de internet e recarregamento da página.

## Publicar uma versão nova

Ao mudar arquivos em `js/`, atualize o `?v=...` das tags `<script>` em `index.html` e `search.html` para os aparelhos não usarem a versão antiga do cache.
