-- ==========================================================
-- Configuração do Supabase para o app de Conferência (sem login)
-- Cole tudo no SQL Editor do Supabase e clique em "Run".
-- Pode rodar mais de uma vez sem problema.
-- ==========================================================

-- 1) Permissões básicas para a chave pública (anon)
grant select                 on public.operations   to anon, authenticated;  -- insert/update: ver 3b
grant select, insert, update on public.routes_state to anon, authenticated;
grant select, insert         on public.scan_events  to anon, authenticated;
grant usage, select on all sequences in schema public to anon, authenticated;

-- 2) Políticas de acesso (sem login)
alter table public.operations   enable row level security;
alter table public.routes_state enable row level security;
alter table public.scan_events  enable row level security;

drop policy if exists "operations_select_all" on public.operations;
drop policy if exists "operations_insert_all" on public.operations;
drop policy if exists "operations_update_all" on public.operations;
create policy "operations_select_all" on public.operations for select using (true);
create policy "operations_insert_all" on public.operations for insert with check (true);
create policy "operations_update_all" on public.operations for update using (true) with check (true);

drop policy if exists "routes_state_select_all" on public.routes_state;
drop policy if exists "routes_state_insert_all" on public.routes_state;
drop policy if exists "routes_state_update_all" on public.routes_state;
create policy "routes_state_select_all" on public.routes_state for select using (true);
create policy "routes_state_insert_all" on public.routes_state for insert with check (true);
create policy "routes_state_update_all" on public.routes_state for update using (true) with check (true);

drop policy if exists "scan_events_select_all" on public.scan_events;
drop policy if exists "scan_events_insert_all" on public.scan_events;
create policy "scan_events_select_all" on public.scan_events for select using (true);
create policy "scan_events_insert_all" on public.scan_events for insert with check (true);

-- 3) Operação inicial
insert into public.operations (code, name, active)
values ('ERD1', 'Expedição ERD1', true)
on conflict (code) do nothing;

-- 3b) PIN único para todas as operações (pedido em excluir rota, limpar o dia,
--     excluir bipagem de placa e salvar operação no Admin).
--     O hash fica em app_config, que a chave pública NÃO consegue ler nem alterar.
--     O app só pergunta ao banco "tem PIN?" e "este PIN está certo?".
--     Definir / trocar / remover o PIN: só por SQL (comandos no final deste arquivo).
create extension if not exists pgcrypto with schema extensions;

create table if not exists public.app_config (
  key   text primary key,
  value text not null
);
alter table public.app_config enable row level security;
revoke all on public.app_config from anon, authenticated;

create or replace function public.pin_enabled()
returns boolean
language sql stable security definer
set search_path = public
as $$
  select exists (select 1 from public.app_config where key = 'pin_hash');
$$;

create or replace function public.check_pin(p_pin text)
returns boolean
language sql stable security definer
set search_path = public, extensions
as $$
  select coalesce(
    (select value = encode(extensions.digest('conferencia:' || coalesce(trim(p_pin), ''), 'sha256'), 'hex')
       from public.app_config where key = 'pin_hash'),
    true);  -- sem PIN cadastrado = liberado
$$;

revoke all on function public.pin_enabled()     from public;
revoke all on function public.check_pin(text)   from public;
grant execute on function public.pin_enabled()   to anon, authenticated;
grant execute on function public.check_pin(text) to anon, authenticated;

-- Operações: a chave pública só cadastra/altera código, nome e ativa
alter table public.operations drop column if exists pin_hash;
revoke insert, update on public.operations from anon, authenticated;
grant insert (code, name, active), update (code, name, active) on public.operations to anon, authenticated;

-- 4) scan_events passa a ser a fonte das bipagens
--    client_id: ID gerado no aparelho; o índice único impede duplicar quando o app reenvia
--    uma bipagem (ex.: depois de ficar sem internet).
alter table public.scan_events add column if not exists client_id text;
alter table public.scan_events add column if not exists device_id text;
create unique index if not exists scan_events_client_id_key on public.scan_events (client_id);

-- Índices para as consultas do app (carga do dia e busca por ID)
create index if not exists scan_events_op_day_id_idx on public.scan_events (operation_code, day, id);
create index if not exists scan_events_package_idx   on public.scan_events (package_id, scanned_at desc);

-- 5) Realtime: bipagens e rotas chegam na hora nos outros aparelhos
do $$
begin
  if not exists (select 1 from pg_publication_tables
                 where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'routes_state') then
    alter publication supabase_realtime add table public.routes_state;
  end if;
  if not exists (select 1 from pg_publication_tables
                 where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'scan_events') then
    alter publication supabase_realtime add table public.scan_events;
  end if;
end $$;

-- 6) Acompanhamento geral calculado no banco (o app não precisa baixar os dados de cada operação)
create or replace function public.day_progress(p_day date)
returns table (
  operation_code text,
  name text,
  routes int,
  total_ids int,
  conferidos int,
  fora int,
  updated_at timestamptz
)
language sql
stable
as $$
  select
    o.code,
    o.name,
    coalesce((select count(*)::int
                from jsonb_each(coalesce(rs.data -> 'routes', '{}'::jsonb))), 0),
    coalesce((select sum(jsonb_array_length(coalesce(e.value -> 'ids', '[]'::jsonb)))::int
                from jsonb_each(coalesce(rs.data -> 'routes', '{}'::jsonb)) as e), 0),
    (select count(distinct se.package_id)::int from public.scan_events se
      where se.operation_code = o.code and se.day = p_day and se.result = 'ok'),
    (select count(distinct se.package_id)::int from public.scan_events se
      where se.operation_code = o.code and se.day = p_day and se.result = 'fora'),
    greatest(rs.updated_at,
             (select max(se.scanned_at) from public.scan_events se
               where se.operation_code = o.code and se.day = p_day))
  from public.operations o
  left join public.routes_state rs on rs.operation_code = o.code and rs.day = p_day
  where o.active
  order by o.code;
$$;

grant execute on function public.day_progress(date) to anon, authenticated;

-- 7) Limpeza automática: apaga dados com mais de 90 dias (todo dia às 05:00 de Porto Velho = 09:00 UTC)
--    Mantém o banco bem abaixo do limite de 500 MB do plano gratuito.
create extension if not exists pg_cron;

select cron.schedule(
  'limpeza-conferencia-90-dias',
  '0 9 * * *',
  $$
    delete from public.scan_events  where day < current_date - 90;
    delete from public.routes_state where day < current_date - 90;
  $$
);

-- 8) Recarrega o cache do schema da API
notify pgrst, 'reload schema';

-- ==========================================================
-- PIN: comandos para rodar À PARTE quando precisar (não rodam junto com o resto)
-- ==========================================================
-- Definir / trocar o PIN: gere o hash de  conferencia:SEU_PIN  em SHA-256 (hex) e use:
--   insert into public.app_config (key, value) values ('pin_hash', '<hash>')
--   on conflict (key) do update set value = excluded.value;
--
-- Remover o PIN (exclusões voltam a não pedir PIN):
--   delete from public.app_config where key = 'pin_hash';
