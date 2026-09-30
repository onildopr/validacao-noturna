-- ==========================================================
-- Configuração do Supabase para o app de Conferência (sem login)
-- Cole tudo no SQL Editor do Supabase e clique em "Run".
-- Pode rodar mais de uma vez sem problema.
-- ==========================================================

-- 1) Permissões básicas para a chave pública (anon)
grant select, insert, update on public.operations   to anon, authenticated;
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

-- 3b) PIN por operação (guarda só o hash; protege exclusões contra acidentes)
alter table public.operations add column if not exists pin_hash text;

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
