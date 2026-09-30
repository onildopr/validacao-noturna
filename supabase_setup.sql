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

-- 2) OPERATIONS: liberar leitura e cadastro sem login
alter table public.operations enable row level security;

drop policy if exists "operations_select_all" on public.operations;
drop policy if exists "operations_insert_all" on public.operations;
drop policy if exists "operations_update_all" on public.operations;

create policy "operations_select_all" on public.operations
  for select using (true);
create policy "operations_insert_all" on public.operations
  for insert with check (true);
create policy "operations_update_all" on public.operations
  for update using (true) with check (true);

-- 3) Operação inicial (ajuste o nome se quiser)
insert into public.operations (code, name, active)
values ('ERD1', 'Expedição ERD1', true)
on conflict (code) do nothing;

-- 4) Realtime: publicar mudanças da routes_state para os outros aparelhos
do $$
begin
  if not exists (
    select 1 from pg_publication_tables
    where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'routes_state'
  ) then
    alter publication supabase_realtime add table public.routes_state;
  end if;
end $$;

-- 5) Recarrega o cache do schema da API
notify pgrst, 'reload schema';
