-- Mode démo uniquement (base locale dans le navigateur).
-- Imite le minimum de Supabase : le schéma « auth », la table des utilisateurs et auth.uid().
create schema if not exists auth;
create table if not exists auth.users (
  id uuid primary key default gen_random_uuid(),
  email text,
  raw_user_meta_data jsonb not null default '{}'
);
create or replace function auth.uid() returns uuid
language sql stable as $$ select nullif(current_setting('app.uid', true), '')::uuid $$;
