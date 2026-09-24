-- =====================================================================
-- Portail-tech V2 - schma Supabase (Postgres)
--  excuter dans Supabase > SQL Editor.
-- Idempotent : on peut le relancer (drop policy / if not exists / on conflict).
-- =====================================================================

-- ---------- PROFILES ----------
create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  nom text,
  role text not null default 'tech' check (role in ('tech','referent','admin')),
  updated_at timestamptz not null default now()
);

-- ---------- MODELES D'ETALON (paramtrables, partags) ----------
create table if not exists public.etalon_modeles (
  id uuid primary key default gen_random_uuid(),
  nom text not null unique,
  ordre integer not null default 50,
  champs jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);

-- ---------- ETALONS (ECME) - base commune partage ----------
create table if not exists public.etalons (
  id uuid primary key default gen_random_uuid(),
  modele_id uuid references public.etalon_modeles(id) on delete set null,
  modele_nom text not null default '',
  designation text not null default '',
  num_serie text,
  num_client text,
  statut text not null default 'en_service',
  date_etalonnage date,
  date_echeance date,
  certificat_ref text,
  certificat_path text,
  certificat_nom text,
  valeurs jsonb not null default '{}'::jsonb,
  champs_libres jsonb not null default '[]'::jsonb,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);
-- si la table existait dj (ancienne version), on ajoute les colonnes manquantes :
alter table public.etalons add column if not exists certificat_path text;
alter table public.etalons add column if not exists certificat_nom text;
alter table public.etalons add column if not exists num_client text;

-- ---------- JOURNAL - log d'activit partag ----------
create table if not exists public.journal (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  ts timestamptz not null default now(),
  type text not null default 'info' check (type in ('ajout','modification','suppression','erreur','info')),
  message text not null default '',
  etalon_id uuid,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);

-- ---------- index (curseur de synchro) ----------
create index if not exists idx_modeles_updated on public.etalon_modeles(updated_at);
create index if not exists idx_etalons_updated on public.etalons(updated_at);
create index if not exists idx_journal_updated on public.journal(updated_at);
create index if not exists idx_journal_user on public.journal(user_id);

-- ---------- updated_at automatique ----------
create or replace function public.set_updated_at()
returns trigger language plpgsql as $$
begin new.updated_at = now(); return new; end; $$;

drop trigger if exists t_profiles_upd on public.profiles;
create trigger t_profiles_upd before update on public.profiles
  for each row execute function public.set_updated_at();

drop trigger if exists t_modeles_upd on public.etalon_modeles;
create trigger t_modeles_upd before update on public.etalon_modeles
  for each row execute function public.set_updated_at();

drop trigger if exists t_etalons_upd on public.etalons;
create trigger t_etalons_upd before update on public.etalons
  for each row execute function public.set_updated_at();

drop trigger if exists t_journal_upd on public.journal;
create trigger t_journal_upd before update on public.journal
  for each row execute function public.set_updated_at();

-- ---------- cration auto du profil  l'inscription ----------
create or replace function public.handle_new_user()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  insert into public.profiles (id, nom) values (new.id, new.email);
  return new;
end; $$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created after insert on auth.users
  for each row execute function public.handle_new_user();

-- =====================================================================
-- RLS  (cl anon publique : toute la scurit est ici)
-- =====================================================================
alter table public.profiles      enable row level security;
alter table public.etalon_modeles enable row level security;
alter table public.etalons       enable row level security;
alter table public.journal       enable row level security;

-- profiles : lecture quipe, criture de son propre profil
drop policy if exists profiles_select on public.profiles;
drop policy if exists profiles_insert on public.profiles;
drop policy if exists profiles_update on public.profiles;
create policy profiles_select on public.profiles for select to authenticated using (true);
create policy profiles_insert on public.profiles for insert to authenticated with check (id = auth.uid());
create policy profiles_update on public.profiles for update to authenticated using (id = auth.uid()) with check (id = auth.uid());

-- modeles : lecture + criture pour tout utilisateur connect
drop policy if exists modeles_select on public.etalon_modeles;
drop policy if exists modeles_write on public.etalon_modeles;
create policy modeles_select on public.etalon_modeles for select to authenticated using (true);
create policy modeles_write  on public.etalon_modeles for all to authenticated using (true) with check (true);

-- etalons : lecture + criture pour tout utilisateur connect
drop policy if exists etalons_select on public.etalons;
drop policy if exists etalons_write on public.etalons;
drop policy if exists etalons_insert on public.etalons;
drop policy if exists etalons_update on public.etalons;
create policy etalons_select on public.etalons for select to authenticated using (true);
create policy etalons_write  on public.etalons for all to authenticated using (true) with check (true);

-- journal : lecture quipe, criture/maj de ses propres entres
drop policy if exists journal_select on public.journal;
drop policy if exists journal_insert on public.journal;
drop policy if exists journal_update on public.journal;
create policy journal_select on public.journal for select to authenticated using (true);
create policy journal_insert on public.journal for insert to authenticated with check (user_id = auth.uid());
create policy journal_update on public.journal for update to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid());

-- =====================================================================
-- MODELES PAR DEFAUT (modifiables ensuite dans l'onglet Paramtrage)
-- =====================================================================
-- Source : champs calqus sur le fichier EDF (DTS-000-DI012). On FORCE la mise
-- à jour même si le modèle existe déjà, pour appliquer ces champs partout.
insert into public.etalon_modeles (nom, ordre, champs) values
('Source', 1, '[
  {"cle":"c_ray","libelle":"Rayonnement","type":"liste","options":["Alpha","Bêta","Gamma","X","Neutron"],"requis":true},
  {"cle":"c_rn","libelle":"Radionucléide","type":"radionucleide","requis":true},
  {"cle":"c_act","libelle":"Activité CE (kBq)","type":"activite_ref","requis":true},
  {"cle":"c_dce","libelle":"Date CE","type":"date_ref","requis":true},
  {"cle":"c_prop","libelle":"Propriétaire","type":"texte"},
  {"cle":"c_mar","libelle":"Marque","type":"texte"},
  {"cle":"c_ref","libelle":"Référence / Drawing","type":"texte"},
  {"cle":"c_typs","libelle":"Type S (EDF)","type":"texte"},
  {"cle":"c_inc","libelle":"Incertitude activité k=2 (%)","type":"nombre"},
  {"cle":"c_f4","libelle":"Flux 4pi sr (pps)","type":"flux"},
  {"cle":"c_f2","libelle":"Flux 2pi sr (pps)","type":"flux"},
  {"cle":"c_diam","libelle":"Diam. actif (mm)","type":"nombre"},
  {"cle":"c_long","libelle":"Long. active (mm)","type":"nombre"},
  {"cle":"c_larg","libelle":"Larg. active (mm)","type":"nombre"},
  {"cle":"c_use","libelle":"Utilisée ?","type":"booleen"}
]'::jsonb)
on conflict (nom) do update set ordre = excluded.ordre, champs = excluded.champs;

-- Les autres modèles : seulement s'ils n'existent pas (ne pas écraser tes modifs)
insert into public.etalon_modeles (nom, ordre, champs) values
('Debitmetre', 2, '[
  {"cle":"c_ma","libelle":"Marque","type":"texte"},
  {"cle":"c_mo","libelle":"Modele","type":"texte"},
  {"cle":"c_ga","libelle":"Gamme","type":"texte"},
  {"cle":"c_un","libelle":"Unite","type":"texte"},
  {"cle":"c_co","libelle":"Coefficient d etalonnage","type":"nombre"}
]'::jsonb),
('Multimedia', 3, '[
  {"cle":"c_ma","libelle":"Marque","type":"texte"},
  {"cle":"c_mo","libelle":"Modele","type":"texte"},
  {"cle":"c_me","libelle":"Type de mesure","type":"texte"},
  {"cle":"c_ga","libelle":"Gamme","type":"texte"}
]'::jsonb),
('Manuel', 9, '[]'::jsonb)
on conflict (nom) do nothing;

-- =====================================================================
-- NOTE : si tu avais dj une ANCIENNE version des tables (avec
-- imputations / colonnes 'type' / 'caracteristiques'), supprime-les
-- d'abord pour repartir propre :
--   drop table if exists public.imputations cascade;
--   drop table if exists public.etalons cascade;
--   puis relance tout ce script.
-- =====================================================================

-- =====================================================================
-- STORAGE : bucket des certificats (priv) + policies
-- (Tu peux aussi crer le bucket  la main : Storage > New bucket >
--  nom "certificats", priv. Mais ce SQL le fait pour toi.)
-- =====================================================================
insert into storage.buckets (id, name, public)
values ('certificats', 'certificats', false)
on conflict (id) do nothing;

drop policy if exists cert_read   on storage.objects;
drop policy if exists cert_insert on storage.objects;
drop policy if exists cert_update on storage.objects;
drop policy if exists cert_delete on storage.objects;
create policy cert_read   on storage.objects for select to authenticated using (bucket_id = 'certificats');
create policy cert_insert on storage.objects for insert to authenticated with check (bucket_id = 'certificats');
create policy cert_update on storage.objects for update to authenticated using (bucket_id = 'certificats') with check (bucket_id = 'certificats');
create policy cert_delete on storage.objects for delete to authenticated using (bucket_id = 'certificats');

-- ---------------------------------------------------------------------------
-- Champ "Client" (CNPE + DP2D) sur TOUS les modeles d'etalon.
-- Ajoute a chaque modele qui ne l'a pas deja -> idempotent.
-- ---------------------------------------------------------------------------
update public.etalon_modeles
set champs = champs || '[
  {"cle":"c_client","libelle":"Client","type":"liste","options":[
    "CNPE Belleville","CNPE Blayais","CNPE Bugey","CNPE Cattenom","CNPE Chinon",
    "CNPE Chooz","CNPE Civaux","CNPE Cruas","CNPE Dampierre","CNPE Flamanville",
    "CNPE Golfech","CNPE Gravelines","CNPE Nogent","CNPE Paluel","CNPE Penly",
    "CNPE Saint-Alban","CNPE Saint-Laurent","CNPE Tricastin",
    "DP2D Brennilis","DP2D Bugey 1","DP2D Chinon A","DP2D Chooz A",
    "DP2D Creys-Malville","DP2D Fessenheim","DP2D Phénix","DP2D Saint-Laurent A"
  ]}
]'::jsonb
where not (champs @> '[{"cle":"c_client"}]'::jsonb);

-- =====================================================================
-- MODULE PLANNING (saisie quotidienne : heures, dose, frais, trajet)
-- Donnees PERSO, isolees par utilisateur (RLS). Une ligne = un jour.
-- Idempotent : relancable sans risque.
-- =====================================================================

-- ---------- PLANNING_JOURS ----------
create table if not exists public.planning_jours (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  date date not null,
  type text not null default 'Travaillé',
  debut text,
  fin text,
  pause numeric(6,2),
  total numeric(6,2),
  h_norm numeric(6,2),
  h_supp numeric(6,2),
  site text,
  contrat text,
  imputation text,
  dose numeric(8,3),
  trajet boolean not null default false,
  t_ad text, t_af text, t_rd text, t_rf text,
  frais jsonb not null default '[]'::jsonb,
  commentaire text,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false,
  unique (user_id, date)   -- un seul jour par date et par utilisateur
);

-- ---------- PLANNING_PARAMS (1 jeu par utilisateur : id = user_id) ----------
create table if not exists public.planning_params (
  id uuid primary key references auth.users(id) on delete cascade,
  horaire numeric(4,2) not null default 7.5,
  matricule text,
  dosi text,
  nom text,
  prenom text,
  sup text,
  codes jsonb not null default '{}'::jsonb,
  trajet_defaut jsonb,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);
alter table public.planning_params add column if not exists sites_favoris jsonb not null default '[]'::jsonb;


-- ---------- index (curseur de synchro updated_at) ----------
create index if not exists idx_planning_updated      on public.planning_jours(updated_at);
create index if not exists idx_planning_user_date     on public.planning_jours(user_id, date);
create index if not exists idx_planningparams_updated on public.planning_params(updated_at);

-- ---------- updated_at automatique (meme fonction set_updated_at) ----------
drop trigger if exists t_planning_upd on public.planning_jours;
create trigger t_planning_upd before update on public.planning_jours
  for each row execute function public.set_updated_at();

drop trigger if exists t_planningparams_upd on public.planning_params;
create trigger t_planningparams_upd before update on public.planning_params
  for each row execute function public.set_updated_at();

-- ---------- RLS : chacun ne voit/ecrit que ses propres lignes ----------
alter table public.planning_jours  enable row level security;
alter table public.planning_params enable row level security;

drop policy if exists planning_jours_all on public.planning_jours;
create policy planning_jours_all on public.planning_jours for all to authenticated
  using (user_id = auth.uid()) with check (user_id = auth.uid());

drop policy if exists planning_params_all on public.planning_params;
create policy planning_params_all on public.planning_params for all to authenticated
  using (id = auth.uid()) with check (id = auth.uid());

-- =====================================================================
-- STORAGE : bucket des photos de notes de frais (prive) + policies
-- (calque sur le bucket "certificats")
-- =====================================================================
insert into storage.buckets (id, name, public)
values ('frais', 'frais', false)
on conflict (id) do nothing;

drop policy if exists frais_read   on storage.objects;
drop policy if exists frais_insert on storage.objects;
drop policy if exists frais_update on storage.objects;
drop policy if exists frais_delete on storage.objects;
create policy frais_read   on storage.objects for select to authenticated using (bucket_id = 'frais');
create policy frais_insert on storage.objects for insert to authenticated with check (bucket_id = 'frais');
create policy frais_update on storage.objects for update to authenticated using (bucket_id = 'frais') with check (bucket_id = 'frais');
create policy frais_delete on storage.objects for delete to authenticated using (bucket_id = 'frais');

-- =====================================================================
-- IMPUTATIONS : référence partagée (codes d'affaire x tâches Bertin)
-- Importée depuis l'Excel "Pointages". Lisible/écrivable par tout
-- utilisateur authentifié (pas de cloisonnement par utilisateur).
-- =====================================================================
create table if not exists public.imputations (
  id uuid primary key default gen_random_uuid(),
  client text,
  nom_projet text,
  num_projet text,            -- n° d'affaire (ex. "25760")
  tache text not null,        -- code tâche (ex. "11.01")
  nom_tache text,             -- libellé (ex. "RPM Bugey")
  commentaires text,
  site boolean not null default false,
  usine boolean not null default false,
  annee int,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false,
  unique (num_projet, tache)  -- clé naturelle pour l'upsert à l'import
);

create index if not exists idx_imputations_updated on public.imputations(updated_at);

drop trigger if exists t_imputations_upd on public.imputations;
create trigger t_imputations_upd before update on public.imputations
  for each row execute function public.set_updated_at();

alter table public.imputations enable row level security;
drop policy if exists imputations_all on public.imputations;
create policy imputations_all on public.imputations for all to authenticated
  using (true) with check (true);

-- Couleur (liseré) par pointage, fixée à l'import (cf. lib/imputations.ts).
alter table public.imputations add column if not exists couleur text;

-- ===== Base matériels (référence partagée, importée depuis l'Excel de suivi) =====
create table if not exists public.materiels (
  scan text primary key,                 -- code GMO² ou repère fonctionnel
  id_type text not null default 'gmo2',  -- 'gmo2' | 'repere'
  type_code text,                        -- 1re partie du GMO² (null si repère)
  id_court text,                         -- BUG070 / repère — n° sur étiquette & CV
  designation text,
  code_model text,
  sn text,
  domaine text,                          -- RPM/KZC/KRS…
  site text,                             -- déduit du préfixe de l'id court
  localisation text,
  etat text not null default 'en_place', -- en_place|magasin|devis|reparation|reforme
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);
create table if not exists public.corim_types (
  type_code text primary key,            -- SCAN de la feuille « correspondance CORIM »
  designation text,                      -- libellé Corim
  type_appareil text,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);

drop trigger if exists trg_materiels_updated on public.materiels;
create trigger trg_materiels_updated before update on public.materiels
  for each row execute function public.set_updated_at();
drop trigger if exists trg_corim_types_updated on public.corim_types;
create trigger trg_corim_types_updated before update on public.corim_types
  for each row execute function public.set_updated_at();

alter table public.materiels enable row level security;
alter table public.corim_types enable row level security;
drop policy if exists materiels_all on public.materiels;
create policy materiels_all on public.materiels for all to authenticated using (true) with check (true);
drop policy if exists corim_types_all on public.corim_types;
create policy corim_types_all on public.corim_types for all to authenticated using (true) with check (true);

-- Épinglage des étalons (favoris en haut de liste ECME)
alter table public.etalons add column if not exists epingle boolean not null default false;

-- ===== Favoris ECME (épinglage par utilisateur) =====
create table if not exists public.ecme_favoris (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  etalon_id uuid not null,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false,
  primary key (user_id, etalon_id)
);
drop trigger if exists trg_ecme_favoris_updated on public.ecme_favoris;
create trigger trg_ecme_favoris_updated before update on public.ecme_favoris
  for each row execute function public.set_updated_at();
alter table public.ecme_favoris enable row level security;
drop policy if exists ecme_favoris_all on public.ecme_favoris;
create policy ecme_favoris_all on public.ecme_favoris for all to authenticated
  using (user_id = auth.uid()) with check (user_id = auth.uid());

-- ===== Module Calcul (formules partagées ou perso) =====
create table if not exists public.calculs (
  id uuid primary key default gen_random_uuid(),
  nom text not null default '',
  scope text not null default 'perso',                 -- 'partage' | 'perso'
  user_id uuid default auth.uid() references auth.users(id) on delete cascade,
  type_source text,
  source_ids jsonb not null default '[]',
  source_filtres jsonb not null default '{}',
  composantes jsonb not null default '[]',
  formules jsonb not null default '[]',
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);
drop trigger if exists trg_calculs_updated on public.calculs;
create trigger trg_calculs_updated before update on public.calculs
  for each row execute function public.set_updated_at();
alter table public.calculs enable row level security;
drop policy if exists calculs_rw on public.calculs;
create policy calculs_rw on public.calculs for all to authenticated
  using (scope = 'partage' or user_id = auth.uid())
  with check (scope = 'partage' or user_id = auth.uid());
alter table public.calculs add column if not exists source_filtres jsonb not null default '{}';

-- =====================================================================
-- Module Intervention : listings de contrôles (partagés/perso) + lignes
-- =====================================================================
create table if not exists public.intervention_listings (
  id uuid primary key default gen_random_uuid(),
  nom text not null default '',
  scope text not null default 'perso',                 -- 'partage' | 'perso'
  user_id uuid default auth.uid() references auth.users(id) on delete cascade,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);

create table if not exists public.intervention_lignes (
  id uuid primary key default gen_random_uuid(),
  listing_id uuid not null references public.intervention_listings(id) on delete cascade,
  scan text not null default '',          -- GMO² complet (clé materiels) ou repère
  id_court text,                          -- ex. BUG070 (imprimé sur l'étiquette)
  designation text,                       -- snapshot au scan
  sn text,                                -- snapshot au scan
  type_controle text not null default 'Préventif',   -- Préventif | Correctif
  operation text not null default 'VP cas 1',        -- VP cas 1 | VP cas 2 | MP cas 1
  conformite text not null default 'Conforme',       -- Conforme | Conforme après intervention | Non conforme
  date_op date,
  validite_ans int,                       -- déduit de l'opération (VP*=1, MP*=3), modifiable
  echeance date,                          -- date_op + validite (si conforme), modifiable
  echeance_manuelle boolean not null default false,
  commentaire text,
  tri_exec text,                          -- trigramme exécutant
  tri_ct text,                            -- trigramme contrôleur technique
  ordre int not null default 0,
  user_id uuid default auth.uid() references auth.users(id) on delete cascade,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);
create index if not exists idx_interv_lignes_listing on public.intervention_lignes(listing_id);

drop trigger if exists trg_interv_listings_updated on public.intervention_listings;
create trigger trg_interv_listings_updated before update on public.intervention_listings
  for each row execute function public.set_updated_at();
drop trigger if exists trg_interv_lignes_updated on public.intervention_lignes;
create trigger trg_interv_lignes_updated before update on public.intervention_lignes
  for each row execute function public.set_updated_at();

alter table public.intervention_listings enable row level security;
alter table public.intervention_lignes enable row level security;

-- Listings : écriture ouverte sur les partagés (même mécanique que les calculs).
drop policy if exists interv_listings_rw on public.intervention_listings;
create policy interv_listings_rw on public.intervention_listings for all to authenticated
  using (scope = 'partage' or user_id = auth.uid())
  with check (scope = 'partage' or user_id = auth.uid());

-- Lignes : suivent la visibilité/écriture de leur listing parent (source unique du scope).
drop policy if exists interv_lignes_rw on public.intervention_lignes;
create policy interv_lignes_rw on public.intervention_lignes for all to authenticated
  using (exists (
    select 1 from public.intervention_listings l
    where l.id = listing_id and (l.scope = 'partage' or l.user_id = auth.uid())
  ))
  with check (exists (
    select 1 from public.intervention_listings l
    where l.id = listing_id and (l.scope = 'partage' or l.user_id = auth.uid())
  ));

-- Trigramme de l'exécutant (préremplissage des lignes d'intervention)
alter table public.planning_params add column if not exists trigramme text;

-- ===========================================================================
-- Bibliothèque RTR : IZ (codes d'accès des régimes de travail radiologique)
-- Perso : chaque technicien ne voit / modifie que ses IZ (synchro multi-appareils).
-- ===========================================================================
create table if not exists public.rtr_iz (
  id uuid primary key default gen_random_uuid(),
  code text not null,                     -- code IZ (scanné ou saisi)
  label text,                             -- libellé court
  site text,
  tranche text,
  local text,                             -- local / repère SIPN
  tags text[] not null default '{}',
  notes text,
  barcode_format text,                    -- CODE128 | CODE39 | EAN13 | ITF
  use_count int not null default 0,
  last_used timestamptz,
  created_at timestamptz not null default now(),
  user_id uuid default auth.uid() references auth.users(id) on delete cascade,
  updated_at timestamptz not null default now(),
  deleted boolean not null default false
);
create index if not exists idx_rtr_iz_user on public.rtr_iz(user_id);
create index if not exists idx_rtr_iz_updated on public.rtr_iz(updated_at);

drop trigger if exists trg_rtr_iz_updated on public.rtr_iz;
create trigger trg_rtr_iz_updated before update on public.rtr_iz
  for each row execute function public.set_updated_at();

alter table public.rtr_iz enable row level security;
drop policy if exists rtr_iz_own on public.rtr_iz;
create policy rtr_iz_own on public.rtr_iz for all to authenticated
  using (user_id = auth.uid())
  with check (user_id = auth.uid());
