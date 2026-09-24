# Portail-tech V2

Suite d'outils radioprotection terrain, **offline-first**, partage entre techniciens.
Stack : **Vite + React + TypeScript + Supabase + PWA**, IndexedDB local (Dexie) comme
source de vrit de l'interface, Supabase comme base commune synchronise.

## Les 3 modules

- **ECME / talons** : base commune des talons (sources, dbitmtres, etc.).
  Lecture pour tous, criture rserve aux **rfrents**. Pour les sources, l'activit
  du jour est calcule automatiquement par dcroissance (A = A0e^(-ln2t/T)).
- **Imputations** : les heures de chaque technicien (donnes **perso**, isoles par RLS).
- **Journal** : main courante **partage** par l'quipe (chacun supprime ses propres entres).

---

## 1. Prrequis

- Node.js 18+ et npm
- Un projet Supabase (gratuit) : https://supabase.com

## 2. Crer le backend Supabase

1. Cre un projet sur Supabase.
2. Ouvre **SQL Editor**, colle le contenu de `supabase/schema.sql`, excute.
   a cre les tables, les index, les triggers et les **RLS**.
3. **Authentication > Providers** : laisse "Email" activ. Pour des comptes crs  la
   main sans email de confirmation, dsactive "Confirm email" dans les rglages Auth.

## 3. Crer les comptes techniciens

- **Authentication > Users > Add user** : un compte par technicien (email + mot de passe).
- Un profil est cr automatiquement (rle `tech` par dfaut).
- Pour donner les droits d'criture sur les talons  un rfrent :
  ```sql
  update public.profiles set role = 'referent' where id = '<uuid-du-user>';
  ```
  (l'uuid se trouve dans Authentication > Users)

## 4. Configurer le front

```bash
cp .env.example .env
```
Renseigne dans `.env` (valeurs dans **Project Settings > API**) :
```
VITE_SUPABASE_URL=https://xxxx.supabase.co
VITE_SUPABASE_ANON_KEY=eyJ...
```
La cl **anon** est publique par design : la scurit passe entirement par les RLS.
Ne jamais mettre la cl `service_role` dans le front.

## 5. Lancer en dev

```bash
npm install
npm run dev
```

## 6. Build de production

```bash
npm run build      # tsc --noEmit + vite build  ->  dist/
npm run preview    # pour tester le build localement
```

### Dploiement

- **Recommand (Cloudflare Pages / Netlify / Vercel)** : pointer sur le repo,
  build `npm run build`, dossier `dist/`, et dclarer les variables
  `VITE_SUPABASE_URL` / `VITE_SUPABASE_ANON_KEY`. La base reste `/`.
- **GitHub Pages** (dpt "projet", ex. `/Portail-tech/`) :
  ```bash
  VITE_BASE=/Portail-tech/ npm run build
  cp dist/index.html dist/404.html   # routage SPA
  ```
  puis publier `dist/`.

## 7. Installer en PWA sur mobile

Ouvrir l'URL dans le navigateur > menu > **Ajouter  l'cran d'accueil**.
L'appli se lance plein cran et fonctionne **hors-ligne** une fois les talons
tlchargs (utile en zone contrle, sans rseau).

---

## Comment a marche (offline-first)

- L'UI lit et crit **uniquement dans IndexedDB** (Dexie) : instantan, hors-ligne.
- `src/lib/sync.ts` rconcilie avec Supabase ds qu'il y a du rseau + une session :
  - **push** des lignes marques `_dirty` (cres/modifies hors-ligne) ;
  - **pull** incrmental par curseur `updated_at` (pagination 500).
- Suppression = **soft-delete** (`deleted = true`) pour ne jamais rater une suppression
  faite hors-ligne. Conflits : dernier qui crit gagne (LWW sur `updated_at`).
- La barre en haut indique l'tat rseau, la dernire synchro et le nombre de
  modifications en attente ; un tap force la synchro.

## Limites connues / pistes d'extension

- Le journal affiche "vous" sur tes entres ; pour afficher le **nom des autres
  auteurs**, synchroniser aussi la table `profiles` dans Dexie et faire la jointure.
- Pas de gestion des pices jointes / certificats PDF (seulement une rfrence + URL).
- Ajouter un site : `src/lib/types.ts` (`SITES`).
- Ajouter un radionuclide : `src/lib/decay.ts` (`HALF_LIVES_DAYS`).
- Ajouter un module : crer `src/modules/<x>/`, une table + RLS dans `schema.sql`,
  une table Dexie dans `src/lib/db.ts`, et l'intgrer  `sync.ts` + au routeur.

## Arborescence

```
portail-tech/
 supabase/schema.sql      # tables + RLS + triggers
 src/
   lib/                   # supabase, db (Dexie), sync, decay, types
   auth/                  # AuthProvider, RequireAuth, Login
   components/            # Layout, SyncStatus, Chips
   modules/
     ecme/                # EtalonsList, EtalonDetail (dcroissance)
     imputations/         # ImputationsView
     journal/             # JournalView
 .env.example
 vite.config.ts           # PWA + base configurable (VITE_BASE)
```
