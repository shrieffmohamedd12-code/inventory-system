-- ============================================================
-- نظام المخازن — مخطط قاعدة البيانات على Supabase
-- شغّل هذا الملف كامل مرة واحدة من: Supabase Dashboard -> SQL Editor -> New query
-- ============================================================

create extension if not exists "pgcrypto";

-- ---------- profiles (مستخدمو النظام) ----------
-- id نفس id بتاع auth.users. email هنا هو إيميل داخلي وهمي (وليس إيميل حقيقي للمستخدم)
-- بيُستخدم فقط عشان نربط اسم المستخدم بحساب الدخول بتاعه في Supabase Auth.
create table if not exists profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  name text not null,
  email text not null unique,
  role text not null default 'staff' check (role in ('admin', 'staff')),
  created_at timestamptz not null default now()
);

-- ---------- factories (المصانع) ----------
create table if not exists factories (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  created_at timestamptz not null default now()
);

-- ---------- items (الأصناف — صف واحد لكل كود داخل مخزن معيّن) ----------
-- الألوان والمقاسات مخزّنة كـ jsonb بنفس شكل بيانات الواجهة تمامًا:
-- colors: [{ id, color, packQty, sizes: [{ id, size, qty }] }]
create table if not exists items (
  id uuid primary key default gen_random_uuid(),
  warehouse_code text not null,
  code text not null,
  name text,
  min_qty integer,
  colors jsonb not null default '[]'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (warehouse_code, code)
);

-- ---------- transactions (سجل حركات الصرف والاستلام) ----------
create table if not exists transactions (
  id uuid primary key default gen_random_uuid(),
  item_id uuid references items(id) on delete set null,
  warehouse_code text,
  code text,
  color text,
  size text,
  kind text,
  type text not null check (type in ('in', 'out')),
  qty integer not null,
  note text,
  by_name text,
  created_at timestamptz not null default now()
);

-- ---------- orders (أوردرات المصانع) ----------
create table if not exists orders (
  id uuid primary key default gen_random_uuid(),
  factory text not null,
  order_date date not null,
  warehouse_code text not null,
  code text not null,
  po text,
  internal_po text,
  color text not null,
  sizes jsonb not null default '[]'::jsonb,
  ordered_qty integer not null default 0,
  received_qty integer not null default 0,
  remaining_at_factory integer not null default 0,
  status text not null default 'open' check (status in ('open', 'closed')),
  receipts jsonb not null default '[]'::jsonb,
  created_by text,
  created_at timestamptz not null default now()
);

-- ---------- فهارس مفيدة ----------
create index if not exists idx_items_warehouse on items (warehouse_code);
create index if not exists idx_transactions_item on transactions (item_id);
create index if not exists idx_transactions_created on transactions (created_at desc);
create index if not exists idx_orders_factory on orders (factory);
create index if not exists idx_orders_status on orders (status);

-- ============================================================
-- دالة مساعدة: هل المستخدم الحالي أدمن؟
-- ============================================================
create or replace function is_admin()
returns boolean
language sql
security definer
set search_path = public
as $$
  select exists (
    select 1 from profiles where id = auth.uid() and role = 'admin'
  );
$$;

-- ============================================================
-- تفعيل RLS على كل الجداول
-- ============================================================
alter table profiles enable row level security;
alter table factories enable row level security;
alter table items enable row level security;
alter table transactions enable row level security;
alter table orders enable row level security;

-- ---------- profiles ----------
-- أي حد (حتى قبل تسجيل الدخول) يقدر يشوف قايمة الأسماء والصلاحيات عشان تظهر شاشة اختيار المستخدم
create policy "profiles are publicly viewable" on profiles
  for select using (true);

-- أول مستخدم بيسجّل نفسه (المسؤول الأول) — يقدر يضيف صف بياناته هو بس
create policy "users can insert their own profile" on profiles
  for insert with check (id = auth.uid());

-- ---------- factories ----------
create policy "authenticated can view factories" on factories
  for select using (auth.role() = 'authenticated');
create policy "authenticated can manage factories" on factories
  for insert with check (auth.role() = 'authenticated');
create policy "authenticated can delete factories" on factories
  for delete using (auth.role() = 'authenticated');

-- ---------- items ----------
create policy "authenticated can view items" on items
  for select using (auth.role() = 'authenticated');
create policy "authenticated can insert items" on items
  for insert with check (auth.role() = 'authenticated');
create policy "authenticated can update items" on items
  for update using (auth.role() = 'authenticated');
create policy "authenticated can delete items" on items
  for delete using (auth.role() = 'authenticated');

-- ---------- transactions ----------
create policy "authenticated can view transactions" on transactions
  for select using (auth.role() = 'authenticated');
create policy "authenticated can insert transactions" on transactions
  for insert with check (auth.role() = 'authenticated');

-- ---------- orders ----------
create policy "authenticated can view orders" on orders
  for select using (auth.role() = 'authenticated');
create policy "authenticated can insert orders" on orders
  for insert with check (auth.role() = 'authenticated');
create policy "authenticated can update orders" on orders
  for update using (auth.role() = 'authenticated');

-- ============================================================
-- ملاحظة أمان مهمة:
-- كل مستخدم مسجّل دخول (أدمن أو موظف مخزن) عنده صلاحية قراءة/كتابة كاملة على
-- الأصناف والحركات والأوردرات والمصانع من جهة قاعدة البيانات مباشرة.
-- التفرقة بين الأدمن والموظف حاليًا متحكم فيها من الواجهة (الأزرار اللي بتظهر/تتخفي) فقط،
-- مش من قاعدة البيانات نفسها. ده مناسب لفريق صغير موثوق فيه، لكن لو حبيت لاحقًا تفرض
-- الصلاحيات من قاعدة البيانات نفسها (مثلاً يمنع موظف المخزن من حذف صنف حتى لو عدّل الكود)،
-- ده تحسين ممكن نضيفه بعدين باستخدام دالة is_admin() في الـ policies.
-- ============================================================
