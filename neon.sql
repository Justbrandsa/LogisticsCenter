begin;

create extension if not exists pgcrypto with schema public;

create schema if not exists private;
revoke all on schema private from public;

create table if not exists private.app_users (
  id uuid primary key default public.gen_random_uuid(),
  name text not null,
  role text not null,
  password_hash text not null,
  active boolean not null default true,
  phone text,
  vehicle text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists app_users_name_unique
  on private.app_users ((lower(btrim(name))));

alter table private.app_users
  drop constraint if exists app_users_role_check;

alter table private.app_users
  add constraint app_users_role_check
  check (role in ('admin', 'sales', 'driver', 'logistics'));

alter table private.app_users add column if not exists last_known_lat numeric(9, 6);
alter table private.app_users add column if not exists last_known_lng numeric(9, 6);
alter table private.app_users add column if not exists last_known_recorded_at timestamptz;

create table if not exists private.suppliers (
  id uuid primary key default public.gen_random_uuid(),
  name text not null,
  contact_person text not null default '',
  contact_number text not null default '',
  factory boolean not null default false,
  created_by uuid not null references private.app_users(id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists suppliers_name_unique
  on private.suppliers ((lower(btrim(name))));

alter table private.suppliers add column if not exists contact_person text;
alter table private.suppliers add column if not exists contact_number text;

do $$
begin
  if not exists (
    select 1
    from information_schema.columns
    where table_schema = 'private'
      and table_name = 'suppliers'
      and column_name = 'factory'
  ) then
    alter table private.suppliers add column factory boolean;
  elsif exists (
    select 1
    from information_schema.columns
    where table_schema = 'private'
      and table_name = 'suppliers'
      and column_name = 'factory'
      and data_type <> 'boolean'
  ) then
    alter table private.suppliers drop column if exists factory_flag;
    alter table private.suppliers add column factory_flag boolean;

    update private.suppliers
    set factory_flag = case
      when factory is null then false
      when lower(btrim(factory)) in ('', 'false', 'f', '0', 'no', 'n') then false
      else true
    end;

    alter table private.suppliers drop column factory;
    alter table private.suppliers rename column factory_flag to factory;
  end if;
end;
$$;

update private.suppliers
set contact_person = ''
where contact_person is null;

update private.suppliers
set contact_number = ''
where contact_number is null;

update private.suppliers
set factory = false
where factory is null;

alter table private.suppliers alter column contact_person set default '';
alter table private.suppliers alter column contact_number set default '';
alter table private.suppliers alter column factory set default false;
alter table private.suppliers alter column contact_person set not null;
alter table private.suppliers alter column contact_number set not null;
alter table private.suppliers alter column factory set not null;

create table if not exists private.locations (
  id uuid primary key default public.gen_random_uuid(),
  supplier_id uuid references private.suppliers(id) on delete restrict,
  location_type text not null default 'supplier',
  name text not null,
  address text not null,
  lat numeric(9, 6),
  lng numeric(9, 6),
  contact_person text not null default '',
  contact_number text not null default '',
  notes text not null default '',
  created_by uuid not null references private.app_users(id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists locations_supplier_name_unique
  on private.locations (supplier_id, lower(btrim(name)));

alter table private.locations add column if not exists location_type text;
alter table private.locations add column if not exists contact_person text;
alter table private.locations add column if not exists contact_number text;

update private.locations l
set location_type = case when s.factory then 'factory' else 'supplier' end,
    contact_person = coalesce(nullif(btrim(l.contact_person), ''), s.contact_person, ''),
    contact_number = coalesce(nullif(btrim(l.contact_number), ''), s.contact_number, '')
from private.suppliers s
where l.supplier_id = s.id
  and (
    l.location_type is null
    or nullif(btrim(l.location_type), '') is null
    or l.contact_person is null
    or l.contact_number is null
  );

update private.locations
set location_type = 'supplier'
where location_type is null
   or nullif(btrim(location_type), '') is null;

update private.locations
set contact_person = ''
where contact_person is null;

update private.locations
set contact_number = ''
where contact_number is null;

alter table private.locations alter column supplier_id drop not null;
alter table private.locations alter column lat drop not null;
alter table private.locations alter column lng drop not null;
alter table private.locations alter column location_type set default 'supplier';
alter table private.locations alter column contact_person set default '';
alter table private.locations alter column contact_number set default '';
alter table private.locations alter column location_type set not null;
alter table private.locations alter column contact_person set not null;
alter table private.locations alter column contact_number set not null;

alter table private.locations
  drop constraint if exists locations_location_type_check;

alter table private.locations
  drop constraint if exists locations_type_check;

alter table private.locations
  add constraint locations_type_check
  check (location_type in ('supplier', 'factory', 'both'));

create table if not exists private.orders (
  id uuid primary key default public.gen_random_uuid(),
  order_number bigint generated by default as identity unique,
  driver_user_id uuid references private.app_users(id) on delete restrict,
  location_id uuid not null references private.locations(id) on delete restrict,
  entry_type text not null default 'delivery' check (entry_type in ('collection', 'delivery')),
  factory_order_number text not null default '',
  inhouse_order_number text not null default '',
  invoice_number text not null default '',
  po_number text not null default '',
  customer_name text not null,
  delivery_address text not null default '',
  priority text not null check (priority in ('high', 'medium', 'low')),
  notes text not null default '',
  driver_flag_type text check (driver_flag_type in ('not_collected', 'not_ready')),
  driver_flag_note text not null default '',
  driver_flagged_at timestamptz,
  driver_flagged_by_user_id uuid references private.app_users(id) on delete set null,
  picked_up_at timestamptz,
  picked_up_by_user_id uuid references private.app_users(id) on delete set null,
  move_to_factory boolean not null default false,
  factory_destination_location_id uuid references private.locations(id) on delete restrict,
  status text not null default 'active' check (status in ('active', 'completed')),
  scheduled_for date not null,
  original_scheduled_for date not null,
  carry_over_count integer not null default 0 check (carry_over_count >= 0),
  created_by_user_id uuid not null references private.app_users(id) on delete restrict,
  created_at timestamptz not null default now(),
  completed_at timestamptz,
  completion_type text,
  completed_by_user_id uuid references private.app_users(id) on delete set null,
  updated_at timestamptz not null default now()
);

alter table private.orders add column if not exists entry_type text;
alter table private.orders add column if not exists factory_order_number text;
alter table private.orders add column if not exists inhouse_order_number text;
alter table private.orders add column if not exists invoice_number text;
alter table private.orders add column if not exists po_number text;
alter table private.orders add column if not exists branding text;
alter table private.orders add column if not exists stock_description text;
alter table private.orders add column if not exists delivery_address text;
alter table private.orders add column if not exists driver_flag_type text;
alter table private.orders add column if not exists driver_flag_note text;
alter table private.orders add column if not exists driver_flagged_at timestamptz;
alter table private.orders add column if not exists driver_flagged_by_user_id uuid references private.app_users(id) on delete set null;
alter table private.orders add column if not exists picked_up_at timestamptz;
alter table private.orders add column if not exists picked_up_by_user_id uuid references private.app_users(id) on delete set null;
alter table private.orders add column if not exists move_to_factory boolean;
alter table private.orders add column if not exists factory_destination_location_id uuid;
alter table private.orders add column if not exists completion_type text;
alter table private.orders add column if not exists completed_by_user_id uuid references private.app_users(id) on delete set null;

update private.orders
set entry_type = 'delivery'
where entry_type is null
   or nullif(btrim(entry_type), '') is null;

update private.orders
set factory_order_number = concat('LEGACY-', order_number)
where factory_order_number is null
   or nullif(btrim(factory_order_number), '') is null;

update private.orders
set inhouse_order_number = concat('ORD-', order_number)
where inhouse_order_number is null
   or nullif(btrim(inhouse_order_number), '') is null;

update private.orders
set move_to_factory = false
where move_to_factory is null;

update private.orders
set factory_destination_location_id = null
where move_to_factory = false;

update private.orders
set invoice_number = ''
where invoice_number is null;

update private.orders
set po_number = ''
where po_number is null;

update private.orders
set branding = ''
where branding is null;

update private.orders
set stock_description = ''
where stock_description is null;

update private.orders
set delivery_address = ''
where delivery_address is null;

update private.orders
set driver_flag_note = ''
where driver_flag_note is null;

update private.orders
set completion_type = case
      when move_to_factory then 'factory'
      else 'office'
    end
where status = 'completed'
  and completion_type is null;

update private.orders
set completion_type = null,
    completed_by_user_id = null
where status <> 'completed';

alter table private.orders alter column driver_user_id drop not null;
alter table private.orders alter column entry_type set default 'delivery';
alter table private.orders alter column factory_order_number set default '';
alter table private.orders alter column inhouse_order_number set default '';
alter table private.orders alter column invoice_number set default '';
alter table private.orders alter column po_number set default '';
alter table private.orders alter column branding set default '';
alter table private.orders alter column stock_description set default '';
alter table private.orders alter column delivery_address set default '';
alter table private.orders alter column driver_flag_note set default '';
alter table private.orders alter column move_to_factory set default false;
alter table private.orders alter column entry_type set not null;
alter table private.orders alter column factory_order_number set not null;
alter table private.orders alter column inhouse_order_number set not null;
alter table private.orders alter column invoice_number set not null;
alter table private.orders alter column po_number set not null;
alter table private.orders alter column branding set not null;
alter table private.orders alter column stock_description set not null;
alter table private.orders alter column delivery_address set not null;
alter table private.orders alter column driver_flag_note set not null;
alter table private.orders alter column move_to_factory set not null;

alter table private.orders
  drop constraint if exists orders_factory_destination_location_id_fkey;

alter table private.orders
  add constraint orders_factory_destination_location_id_fkey
  foreign key (factory_destination_location_id) references private.locations(id) on delete restrict;

alter table private.orders
  drop constraint if exists orders_driver_flag_type_check;

alter table private.orders
  add constraint orders_driver_flag_type_check
  check (driver_flag_type in ('not_collected', 'not_ready') or driver_flag_type is null);

alter table private.orders
  drop constraint if exists orders_completion_type_check;

alter table private.orders
  add constraint orders_completion_type_check
  check (completion_type in ('office', 'factory') or completion_type is null);

create index if not exists orders_driver_scheduled_idx
  on private.orders (driver_user_id, scheduled_for);

create index if not exists orders_status_scheduled_idx
  on private.orders (status, scheduled_for);

create index if not exists orders_factory_destination_idx
  on private.orders (factory_destination_location_id);

create table if not exists private.stock_items (
  id uuid primary key default public.gen_random_uuid(),
  name text not null,
  sku text not null default '',
  quote_number text not null default '',
  invoice_number text not null default '',
  sales_order_number text not null default '',
  po_number text not null default '',
  unit text not null default 'units',
  notes text not null default '',
  created_source text not null default 'manual',
  created_by_user_id uuid not null references private.app_users(id) on delete restrict,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table private.stock_items add column if not exists quote_number text;
alter table private.stock_items add column if not exists invoice_number text;
alter table private.stock_items add column if not exists sales_order_number text;
alter table private.stock_items add column if not exists po_number text;
alter table private.stock_items add column if not exists created_source text;

update private.stock_items
set quote_number = ''
where quote_number is null;

update private.stock_items
set invoice_number = ''
where invoice_number is null;

update private.stock_items
set sales_order_number = ''
where sales_order_number is null;

update private.stock_items
set po_number = ''
where po_number is null;

update private.stock_items
set created_source = 'manual'
where created_source is null
   or nullif(btrim(created_source), '') is null;

update private.stock_items s
set created_source = 'order'
where not exists (
    select 1
    from private.stock_movements m
    where m.stock_item_id = s.id
  )
  and exists (
    select 1
    from private.orders o
    where lower(btrim(s.name)) = lower(btrim(o.stock_description))
      and lower(btrim(s.quote_number)) = lower(btrim(o.inhouse_order_number))
      and lower(btrim(s.invoice_number)) = lower(btrim(o.invoice_number))
      and lower(btrim(s.sales_order_number)) = lower(btrim(o.factory_order_number))
      and lower(btrim(s.po_number)) = lower(btrim(o.po_number))
  );

alter table private.stock_items alter column quote_number set default '';
alter table private.stock_items alter column invoice_number set default '';
alter table private.stock_items alter column sales_order_number set default '';
alter table private.stock_items alter column po_number set default '';
alter table private.stock_items alter column created_source set default 'manual';
alter table private.stock_items alter column quote_number set not null;
alter table private.stock_items alter column invoice_number set not null;
alter table private.stock_items alter column sales_order_number set not null;
alter table private.stock_items alter column po_number set not null;
alter table private.stock_items alter column created_source set not null;

alter table private.stock_items
  drop constraint if exists stock_items_created_source_check;

alter table private.stock_items
  add constraint stock_items_created_source_check
  check (created_source in ('manual', 'order'));

drop index if exists stock_items_name_unique;
drop index if exists stock_items_quote_name_unique;
drop index if exists stock_items_reference_name_unique;

create unique index if not exists stock_items_reference_name_unique
  on private.stock_items (
    (lower(btrim(name))),
    (lower(btrim(quote_number))),
    (lower(btrim(invoice_number))),
    (lower(btrim(sales_order_number))),
    (lower(btrim(po_number)))
  );

create unique index if not exists stock_items_sku_unique
  on private.stock_items ((lower(btrim(sku))))
  where nullif(btrim(sku), '') is not null;

create table if not exists private.stock_movements (
  id uuid primary key default public.gen_random_uuid(),
  stock_item_id uuid not null references private.stock_items(id) on delete cascade,
  movement_type text not null check (movement_type in ('in', 'out')),
  quantity integer not null check (quantity > 0),
  supplier_name text not null default '',
  driver_user_id uuid references private.app_users(id) on delete restrict,
  notes text not null default '',
  created_by_user_id uuid not null references private.app_users(id) on delete restrict,
  created_at timestamptz not null default now()
);

create index if not exists stock_movements_item_created_idx
  on private.stock_movements (stock_item_id, created_at desc);

create index if not exists stock_movements_driver_created_idx
  on private.stock_movements (driver_user_id, created_at desc);

with seedable_stock_items as (
  select
    s.id,
    regexp_replace(lower(btrim(s.unit)), '\s*units?$', '')::integer as quantity,
    s.created_by_user_id,
    s.created_at
  from private.stock_items s
  where lower(btrim(s.unit)) ~ '^\d+\s*(unit|units)?$'
    and not exists (
      select 1
      from private.stock_movements m
      where m.stock_item_id = s.id
    )
)
insert into private.stock_movements (
  stock_item_id,
  movement_type,
  quantity,
  supplier_name,
  notes,
  created_by_user_id,
  created_at
)
select
  id,
  'in',
  quantity,
  'Opening stock',
  'Opening stock backfilled from the item unit value.',
  created_by_user_id,
  created_at
from seedable_stock_items
where quantity > 0;

update private.stock_items
set unit = 'units'
where lower(btrim(unit)) ~ '^\d+\s*(unit|units)?$';

create table if not exists private.artwork_requests (
  id uuid primary key default public.gen_random_uuid(),
  stock_item_id uuid not null references private.stock_items(id) on delete cascade,
  requested_quantity integer not null check (requested_quantity > 0),
  notes text not null default '',
  sent_to text not null default '',
  requested_by_user_id uuid not null references private.app_users(id) on delete restrict,
  sent_at timestamptz not null default now()
);

create index if not exists artwork_requests_item_sent_idx
  on private.artwork_requests (stock_item_id, sent_at desc);

alter table private.stock_movements
  drop constraint if exists stock_movements_stock_item_id_fkey;

alter table private.stock_movements
  add constraint stock_movements_stock_item_id_fkey
  foreign key (stock_item_id) references private.stock_items(id) on delete cascade;

alter table private.artwork_requests
  drop constraint if exists artwork_requests_stock_item_id_fkey;

alter table private.artwork_requests
  add constraint artwork_requests_stock_item_id_fkey
  foreign key (stock_item_id) references private.stock_items(id) on delete cascade;

create table if not exists private.app_sessions (
  token uuid primary key default public.gen_random_uuid(),
  user_id uuid not null references private.app_users(id) on delete cascade,
  created_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now(),
  expires_at timestamptz not null default (now() + interval '14 days')
);

create index if not exists app_sessions_user_idx
  on private.app_sessions (user_id);

alter table private.app_users enable row level security;
alter table private.suppliers enable row level security;
alter table private.locations enable row level security;
alter table private.orders enable row level security;
alter table private.stock_items enable row level security;
alter table private.stock_movements enable row level security;
alter table private.artwork_requests enable row level security;
alter table private.app_sessions enable row level security;

create or replace function private.touch_updated_at()
returns trigger
language plpgsql
set search_path = ''
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists trg_app_users_touch on private.app_users;
create trigger trg_app_users_touch
before update on private.app_users
for each row
execute procedure private.touch_updated_at();

drop trigger if exists trg_suppliers_touch on private.suppliers;
create trigger trg_suppliers_touch
before update on private.suppliers
for each row
execute procedure private.touch_updated_at();

drop trigger if exists trg_locations_touch on private.locations;
create trigger trg_locations_touch
before update on private.locations
for each row
execute procedure private.touch_updated_at();

drop trigger if exists trg_orders_touch on private.orders;
create trigger trg_orders_touch
before update on private.orders
for each row
execute procedure private.touch_updated_at();

drop trigger if exists trg_stock_items_touch on private.stock_items;
create trigger trg_stock_items_touch
before update on private.stock_items
for each row
execute procedure private.touch_updated_at();

create or replace function private.today_local()
returns date
language sql
stable
set search_path = ''
as $$
  select (now() at time zone 'Africa/Johannesburg')::date;
$$;

create or replace function private.week_start(p_day date)
returns date
language sql
immutable
set search_path = ''
as $$
  select (p_day - ((extract(isodow from p_day)::int) - 1))::date;
$$;

create or replace function private.hash_password(p_password text)
returns text
language sql
strict
set search_path = ''
as $$
  select public.crypt(p_password, public.gen_salt('bf'));
$$;

create or replace function private.password_matches(p_password text, p_hash text)
returns boolean
language sql
strict
set search_path = ''
as $$
  select p_hash = public.crypt(p_password, p_hash);
$$;

create or replace function private.stock_on_hand(p_stock_item_id uuid)
returns integer
language sql
stable
set search_path = ''
as $$
  select coalesce(
    sum(case when movement_type = 'in' then quantity else -quantity end),
    0
  )::integer
  from private.stock_movements
  where stock_item_id = p_stock_item_id;
$$;

create or replace function private.build_user_json(p_user private.app_users)
returns jsonb
language sql
stable
set search_path = ''
as $$
  select jsonb_build_object(
    'id', p_user.id,
    'name', p_user.name,
    'role', p_user.role,
    'active', p_user.active,
    'phone', coalesce(p_user.phone, ''),
    'vehicle', coalesce(p_user.vehicle, ''),
    'lastKnownLat', p_user.last_known_lat,
    'lastKnownLng', p_user.last_known_lng,
    'lastKnownRecordedAt', p_user.last_known_recorded_at,
    'createdAt', p_user.created_at
  );
$$;

create or replace function private.issue_session(p_user_id uuid)
returns uuid
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_token uuid := public.gen_random_uuid();
begin
  delete from private.app_sessions
  where user_id = p_user_id
    and expires_at < now();

  insert into private.app_sessions (token, user_id)
  values (v_token, p_user_id);

  return v_token;
end;
$$;

create or replace function private.require_user(p_token uuid, p_roles text[] default null)
returns private.app_users
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_user private.app_users;
begin
  select u.*
  into v_user
  from private.app_sessions s
  join private.app_users u on u.id = s.user_id
  where s.token = p_token
    and s.expires_at > now()
    and u.active
  order by s.created_at desc
  limit 1;

  if v_user.id is null then
    raise exception 'Invalid session';
  end if;

  if p_roles is not null and not (v_user.role = any (p_roles)) then
    raise exception 'Permission denied';
  end if;

  update private.app_sessions
  set last_seen_at = now()
  where token = p_token;

  return v_user;
end;
$$;

drop function if exists private.roll_forward_open_orders(date);

create or replace function private.roll_forward_open_orders(p_today date default private.today_local())
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_result jsonb := jsonb_build_object(
    'today', p_today,
    'updatedOrders', 0,
    'carriedOrders', '[]'::jsonb
  );
begin
  with carried_orders as (
    select
      o.id,
      concat('ORD-', o.order_number) as reference,
      o.order_number,
      o.entry_type,
      o.move_to_factory,
      o.factory_destination_location_id,
      coalesce(f.name, '') as factory_destination_name,
      coalesce(f.address, '') as factory_destination_address,
      o.factory_order_number,
      o.inhouse_order_number,
      o.factory_order_number as sales_order_number,
      o.inhouse_order_number as quote_number,
      o.invoice_number,
      o.po_number,
      o.delivery_address,
      o.customer_name,
      o.driver_user_id,
      coalesce(d.name, '') as driver_name,
      o.location_id,
      l.name as location_name,
      l.address as location_address,
      l.location_type,
      l.contact_person as location_contact_person,
      l.contact_number as location_contact_number,
      'high'::text as priority,
      o.notes,
      o.driver_flag_type,
      o.driver_flag_note,
      o.driver_flagged_at,
      o.driver_flagged_by_user_id,
      coalesce(g.name, '') as driver_flagged_by_name,
      o.picked_up_at,
      o.picked_up_by_user_id,
      coalesce(i.name, '') as picked_up_by_name,
      o.status,
      o.scheduled_for as previous_scheduled_for,
      p_today as scheduled_for,
      o.original_scheduled_for,
      o.carry_over_count + greatest((p_today - o.scheduled_for), 1) as carry_over_count,
      o.created_by_user_id,
      c.name as created_by_name,
      c.role as created_by_role,
      o.created_at,
      o.completed_at,
      o.completion_type,
      o.completed_by_user_id,
      coalesce(h.name, '') as completed_by_name
    from private.orders o
    left join private.app_users d on d.id = o.driver_user_id
    join private.locations l on l.id = o.location_id
    left join private.locations f on f.id = o.factory_destination_location_id
    left join private.app_users g on g.id = o.driver_flagged_by_user_id
    left join private.app_users h on h.id = o.completed_by_user_id
    left join private.app_users i on i.id = o.picked_up_by_user_id
    join private.app_users c on c.id = o.created_by_user_id
    where o.status = 'active'
      and o.driver_user_id is not null
      and o.scheduled_for < p_today
  ),
  updated_orders as (
    update private.orders o
    set scheduled_for = p_today,
        priority = 'high',
        carry_over_count = o.carry_over_count + greatest((p_today - o.scheduled_for), 1),
        updated_at = now()
    from carried_orders carried
    where o.id = carried.id
    returning
      jsonb_build_object(
        'id', carried.id,
        'reference', carried.reference,
        'orderNumber', carried.order_number,
        'entryType', carried.entry_type,
        'moveToFactory', carried.move_to_factory,
        'factoryDestinationLocationId', carried.factory_destination_location_id,
        'factoryDestinationName', carried.factory_destination_name,
        'factoryDestinationAddress', carried.factory_destination_address,
        'factoryOrderNumber', carried.factory_order_number,
        'inhouseOrderNumber', carried.inhouse_order_number,
        'salesOrderNumber', carried.sales_order_number,
        'quoteNumber', carried.quote_number,
        'invoiceNumber', carried.invoice_number,
        'poNumber', carried.po_number,
        'deliveryAddress', carried.delivery_address,
        'customerName', carried.customer_name,
        'driverUserId', carried.driver_user_id,
        'driverName', carried.driver_name,
        'locationId', carried.location_id,
        'locationName', carried.location_name,
        'locationAddress', carried.location_address,
        'locationType', carried.location_type,
        'locationContactPerson', carried.location_contact_person,
        'locationContactNumber', carried.location_contact_number,
        'priority', carried.priority,
        'notes', carried.notes,
        'driverFlagType', carried.driver_flag_type,
        'driverFlagNote', carried.driver_flag_note,
        'driverFlaggedAt', carried.driver_flagged_at,
        'driverFlaggedByUserId', carried.driver_flagged_by_user_id,
        'driverFlaggedByName', carried.driver_flagged_by_name,
        'pickedUpAt', carried.picked_up_at,
        'pickedUpByUserId', carried.picked_up_by_user_id,
        'pickedUpByName', carried.picked_up_by_name,
        'completionType', carried.completion_type,
        'completedByUserId', carried.completed_by_user_id,
        'completedByName', carried.completed_by_name,
        'status', carried.status,
        'previousScheduledFor', carried.previous_scheduled_for,
        'scheduledFor', carried.scheduled_for,
        'originalScheduledFor', carried.original_scheduled_for,
        'carryOverCount', carried.carry_over_count,
        'createdByUserId', carried.created_by_user_id,
        'createdByName', carried.created_by_name,
        'createdByRole', carried.created_by_role,
        'createdAt', carried.created_at,
        'completedAt', carried.completed_at
      ) as order_json,
      carried.driver_name,
      carried.location_name,
      carried.order_number
  )
  select jsonb_build_object(
    'today', p_today,
    'updatedOrders', count(*),
    'carriedOrders',
      coalesce(
        jsonb_agg(order_json order by lower(driver_name), lower(location_name), order_number),
        '[]'::jsonb
      )
  )
  into v_result
  from updated_orders;

  return coalesce(
    v_result,
    jsonb_build_object(
      'today', p_today,
      'updatedOrders', 0,
      'carriedOrders', '[]'::jsonb
    )
  );
end;
$$;

create or replace function public.get_login_state()
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_today date := private.today_local();
  v_week_start date := private.week_start(v_today);
begin
  return jsonb_build_object(
    'today', v_today,
    'weekStart', v_week_start,
    'weekEnd', (v_week_start + 6),
    'hasUsers', exists(select 1 from private.app_users)
  );
end;
$$;

create or replace function public.bootstrap_admin(p_name text, p_password text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_name text := nullif(btrim(p_name), '');
  v_password text := nullif(btrim(p_password), '');
  v_user private.app_users;
  v_token uuid;
begin
  if exists(select 1 from private.app_users) then
    raise exception 'The first admin account already exists.';
  end if;

  if v_name is null then
    raise exception 'Name is required.';
  end if;

  if v_password is null or length(v_password) < 4 then
    raise exception 'Password must be at least 4 characters long.';
  end if;

  insert into private.app_users (name, role, password_hash)
  values (v_name, 'admin', private.hash_password(v_password))
  returning * into v_user;

  v_token := private.issue_session(v_user.id);

  return jsonb_build_object(
    'token', v_token,
    'user', private.build_user_json(v_user)
  );
end;
$$;

create or replace function public.login_user(p_name text, p_password text)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_name text := nullif(btrim(p_name), '');
  v_password text := nullif(btrim(p_password), '');
  v_user private.app_users;
  v_token uuid;
begin
  if v_name is null or v_password is null then
    raise exception 'Name and password are required.';
  end if;

  select *
  into v_user
  from private.app_users
  where lower(btrim(name)) = lower(v_name)
  limit 1;

  if v_user.id is null or not private.password_matches(v_password, v_user.password_hash) then
    raise exception 'Invalid name or password.';
  end if;

  if not v_user.active then
    raise exception 'This account is inactive.';
  end if;

  v_token := private.issue_session(v_user.id);

  return jsonb_build_object(
    'token', v_token,
    'user', private.build_user_json(v_user)
  );
end;
$$;

create or replace function public.logout_user(p_token uuid)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
begin
  delete from private.app_sessions where token = p_token;
  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.run_daily_rollover()
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_today date := private.today_local();
  v_rollover jsonb;
begin
  v_rollover := private.roll_forward_open_orders(v_today);
  return coalesce(
    v_rollover,
    jsonb_build_object(
      'today', v_today,
      'updatedOrders', 0,
      'carriedOrders', '[]'::jsonb
    )
  );
end;
$$;

create or replace function public.get_app_snapshot(p_token uuid)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_today date := private.today_local();
  v_rollover jsonb := jsonb_build_object(
    'today', v_today,
    'updatedOrders', 0,
    'carriedOrders', '[]'::jsonb
  );
  v_users jsonb := '[]'::jsonb;
  v_suppliers jsonb := '[]'::jsonb;
  v_locations jsonb := '[]'::jsonb;
  v_orders jsonb := '[]'::jsonb;
  v_stock_items jsonb := '[]'::jsonb;
  v_stock_movements jsonb := '[]'::jsonb;
  v_artwork_requests jsonb := '[]'::jsonb;
begin
  v_actor := private.require_user(p_token);
  v_rollover := private.roll_forward_open_orders(v_today);

  if v_actor.role = 'admin' then
    select coalesce(
      jsonb_agg(private.build_user_json(u) order by lower(u.name)),
      '[]'::jsonb
    )
    into v_users
    from private.app_users u;
  elsif v_actor.role in ('sales', 'logistics', 'driver') then
    select coalesce(
      jsonb_agg(private.build_user_json(u) order by lower(u.name)),
      '[]'::jsonb
    )
    into v_users
    from private.app_users u
    where u.role = 'driver'
      and u.active;
  end if;

  if v_actor.role in ('admin', 'sales') then
    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', s.id,
          'name', s.name,
          'contactPerson', s.contact_person,
          'contactNumber', s.contact_number,
          'factory', s.factory,
          'createdAt', s.created_at
        )
        order by lower(s.name)
      ),
      '[]'::jsonb
    )
    into v_suppliers
    from private.suppliers s;

    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', l.id,
          'locationType', l.location_type,
          'name', l.name,
          'address', l.address,
          'lat', l.lat,
          'lng', l.lng,
          'contactPerson', l.contact_person,
          'contactNumber', l.contact_number,
          'notes', l.notes,
          'createdAt', l.created_at
        )
        order by lower(l.name)
      ),
      '[]'::jsonb
    )
    into v_locations
    from private.locations l;

    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', o.id,
          'reference', concat('ORD-', o.order_number),
          'orderNumber', o.order_number,
          'entryType', o.entry_type,
          'moveToFactory', o.move_to_factory,
          'factoryDestinationLocationId', o.factory_destination_location_id,
          'factoryDestinationName', coalesce(f.name, ''),
          'factoryDestinationAddress', coalesce(f.address, ''),
          'factoryOrderNumber', o.factory_order_number,
          'inhouseOrderNumber', o.inhouse_order_number,
          'salesOrderNumber', o.factory_order_number,
          'quoteNumber', o.inhouse_order_number,
          'invoiceNumber', o.invoice_number,
          'poNumber', o.po_number,
          'deliveryAddress', o.delivery_address,
          'branding', o.branding,
          'stockDescription', o.stock_description,
          'customerName', o.customer_name,
          'driverUserId', o.driver_user_id,
          'driverName', coalesce(d.name, ''),
          'locationId', o.location_id,
          'locationName', l.name,
          'locationAddress', l.address,
          'locationType', l.location_type,
          'locationContactPerson', l.contact_person,
          'locationContactNumber', l.contact_number,
          'priority', o.priority,
          'notes', o.notes,
          'driverFlagType', o.driver_flag_type,
          'driverFlagNote', o.driver_flag_note,
          'driverFlaggedAt', o.driver_flagged_at,
          'driverFlaggedByUserId', o.driver_flagged_by_user_id,
          'driverFlaggedByName', coalesce(g.name, ''),
          'pickedUpAt', o.picked_up_at,
          'pickedUpByUserId', o.picked_up_by_user_id,
          'pickedUpByName', coalesce(i.name, ''),
          'completionType', o.completion_type,
          'completedByUserId', o.completed_by_user_id,
          'completedByName', coalesce(h.name, ''),
          'status', o.status,
          'scheduledFor', o.scheduled_for,
          'originalScheduledFor', o.original_scheduled_for,
          'carryOverCount', o.carry_over_count,
          'createdByUserId', o.created_by_user_id,
          'createdByName', c.name,
          'createdByRole', c.role,
          'createdAt', o.created_at,
          'completedAt', o.completed_at
        )
        order by
          case when o.status = 'active' then 0 else 1 end,
          o.created_at desc,
          o.order_number desc
      ),
      '[]'::jsonb
    )
    into v_orders
    from private.orders o
    left join private.app_users d on d.id = o.driver_user_id
    join private.locations l on l.id = o.location_id
    left join private.locations f on f.id = o.factory_destination_location_id
    left join private.app_users g on g.id = o.driver_flagged_by_user_id
    left join private.app_users h on h.id = o.completed_by_user_id
    left join private.app_users i on i.id = o.picked_up_by_user_id
    join private.app_users c on c.id = o.created_by_user_id
    ;
  else
    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', s.id,
          'name', s.name,
          'contactPerson', s.contact_person,
          'contactNumber', s.contact_number,
          'factory', s.factory,
          'createdAt', s.created_at
        )
        order by lower(s.name)
      ),
      '[]'::jsonb
    )
    into v_suppliers
    from private.suppliers s
    where exists (
      select 1
      from private.locations l
      join private.orders o on o.location_id = l.id
      where l.supplier_id = s.id
        and o.driver_user_id = v_actor.id
    );

    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', l.id,
          'locationType', l.location_type,
          'name', l.name,
          'address', l.address,
          'lat', l.lat,
          'lng', l.lng,
          'contactPerson', l.contact_person,
          'contactNumber', l.contact_number,
          'notes', l.notes,
          'createdAt', l.created_at
        )
        order by lower(l.name)
      ),
      '[]'::jsonb
    )
    into v_locations
    from private.locations l
    where exists (
      select 1
      from private.orders o
      where o.location_id = l.id
        and o.driver_user_id = v_actor.id
    );

    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', o.id,
          'reference', concat('ORD-', o.order_number),
          'orderNumber', o.order_number,
          'entryType', o.entry_type,
          'moveToFactory', o.move_to_factory,
          'factoryDestinationLocationId', o.factory_destination_location_id,
          'factoryDestinationName', coalesce(f.name, ''),
          'factoryDestinationAddress', coalesce(f.address, ''),
          'factoryOrderNumber', o.factory_order_number,
          'inhouseOrderNumber', o.inhouse_order_number,
          'salesOrderNumber', o.factory_order_number,
          'quoteNumber', o.inhouse_order_number,
          'invoiceNumber', o.invoice_number,
          'poNumber', o.po_number,
          'deliveryAddress', o.delivery_address,
          'branding', o.branding,
          'stockDescription', o.stock_description,
          'customerName', o.customer_name,
          'driverUserId', o.driver_user_id,
          'driverName', v_actor.name,
          'locationId', o.location_id,
          'locationName', l.name,
          'locationAddress', l.address,
          'locationType', l.location_type,
          'locationContactPerson', l.contact_person,
          'locationContactNumber', l.contact_number,
          'priority', o.priority,
          'notes', o.notes,
          'driverFlagType', o.driver_flag_type,
          'driverFlagNote', o.driver_flag_note,
          'driverFlaggedAt', o.driver_flagged_at,
          'driverFlaggedByUserId', o.driver_flagged_by_user_id,
          'driverFlaggedByName', coalesce(g.name, ''),
          'pickedUpAt', o.picked_up_at,
          'pickedUpByUserId', o.picked_up_by_user_id,
          'pickedUpByName', coalesce(i.name, ''),
          'completionType', o.completion_type,
          'completedByUserId', o.completed_by_user_id,
          'completedByName', coalesce(h.name, ''),
          'status', o.status,
          'scheduledFor', o.scheduled_for,
          'originalScheduledFor', o.original_scheduled_for,
          'carryOverCount', o.carry_over_count,
          'createdByUserId', o.created_by_user_id,
          'createdByName', c.name,
          'createdByRole', c.role,
          'createdAt', o.created_at,
          'completedAt', o.completed_at
        )
        order by
          case when o.status = 'active' then 0 else 1 end,
          o.created_at desc,
          o.order_number desc
      ),
      '[]'::jsonb
    )
    into v_orders
    from private.orders o
    join private.locations l on l.id = o.location_id
    left join private.locations f on f.id = o.factory_destination_location_id
    left join private.app_users g on g.id = o.driver_flagged_by_user_id
    left join private.app_users h on h.id = o.completed_by_user_id
    left join private.app_users i on i.id = o.picked_up_by_user_id
    join private.app_users c on c.id = o.created_by_user_id
    where o.driver_user_id = v_actor.id
    ;
  end if;

  if v_actor.role in ('admin', 'logistics', 'sales') then
    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', s.id,
          'name', s.name,
          'sku', s.sku,
          'quoteNumber', s.quote_number,
          'invoiceNumber', s.invoice_number,
          'salesOrderNumber', s.sales_order_number,
          'poNumber', s.po_number,
          'unit', s.unit,
          'notes', s.notes,
          'createdSource', s.created_source,
          'onHandQuantity', private.stock_on_hand(s.id),
          'createdAt', s.created_at,
          'updatedAt', s.updated_at
        )
        order by lower(s.name)
      ),
      '[]'::jsonb
    )
    into v_stock_items
    from private.stock_items s;

    select coalesce(
      jsonb_agg(
        jsonb_build_object(
          'id', m.id,
          'stockItemId', m.stock_item_id,
          'itemName', s.name,
          'sku', s.sku,
          'quoteNumber', s.quote_number,
          'invoiceNumber', s.invoice_number,
          'salesOrderNumber', s.sales_order_number,
          'poNumber', s.po_number,
          'unit', s.unit,
          'movementType', m.movement_type,
          'quantity', m.quantity,
          'supplierName', m.supplier_name,
          'driverUserId', m.driver_user_id,
          'driverName', coalesce(d.name, ''),
          'notes', m.notes,
          'createdByUserId', m.created_by_user_id,
          'createdByName', c.name,
          'createdAt', m.created_at
        )
        order by m.created_at desc
      ),
      '[]'::jsonb
    )
    into v_stock_movements
    from private.stock_movements m
    join private.stock_items s on s.id = m.stock_item_id
    left join private.app_users d on d.id = m.driver_user_id
    join private.app_users c on c.id = m.created_by_user_id;

    if v_actor.role in ('admin', 'logistics') then
      select coalesce(
        jsonb_agg(
          jsonb_build_object(
            'id', r.id,
            'stockItemId', r.stock_item_id,
            'itemName', s.name,
            'sku', s.sku,
            'quoteNumber', s.quote_number,
            'invoiceNumber', s.invoice_number,
            'salesOrderNumber', s.sales_order_number,
            'poNumber', s.po_number,
            'requestedQuantity', r.requested_quantity,
            'notes', r.notes,
            'sentTo', r.sent_to,
            'requestedByUserId', r.requested_by_user_id,
            'requestedByName', u.name,
            'sentAt', r.sent_at
          )
          order by r.sent_at desc
        ),
        '[]'::jsonb
      )
      into v_artwork_requests
      from private.artwork_requests r
      join private.stock_items s on s.id = r.stock_item_id
      join private.app_users u on u.id = r.requested_by_user_id;
    end if;
  end if;

  return jsonb_build_object(
    'today', v_today,
    'rollover', coalesce(v_rollover, jsonb_build_object(
      'today', v_today,
      'updatedOrders', 0,
      'carriedOrders', '[]'::jsonb
    )),
    'user', private.build_user_json(v_actor),
    'users', v_users,
    'suppliers', v_suppliers,
    'locations', v_locations,
    'orders', v_orders,
    'stockItems', v_stock_items,
    'stockMovements', v_stock_movements,
    'artworkRequests', v_artwork_requests
  );
end;
$$;

create or replace function public.record_driver_position(
  p_token uuid,
  p_lat numeric,
  p_lng numeric
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_recorded_at timestamptz := now();
begin
  v_actor := private.require_user(p_token, array['driver']);

  if p_lat is null or p_lng is null then
    raise exception 'Latitude and longitude are required.';
  end if;

  if p_lat < -90 or p_lat > 90 then
    raise exception 'Latitude must be between -90 and 90.';
  end if;

  if p_lng < -180 or p_lng > 180 then
    raise exception 'Longitude must be between -180 and 180.';
  end if;

  update private.app_users
  set last_known_lat = p_lat,
      last_known_lng = p_lng,
      last_known_recorded_at = v_recorded_at
  where id = v_actor.id;

  return jsonb_build_object(
    'ok', true,
    'recordedAt', v_recorded_at
  );
end;
$$;

create or replace function public.create_user_account(
  p_token uuid,
  p_name text,
  p_password text,
  p_role text,
  p_phone text default null,
  p_vehicle text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_name text := nullif(btrim(p_name), '');
  v_password text := nullif(btrim(p_password), '');
  v_role text := lower(nullif(btrim(p_role), ''));
begin
  v_actor := private.require_user(p_token, array['admin']);

  if v_name is null then
    raise exception 'Name is required.';
  end if;

  if v_password is null or length(v_password) < 4 then
    raise exception 'Password must be at least 4 characters long.';
  end if;

  if v_role not in ('admin', 'sales', 'driver', 'logistics') then
    raise exception 'Invalid role.';
  end if;

  if exists (
    select 1
    from private.app_users
    where lower(btrim(name)) = lower(v_name)
  ) then
    raise exception 'That name is already in use.';
  end if;

  if v_role = 'driver' and nullif(btrim(p_phone), '') is null then
    raise exception 'Driver accounts require a phone number.';
  end if;

  insert into private.app_users (
    name,
    role,
    password_hash,
    phone,
    vehicle
  )
  values (
    v_name,
    v_role,
    private.hash_password(v_password),
    case when v_role = 'driver' then nullif(btrim(p_phone), '') else null end,
    null
  );

  return jsonb_build_object('ok', true, 'createdBy', v_actor.id);
end;
$$;

drop function if exists public.update_user_account(uuid, uuid, text, text, text);
drop function if exists public.update_user_account(uuid, uuid, text, text, text, text);

create or replace function public.update_user_account(
  p_token uuid,
  p_user_id uuid,
  p_name text,
  p_role text,
  p_phone text default null,
  p_password text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_target private.app_users;
  v_name text := nullif(btrim(p_name), '');
  v_role text := lower(nullif(btrim(p_role), ''));
  v_phone text := nullif(btrim(p_phone), '');
  v_password text := nullif(btrim(p_password), '');
begin
  v_actor := private.require_user(p_token, array['admin']);

  select *
  into v_target
  from private.app_users
  where id = p_user_id;

  if v_target.id is null then
    raise exception 'User not found.';
  end if;

  if v_name is null then
    raise exception 'Name is required.';
  end if;

  if v_role not in ('admin', 'sales', 'driver', 'logistics') then
    raise exception 'Invalid role.';
  end if;

  if v_password is not null and length(v_password) < 4 then
    raise exception 'Password must be at least 4 characters long.';
  end if;

  if exists (
    select 1
    from private.app_users
    where id <> p_user_id
      and lower(btrim(name)) = lower(v_name)
  ) then
    raise exception 'That name is already in use.';
  end if;

  if v_role = 'driver' and v_phone is null then
    raise exception 'Driver accounts require a phone number.';
  end if;

  if v_target.id = v_actor.id and v_role <> v_actor.role then
    raise exception 'You cannot change your own role.';
  end if;

  update private.app_users
  set name = v_name,
      role = v_role,
      phone = case when v_role = 'driver' then v_phone else null end,
      password_hash = case
        when v_password is null then password_hash
        else private.hash_password(v_password)
      end
  where id = p_user_id;

  return jsonb_build_object('ok', true, 'updatedBy', v_actor.id);
end;
$$;

create or replace function public.toggle_user_active(
  p_token uuid,
  p_user_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_target private.app_users;
begin
  v_actor := private.require_user(p_token, array['admin']);

  select * into v_target
  from private.app_users
  where id = p_user_id;

  if v_target.id is null then
    raise exception 'User not found.';
  end if;

  if v_target.id = v_actor.id then
    raise exception 'You cannot change your own active status.';
  end if;

  update private.app_users
  set active = not active
  where id = v_target.id;

  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.delete_user_account(
  p_token uuid,
  p_user_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_target private.app_users;
begin
  v_actor := private.require_user(p_token, array['admin']);

  select * into v_target
  from private.app_users
  where id = p_user_id;

  if v_target.id is null then
    raise exception 'User not found.';
  end if;

  if v_target.id = v_actor.id then
    raise exception 'You cannot delete your own account.';
  end if;

  if exists (
    select 1
    from private.orders
    where driver_user_id = v_target.id
       or created_by_user_id = v_target.id
  ) then
    raise exception 'This account has order history. Disable it instead of deleting it.';
  end if;

  delete from private.app_sessions where user_id = v_target.id;
  delete from private.app_users where id = v_target.id;

  return jsonb_build_object('ok', true);
end;
$$;

drop function if exists public.create_supplier(uuid, text);
drop function if exists public.create_supplier(uuid, text, text, text);
drop function if exists public.create_supplier(uuid, text, text, boolean);
drop function if exists public.create_supplier(uuid, text, text, text, boolean);

create or replace function public.create_supplier(
  p_token uuid,
  p_name text,
  p_contact_person text,
  p_contact_number text,
  p_factory boolean
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_name text := nullif(btrim(p_name), '');
  v_contact_person text := coalesce(nullif(btrim(p_contact_person), ''), '');
  v_contact_number text := nullif(btrim(p_contact_number), '');
  v_factory boolean := coalesce(p_factory, false);
begin
  v_actor := private.require_user(p_token, array['admin']);

  if v_name is null or v_contact_number is null then
    raise exception 'Supplier name and contact number are required.';
  end if;

  if exists (
    select 1
    from private.suppliers
    where lower(btrim(name)) = lower(v_name)
  ) then
    raise exception 'That supplier already exists.';
  end if;

  insert into private.suppliers (name, contact_person, contact_number, factory, created_by)
  values (v_name, v_contact_person, v_contact_number, v_factory, v_actor.id);

  return jsonb_build_object('ok', true);
end;
$$;

drop function if exists public.update_supplier(uuid, uuid, text, text, text);
drop function if exists public.update_supplier(uuid, uuid, text, text, boolean);
drop function if exists public.update_supplier(uuid, uuid, text, text, text, boolean);

create or replace function public.update_supplier(
  p_token uuid,
  p_supplier_id uuid,
  p_name text,
  p_contact_person text,
  p_contact_number text,
  p_factory boolean
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_name text := nullif(btrim(p_name), '');
  v_contact_person text := coalesce(nullif(btrim(p_contact_person), ''), '');
  v_contact_number text := nullif(btrim(p_contact_number), '');
  v_factory boolean := coalesce(p_factory, false);
begin
  v_actor := private.require_user(p_token, array['admin']);

  if v_name is null or v_contact_number is null then
    raise exception 'Supplier name and contact number are required.';
  end if;

  if not exists (
    select 1
    from private.suppliers
    where id = p_supplier_id
  ) then
    raise exception 'Supplier not found.';
  end if;

  if exists (
    select 1
    from private.suppliers
    where id <> p_supplier_id
      and lower(btrim(name)) = lower(v_name)
  ) then
    raise exception 'That supplier already exists.';
  end if;

  update private.suppliers
  set name = v_name,
      contact_person = v_contact_person,
      contact_number = v_contact_number,
      factory = v_factory
  where id = p_supplier_id;

  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.delete_supplier(
  p_token uuid,
  p_supplier_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
begin
  v_actor := private.require_user(p_token, array['admin']);

  if exists (
    select 1
    from private.locations
    where supplier_id = p_supplier_id
  ) then
    raise exception 'Delete supplier locations first.';
  end if;

  delete from private.suppliers
  where id = p_supplier_id;

  return jsonb_build_object('ok', true, 'deletedBy', v_actor.id);
end;
$$;

drop function if exists public.create_location(uuid, uuid, text, text, numeric, numeric, text);
drop function if exists public.create_location(uuid, text, text, text, numeric, numeric, text, text);
drop function if exists public.update_location(uuid, uuid, text, text, text, numeric, numeric, text, text);

create or replace function public.create_location(
  p_token uuid,
  p_location_type text,
  p_name text,
  p_address text,
  p_lat numeric,
  p_lng numeric,
  p_contact_person text,
  p_contact_number text
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_location_type text := lower(nullif(btrim(p_location_type), ''));
  v_name text := nullif(btrim(p_name), '');
  v_address text := nullif(btrim(p_address), '');
  v_contact_person text := coalesce(nullif(btrim(p_contact_person), ''), '');
  v_contact_number text := coalesce(nullif(btrim(p_contact_number), ''), '');
begin
  v_actor := private.require_user(p_token, array['admin']);

  if v_location_type not in ('supplier', 'factory', 'both') then
    raise exception 'Location type must be supplier, factory, or both.';
  end if;

  if v_name is null or v_address is null then
    raise exception 'Location name and address are required.';
  end if;

  if (p_lat is null) <> (p_lng is null) then
    raise exception 'Latitude and longitude must both be provided, or both left blank.';
  end if;

  insert into private.locations (
    supplier_id,
    location_type,
    name,
    address,
    lat,
    lng,
    contact_person,
    contact_number,
    notes,
    created_by
  )
  values (
    null,
    v_location_type,
    v_name,
    v_address,
    p_lat,
    p_lng,
    v_contact_person,
    v_contact_number,
    '',
    v_actor.id
  );

  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.update_location(
  p_token uuid,
  p_location_id uuid,
  p_location_type text,
  p_name text,
  p_address text,
  p_lat numeric,
  p_lng numeric,
  p_contact_person text,
  p_contact_number text
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_location_type text := lower(nullif(btrim(p_location_type), ''));
  v_name text := nullif(btrim(p_name), '');
  v_address text := nullif(btrim(p_address), '');
  v_contact_person text := coalesce(nullif(btrim(p_contact_person), ''), '');
  v_contact_number text := coalesce(nullif(btrim(p_contact_number), ''), '');
begin
  v_actor := private.require_user(p_token, array['admin']);

  if v_location_type not in ('supplier', 'factory', 'both') then
    raise exception 'Location type must be supplier, factory, or both.';
  end if;

  if v_name is null or v_address is null then
    raise exception 'Location name and address are required.';
  end if;

  if (p_lat is null) <> (p_lng is null) then
    raise exception 'Latitude and longitude must both be provided, or both left blank.';
  end if;

  if not exists (
    select 1
    from private.locations
    where id = p_location_id
  ) then
    raise exception 'Location not found.';
  end if;

  update private.locations
  set location_type = v_location_type,
      name = v_name,
      address = v_address,
      lat = p_lat,
      lng = p_lng,
      contact_person = v_contact_person,
      contact_number = v_contact_number
  where id = p_location_id;

  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.delete_location(
  p_token uuid,
  p_location_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
begin
  v_actor := private.require_user(p_token, array['admin']);

  if exists (
    select 1
    from private.orders
    where location_id = p_location_id
  ) then
    raise exception 'This location has order history and cannot be deleted.';
  end if;

  delete from private.locations
  where id = p_location_id;

  return jsonb_build_object('ok', true, 'deletedBy', v_actor.id);
end;
$$;

drop function if exists public.create_stock_item(uuid, text, text, text, text);
drop function if exists public.create_stock_item(uuid, text, text, text, text, text, text, text);
drop function if exists public.create_stock_item(uuid, text, text, text, text, text, text, text, text);
drop function if exists public.create_stock_item(uuid, text, text, text, text, text, text, text, integer, text);

create or replace function public.create_stock_item(
  p_token uuid,
  p_name text,
  p_sku text default null,
  p_quote_number text default null,
  p_invoice_number text default null,
  p_sales_order_number text default null,
  p_po_number text default null,
  p_unit text default null,
  p_initial_quantity integer default null,
  p_notes text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_name text := nullif(btrim(p_name), '');
  v_sku text := coalesce(nullif(btrim(p_sku), ''), '');
  v_quote_number text := coalesce(nullif(btrim(p_quote_number), ''), '');
  v_invoice_number text := coalesce(nullif(btrim(p_invoice_number), ''), '');
  v_sales_order_number text := coalesce(nullif(btrim(p_sales_order_number), ''), '');
  v_po_number text := coalesce(nullif(btrim(p_po_number), ''), '');
  v_unit text := coalesce(nullif(btrim(p_unit), ''), 'units');
  v_initial_quantity integer := coalesce(p_initial_quantity, 0);
  v_notes text := coalesce(nullif(btrim(p_notes), ''), '');
  v_stock_item_id uuid;
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'logistics') then
    raise exception 'Permission denied';
  end if;

  if v_name is null then
    raise exception 'Stock item name is required.';
  end if;

  if v_quote_number = ''
     and v_invoice_number = ''
     and v_sales_order_number = ''
     and v_po_number = '' then
    raise exception 'Enter at least one quote, sales order, invoice, or PO number.';
  end if;

  if v_initial_quantity <= 0 then
    raise exception 'Opening stock quantity must be greater than zero.';
  end if;

  if exists (
    select 1
    from private.stock_items
    where lower(btrim(name)) = lower(v_name)
      and lower(btrim(quote_number)) = lower(v_quote_number)
      and lower(btrim(invoice_number)) = lower(v_invoice_number)
      and lower(btrim(sales_order_number)) = lower(v_sales_order_number)
      and lower(btrim(po_number)) = lower(v_po_number)
  ) then
    raise exception 'That stock item already exists for the same reference combination.';
  end if;

  if v_sku <> '' and exists (
    select 1
    from private.stock_items
    where lower(btrim(sku)) = lower(v_sku)
  ) then
    raise exception 'That stock code is already in use.';
  end if;

  insert into private.stock_items (
    name,
    sku,
    quote_number,
    invoice_number,
    sales_order_number,
    po_number,
    unit,
    notes,
    created_source,
    created_by_user_id
  )
  values (
    v_name,
    v_sku,
    v_quote_number,
    v_invoice_number,
    v_sales_order_number,
    v_po_number,
    v_unit,
    v_notes,
    'manual',
    v_actor.id
  )
  returning id into v_stock_item_id;

  insert into private.stock_movements (
    stock_item_id,
    movement_type,
    quantity,
    supplier_name,
    notes,
    created_by_user_id
  )
  values (
    v_stock_item_id,
    'in',
    v_initial_quantity,
    'Opening stock',
    'Opening stock logged when the item was created.',
    v_actor.id
  );

  return jsonb_build_object('ok', true, 'stockItemId', v_stock_item_id);
end;
$$;

drop function if exists public.update_stock_item(uuid, uuid, text, text, text, text, text, text, text, text);

create or replace function public.update_stock_item(
  p_token uuid,
  p_stock_item_id uuid,
  p_name text,
  p_sku text default null,
  p_quote_number text default null,
  p_invoice_number text default null,
  p_sales_order_number text default null,
  p_po_number text default null,
  p_unit text default null,
  p_notes text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_stock_item private.stock_items;
  v_name text := nullif(btrim(p_name), '');
  v_sku text := coalesce(nullif(btrim(p_sku), ''), '');
  v_quote_number text := coalesce(nullif(btrim(p_quote_number), ''), '');
  v_invoice_number text := coalesce(nullif(btrim(p_invoice_number), ''), '');
  v_sales_order_number text := coalesce(nullif(btrim(p_sales_order_number), ''), '');
  v_po_number text := coalesce(nullif(btrim(p_po_number), ''), '');
  v_unit text := coalesce(nullif(btrim(p_unit), ''), 'units');
  v_notes text := coalesce(nullif(btrim(p_notes), ''), '');
begin
  v_actor := private.require_user(p_token, array['admin']);

  select *
  into v_stock_item
  from private.stock_items
  where id = p_stock_item_id;

  if v_stock_item.id is null then
    raise exception 'Stock item not found.';
  end if;

  if v_name is null then
    raise exception 'Stock item name is required.';
  end if;

  if v_quote_number = ''
     and v_invoice_number = ''
     and v_sales_order_number = ''
     and v_po_number = '' then
    raise exception 'Enter at least one quote, sales order, invoice, or PO number.';
  end if;

  if exists (
    select 1
    from private.stock_items
    where id <> p_stock_item_id
      and lower(btrim(name)) = lower(v_name)
      and lower(btrim(quote_number)) = lower(v_quote_number)
      and lower(btrim(invoice_number)) = lower(v_invoice_number)
      and lower(btrim(sales_order_number)) = lower(v_sales_order_number)
      and lower(btrim(po_number)) = lower(v_po_number)
  ) then
    raise exception 'That stock item already exists for the same reference combination.';
  end if;

  if v_sku <> '' and exists (
    select 1
    from private.stock_items
    where id <> p_stock_item_id
      and lower(btrim(sku)) = lower(v_sku)
  ) then
    raise exception 'That stock code is already in use.';
  end if;

  update private.stock_items
  set name = v_name,
      sku = v_sku,
      quote_number = v_quote_number,
      invoice_number = v_invoice_number,
      sales_order_number = v_sales_order_number,
      po_number = v_po_number,
      unit = v_unit,
      notes = v_notes
  where id = p_stock_item_id;

  return jsonb_build_object('ok', true, 'updatedBy', v_actor.id);
end;
$$;

create or replace function public.delete_stock_item(
  p_token uuid,
  p_stock_item_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
begin
  v_actor := private.require_user(p_token, array['admin']);

  if not exists (
    select 1
    from private.stock_items
    where id = p_stock_item_id
  ) then
    raise exception 'Stock item not found.';
  end if;

  delete from private.stock_items
  where id = p_stock_item_id;

  return jsonb_build_object('ok', true, 'deletedBy', v_actor.id);
end;
$$;

drop function if exists public.record_stock_movement(uuid, uuid, text, integer, text, uuid, text);

create or replace function public.record_stock_movement(
  p_token uuid,
  p_stock_item_id uuid,
  p_movement_type text,
  p_quantity integer,
  p_supplier_name text default null,
  p_driver_user_id uuid default null,
  p_notes text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_stock_item private.stock_items;
  v_driver private.app_users;
  v_movement_type text := lower(nullif(btrim(p_movement_type), ''));
  v_quantity integer := coalesce(p_quantity, 0);
  v_supplier_name text := coalesce(nullif(btrim(p_supplier_name), ''), '');
  v_notes text := coalesce(nullif(btrim(p_notes), ''), '');
  v_on_hand integer := 0;
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'logistics') then
    raise exception 'Permission denied';
  end if;

  if v_movement_type not in ('in', 'out') then
    raise exception 'Movement type must be in or out.';
  end if;

  if v_quantity <= 0 then
    raise exception 'Quantity must be greater than zero.';
  end if;

  select *
  into v_stock_item
  from private.stock_items
  where id = p_stock_item_id;

  if v_stock_item.id is null then
    raise exception 'Stock item not found.';
  end if;

  if v_movement_type = 'in' and v_supplier_name = '' then
    raise exception 'Supplier is required for stock coming in.';
  end if;

  if v_movement_type = 'out' then
    select *
    into v_driver
    from private.app_users
    where id = p_driver_user_id
      and role = 'driver'
      and active;

    if v_driver.id is null then
      raise exception 'Driver is required for stock going out.';
    end if;

    v_on_hand := private.stock_on_hand(p_stock_item_id);
    if v_on_hand < v_quantity then
      raise exception 'Not enough stock on hand for that movement.';
    end if;
  end if;

  insert into private.stock_movements (
    stock_item_id,
    movement_type,
    quantity,
    supplier_name,
    driver_user_id,
    notes,
    created_by_user_id
  )
  values (
    p_stock_item_id,
    v_movement_type,
    v_quantity,
    case when v_movement_type = 'in' then v_supplier_name else '' end,
    case when v_movement_type = 'out' then p_driver_user_id else null end,
    v_notes,
    v_actor.id
  );

  return jsonb_build_object('ok', true);
end;
$$;

drop function if exists public.update_stock_movement(uuid, uuid, uuid, text, integer, text, uuid, text);

create or replace function public.update_stock_movement(
  p_token uuid,
  p_stock_movement_id uuid,
  p_stock_item_id uuid,
  p_movement_type text,
  p_quantity integer,
  p_supplier_name text default null,
  p_driver_user_id uuid default null,
  p_notes text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_existing private.stock_movements;
  v_stock_item private.stock_items;
  v_driver private.app_users;
  v_movement_type text := lower(nullif(btrim(p_movement_type), ''));
  v_quantity integer := coalesce(p_quantity, 0);
  v_supplier_name text := coalesce(nullif(btrim(p_supplier_name), ''), '');
  v_notes text := coalesce(nullif(btrim(p_notes), ''), '');
  v_existing_effect integer := 0;
  v_new_effect integer := 0;
  v_old_on_hand integer := 0;
  v_new_on_hand integer := 0;
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'logistics') then
    raise exception 'Permission denied';
  end if;

  if v_movement_type not in ('in', 'out') then
    raise exception 'Movement type must be in or out.';
  end if;

  if v_quantity <= 0 then
    raise exception 'Quantity must be greater than zero.';
  end if;

  select *
  into v_existing
  from private.stock_movements
  where id = p_stock_movement_id;

  if v_existing.id is null then
    raise exception 'Stock movement not found.';
  end if;

  select *
  into v_stock_item
  from private.stock_items
  where id = p_stock_item_id;

  if v_stock_item.id is null then
    raise exception 'Stock item not found.';
  end if;

  if v_movement_type = 'in' and v_supplier_name = '' then
    raise exception 'Supplier is required for stock coming in.';
  end if;

  if v_movement_type = 'out' then
    select *
    into v_driver
    from private.app_users
    where id = p_driver_user_id
      and role = 'driver'
      and active;

    if v_driver.id is null then
      raise exception 'Driver is required for stock going out.';
    end if;
  end if;

  v_existing_effect := case when v_existing.movement_type = 'in' then v_existing.quantity else -v_existing.quantity end;
  v_new_effect := case when v_movement_type = 'in' then v_quantity else -v_quantity end;

  if v_existing.stock_item_id = p_stock_item_id then
    v_old_on_hand := private.stock_on_hand(v_existing.stock_item_id) - v_existing_effect + v_new_effect;
    if v_old_on_hand < 0 then
      raise exception 'Not enough stock on hand for that movement.';
    end if;
  else
    v_old_on_hand := private.stock_on_hand(v_existing.stock_item_id) - v_existing_effect;
    if v_old_on_hand < 0 then
      raise exception 'Editing this movement would leave the original stock item below zero.';
    end if;

    v_new_on_hand := private.stock_on_hand(p_stock_item_id) + v_new_effect;
    if v_new_on_hand < 0 then
      raise exception 'Not enough stock on hand for that movement.';
    end if;
  end if;

  update private.stock_movements
  set stock_item_id = p_stock_item_id,
      movement_type = v_movement_type,
      quantity = v_quantity,
      supplier_name = case when v_movement_type = 'in' then v_supplier_name else '' end,
      driver_user_id = case when v_movement_type = 'out' then p_driver_user_id else null end,
      notes = v_notes
  where id = v_existing.id;

  return jsonb_build_object('ok', true);
end;
$$;

drop function if exists public.create_artwork_request(uuid, uuid, integer, text, text);

create or replace function public.create_artwork_request(
  p_token uuid,
  p_stock_item_id uuid,
  p_requested_quantity integer,
  p_notes text default null,
  p_sent_to text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_stock_item private.stock_items;
  v_requested_quantity integer := coalesce(p_requested_quantity, 0);
  v_notes text := coalesce(nullif(btrim(p_notes), ''), '');
  v_sent_to text := coalesce(nullif(btrim(p_sent_to), ''), '');
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'logistics') then
    raise exception 'Permission denied';
  end if;

  if v_requested_quantity <= 0 then
    raise exception 'Requested quantity must be greater than zero.';
  end if;

  select *
  into v_stock_item
  from private.stock_items
  where id = p_stock_item_id;

  if v_stock_item.id is null then
    raise exception 'Stock item not found.';
  end if;

  insert into private.artwork_requests (
    stock_item_id,
    requested_quantity,
    notes,
    sent_to,
    requested_by_user_id
  )
  values (
    p_stock_item_id,
    v_requested_quantity,
    v_notes,
    v_sent_to,
    v_actor.id
  );

  return jsonb_build_object('ok', true);
end;
$$;

drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, boolean);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, text, boolean);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, boolean, text);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, boolean, text, boolean);

drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, boolean, text, boolean);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, text, text, boolean, text, boolean);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, text, text, boolean, text, boolean, uuid);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, text, text, text, text, boolean, text, boolean, uuid);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, text, text, text, text, text, boolean, text, boolean, uuid);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, text, text, text, text, text, text, boolean, text, boolean, uuid);
drop function if exists public.create_order(uuid, uuid, uuid, text, text, text, text, text, text, text, text, text, boolean, text, boolean, uuid, text[]);

create or replace function public.create_order(
  p_token uuid,
  p_driver_user_id uuid,
  p_location_id uuid,
  p_entry_type text,
  p_quote_number text,
  p_sales_order_number text default null,
  p_invoice_number text default null,
  p_po_number text default null,
  p_branding text default null,
  p_stock_description text default null,
  p_stock_item_names text[] default null,
  p_delivery_address text default null,
  p_priority text default 'medium',
  p_allow_duplicate boolean default false,
  p_notice text default null,
  p_move_to_factory boolean default false,
  p_factory_destination_location_id uuid default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_driver private.app_users;
  v_location private.locations;
  v_factory_destination private.locations;
  v_stock_item_id uuid;
  v_first_stock_item_id uuid;
  v_stock_item_created boolean := false;
  v_stock_item_created_count integer := 0;
  v_stock_item_count integer := 0;
  v_today date := private.today_local();
  v_entry_type text := lower(nullif(btrim(p_entry_type), ''));
  v_quote_number text := nullif(btrim(p_quote_number), '');
  v_sales_order_number text := coalesce(nullif(btrim(p_sales_order_number), ''), '');
  v_invoice_number text := coalesce(nullif(btrim(p_invoice_number), ''), '');
  v_po_number text := coalesce(nullif(btrim(p_po_number), ''), '');
  v_branding text := coalesce(nullif(btrim(p_branding), ''), '');
  v_stock_description text := nullif(btrim(p_stock_description), '');
  v_delivery_address text := coalesce(nullif(btrim(p_delivery_address), ''), '');
  v_priority text := lower(coalesce(nullif(btrim(p_priority), ''), 'medium'));
  v_notice text := coalesce(nullif(btrim(p_notice), ''), '');
  v_stock_item_notes text := case
    when coalesce(nullif(btrim(p_branding), ''), '') <> '' then concat('Branding: ', btrim(p_branding), '.')
    else ''
  end;
  v_move_to_factory boolean := coalesce(p_move_to_factory, false);
  v_factory_destination_location_id uuid := case when coalesce(p_move_to_factory, false) then p_factory_destination_location_id else null end;
  v_stock_item_names text[] := '{}'::text[];
  v_processed_stock_item_names text[] := '{}'::text[];
  v_stock_item_name text;
  v_stock_item_name_key text;
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'sales') then
    raise exception 'Permission denied';
  end if;

  if v_entry_type not in ('collection', 'delivery') then
    raise exception 'Entry type must be collection or delivery.';
  end if;

  if v_quote_number is null then
    raise exception 'Quote number is required.';
  end if;

  if v_stock_description is null then
    raise exception 'Stock description is required.';
  end if;

  if v_entry_type = 'delivery' and v_delivery_address = '' then
    raise exception 'Delivery address is required for delivery entries.';
  elsif v_entry_type <> 'delivery' then
    v_delivery_address := '';
  end if;

  if v_priority not in ('high', 'medium', 'low') then
    raise exception 'Choose a valid priority.';
  end if;

  if p_driver_user_id is not null then
    select *
    into v_driver
    from private.app_users
    where id = p_driver_user_id
      and role = 'driver'
      and active;

    if v_driver.id is null then
      raise exception 'Driver not found.';
    end if;
  end if;

  select *
  into v_location
  from private.locations
  where id = p_location_id;

  if v_location.id is null then
    raise exception 'Location not found.';
  end if;

  if v_move_to_factory and v_entry_type <> 'collection' then
    raise exception 'Only collection entries can be marked to move stock to a factory.';
  end if;

  if v_move_to_factory and v_factory_destination_location_id is null then
    raise exception 'Select which factory the collected stock should go to.';
  end if;

  if v_move_to_factory then
    select *
    into v_factory_destination
    from private.locations
    where id = v_factory_destination_location_id
      and location_type in ('factory', 'both');

    if v_factory_destination.id is null then
      raise exception 'Factory destination not found.';
    end if;
  else
    v_factory_destination_location_id := null;
  end if;

  v_stock_item_names := coalesce(p_stock_item_names, array[v_stock_description]);

  if coalesce(array_length(v_stock_item_names, 1), 0) = 0 then
    v_stock_item_names := array[v_stock_description];
  end if;

  if p_driver_user_id is not null
     and exists (
       select 1
       from private.orders
       where driver_user_id = p_driver_user_id
         and inhouse_order_number = v_quote_number
         and status = 'active'
     )
     and not (v_actor.role = 'admin' and coalesce(p_allow_duplicate, false)) then
    raise exception 'Duplicate blocked. This driver already has an active entry for that quote number.';
  end if;

  if p_driver_user_id is not null
     and exists (
       select 1
       from private.orders
       where driver_user_id = p_driver_user_id
         and location_id = p_location_id
         and status = 'completed'
         and scheduled_for = v_today
     )
     and not (v_actor.role = 'admin' and coalesce(p_allow_duplicate, false)) then
    raise exception 'Completed stop blocked. This driver has already completed that pickup location today. Admin authorization is required to send them back.';
  end if;

  insert into private.orders (
    customer_name,
    driver_user_id,
    location_id,
    entry_type,
    factory_order_number,
    inhouse_order_number,
    invoice_number,
    po_number,
    branding,
    stock_description,
    delivery_address,
    priority,
    notes,
    move_to_factory,
    factory_destination_location_id,
    scheduled_for,
    original_scheduled_for,
    created_by_user_id
  )
  values (
    v_quote_number,
    p_driver_user_id,
    p_location_id,
    v_entry_type,
    v_sales_order_number,
    v_quote_number,
    v_invoice_number,
    v_po_number,
    v_branding,
    v_stock_description,
    v_delivery_address,
    v_priority,
    v_notice,
    v_move_to_factory,
    v_factory_destination_location_id,
    v_today,
    v_today,
    v_actor.id
  );

  foreach v_stock_item_name in array v_stock_item_names loop
    v_stock_item_name := nullif(btrim(v_stock_item_name), '');
    continue when v_stock_item_name is null;

    v_stock_item_name_key := lower(v_stock_item_name);
    if array_position(v_processed_stock_item_names, v_stock_item_name_key) is not null then
      continue;
    end if;

    v_processed_stock_item_names := array_append(v_processed_stock_item_names, v_stock_item_name_key);
    v_stock_item_count := v_stock_item_count + 1;

    select s.id
    into v_stock_item_id
    from private.stock_items s
    where lower(btrim(s.name)) = v_stock_item_name_key
      and lower(btrim(s.quote_number)) = lower(v_quote_number)
      and lower(btrim(s.invoice_number)) = lower(v_invoice_number)
      and lower(btrim(s.sales_order_number)) = lower(v_sales_order_number)
      and lower(btrim(s.po_number)) = lower(v_po_number)
    limit 1;

    if v_stock_item_id is null then
      begin
        insert into private.stock_items (
          name,
          sku,
          quote_number,
          invoice_number,
          sales_order_number,
          po_number,
          unit,
          notes,
          created_source,
          created_by_user_id
        )
        values (
          v_stock_item_name,
          '',
          v_quote_number,
          v_invoice_number,
          v_sales_order_number,
          v_po_number,
          'units',
          v_stock_item_notes,
          'order',
          v_actor.id
        )
        returning id into v_stock_item_id;

        v_stock_item_created_count := v_stock_item_created_count + 1;
      exception
        when unique_violation then
          select s.id
          into v_stock_item_id
          from private.stock_items s
          where lower(btrim(s.name)) = v_stock_item_name_key
            and lower(btrim(s.quote_number)) = lower(v_quote_number)
            and lower(btrim(s.invoice_number)) = lower(v_invoice_number)
            and lower(btrim(s.sales_order_number)) = lower(v_sales_order_number)
            and lower(btrim(s.po_number)) = lower(v_po_number)
          limit 1;
      end;
    end if;

    if v_first_stock_item_id is null then
      v_first_stock_item_id := v_stock_item_id;
    end if;
  end loop;

  v_stock_item_created := v_stock_item_created_count > 0;

  return jsonb_build_object(
    'ok', true,
    'stockItemId', v_first_stock_item_id,
    'stockItemCreated', v_stock_item_created,
    'stockItemCount', v_stock_item_count
  );
end;
$$;

drop function if exists public.assign_order(uuid, uuid, uuid);
drop function if exists public.assign_order(uuid, uuid, uuid, boolean);

create or replace function public.assign_order(
  p_token uuid,
  p_order_id uuid,
  p_driver_user_id uuid default null,
  p_allow_duplicate boolean default false
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_order private.orders;
  v_driver private.app_users;
  v_today date := private.today_local();
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'sales', 'driver') then
    raise exception 'Permission denied';
  end if;

  select *
  into v_order
  from private.orders
  where id = p_order_id;

  if v_order.id is null then
    raise exception 'Order not found.';
  end if;

  if v_order.status <> 'active' then
    raise exception 'Only active entries can be reassigned.';
  end if;

  if v_actor.role = 'driver' then
    if v_order.driver_user_id is distinct from v_actor.id then
      raise exception 'Drivers can only transfer their own assigned entries.';
    end if;

    if p_driver_user_id is null then
      raise exception 'Drivers must transfer the entry to another driver.';
    end if;

    if p_driver_user_id = v_actor.id then
      raise exception 'Choose another driver for the transfer.';
    end if;
  end if;

  if p_driver_user_id is not null then
    select *
    into v_driver
    from private.app_users
    where id = p_driver_user_id
      and role = 'driver'
      and active;

    if v_driver.id is null then
      raise exception 'Driver not found.';
    end if;

    if exists (
      select 1
      from private.orders
      where id <> p_order_id
        and driver_user_id = p_driver_user_id
        and inhouse_order_number = v_order.inhouse_order_number
        and status = 'active'
    ) and not (v_actor.role = 'admin' and coalesce(p_allow_duplicate, false)) then
      raise exception 'Duplicate blocked. This driver already has an active entry for that quote number.';
    end if;

    if exists (
      select 1
      from private.orders
      where id <> p_order_id
        and driver_user_id = p_driver_user_id
        and location_id = v_order.location_id
        and status = 'completed'
        and scheduled_for = v_today
    ) and not (v_actor.role = 'admin' and coalesce(p_allow_duplicate, false)) then
      raise exception 'Completed stop blocked. This driver has already completed that pickup location today. Admin authorization is required to send them back.';
    end if;
  end if;

  update private.orders
  set driver_user_id = p_driver_user_id,
      scheduled_for = case
        when v_order.driver_user_id is null or p_driver_user_id is null then v_today
        else scheduled_for
      end,
      original_scheduled_for = case
        when v_order.driver_user_id is null or p_driver_user_id is null then v_today
        else original_scheduled_for
      end,
      priority = case
        when v_order.driver_user_id is not null
          and p_driver_user_id is null
          and v_order.carry_over_count > 0
          and priority = 'high'
        then 'medium'
        else priority
      end,
      carry_over_count = case
        when v_order.driver_user_id is null or p_driver_user_id is null then 0
        else carry_over_count
      end,
      updated_at = now()
  where id = p_order_id;

  return jsonb_build_object('ok', true);
end;
$$;

create or replace function public.set_order_priority(
  p_token uuid,
  p_order_id uuid,
  p_priority text default 'medium'
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_order private.orders;
  v_priority text := lower(coalesce(nullif(btrim(p_priority), ''), 'medium'));
begin
  v_actor := private.require_user(p_token);

  if v_actor.role <> 'admin' then
    raise exception 'Permission denied';
  end if;

  if v_priority not in ('high', 'medium', 'low') then
    raise exception 'Choose a valid priority.';
  end if;

  select *
  into v_order
  from private.orders
  where id = p_order_id;

  if v_order.id is null then
    raise exception 'Order not found.';
  end if;

  if v_order.status <> 'active' then
    raise exception 'Only active entries can change priority.';
  end if;

  update private.orders
  set priority = v_priority,
      updated_at = now()
  where id = p_order_id;

  return jsonb_build_object('ok', true, 'priority', v_priority);
end;
$$;

create or replace function public.clear_all_order_priorities(p_token uuid)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_updated_orders integer := 0;
begin
  v_actor := private.require_user(p_token);

  if v_actor.role <> 'admin' then
    raise exception 'Permission denied';
  end if;

  update private.orders
  set priority = 'medium',
      updated_at = now()
  where status = 'active'
    and priority = 'high';

  get diagnostics v_updated_orders = row_count;

  return jsonb_build_object(
    'ok', true,
    'updatedOrders', v_updated_orders
  );
end;
$$;

create or replace function public.clear_order_rollovers(p_token uuid)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_updated_orders integer := 0;
begin
  v_actor := private.require_user(p_token);

  if v_actor.role <> 'admin' then
    raise exception 'Permission denied';
  end if;

  update private.orders
  set carry_over_count = 0,
      original_scheduled_for = scheduled_for,
      updated_at = now()
  where status = 'active'
    and carry_over_count > 0;

  get diagnostics v_updated_orders = row_count;

  return jsonb_build_object(
    'ok', true,
    'updatedOrders', v_updated_orders
  );
end;
$$;

create or replace function public.set_order_flag(
  p_token uuid,
  p_order_id uuid,
  p_flag_type text default null,
  p_note text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_order private.orders;
  v_flag_type text := lower(nullif(btrim(p_flag_type), ''));
  v_note text := coalesce(nullif(btrim(p_note), ''), '');
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'driver') then
    raise exception 'Permission denied';
  end if;

  select *
  into v_order
  from private.orders
  where id = p_order_id;

  if v_order.id is null then
    raise exception 'Order not found.';
  end if;

  if v_order.status <> 'active' then
    raise exception 'Only active entries can be flagged.';
  end if;

  if v_actor.role = 'driver' and v_order.driver_user_id is distinct from v_actor.id then
    raise exception 'Drivers can only flag their own assigned orders.';
  end if;

  if v_flag_type is not null and v_flag_type not in ('not_collected', 'not_ready') then
    raise exception 'Choose a valid follow-up reason.';
  end if;

  update private.orders
  set driver_flag_type = v_flag_type,
      driver_flag_note = case when v_flag_type is null then '' else v_note end,
      driver_flagged_at = case when v_flag_type is null then null else now() end,
      driver_flagged_by_user_id = case when v_flag_type is null then null else v_actor.id end,
      updated_at = now()
  where id = p_order_id;

  return jsonb_build_object('ok', true, 'flagType', coalesce(v_flag_type, ''));
end;
$$;

drop function if exists public.pick_up_order(uuid, uuid);

create or replace function public.pick_up_order(
  p_token uuid,
  p_order_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_order private.orders;
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'driver') then
    raise exception 'Permission denied';
  end if;

  select *
  into v_order
  from private.orders
  where id = p_order_id;

  if v_order.id is null then
    raise exception 'Order not found.';
  end if;

  if v_actor.role = 'driver' and v_order.driver_user_id is distinct from v_actor.id then
    raise exception 'Drivers can only pick up their own assigned orders.';
  end if;

  if v_order.status <> 'active' then
    raise exception 'Only active entries can be marked as picked up.';
  end if;

  if v_order.picked_up_at is not null then
    return jsonb_build_object('ok', true);
  end if;

  update private.orders
  set picked_up_at = now(),
      picked_up_by_user_id = v_actor.id,
      updated_at = now()
  where id = p_order_id
    and status = 'active'
    and picked_up_at is null;

  return jsonb_build_object('ok', true, 'pickedUpBy', v_actor.id);
end;
$$;

drop function if exists public.complete_order(uuid, uuid);
drop function if exists public.complete_order(uuid, uuid, text);

create or replace function public.complete_order(
  p_token uuid,
  p_order_id uuid,
  p_completion_type text default null
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
  v_order private.orders;
  v_completion_type text := lower(coalesce(nullif(btrim(p_completion_type), ''), 'office'));
begin
  v_actor := private.require_user(p_token);

  if v_actor.role not in ('admin', 'driver') then
    raise exception 'Permission denied';
  end if;

  select *
  into v_order
  from private.orders
  where id = p_order_id;

  if v_order.id is null then
    raise exception 'Order not found.';
  end if;

  if v_actor.role = 'driver' and v_order.driver_user_id is distinct from v_actor.id then
    raise exception 'Drivers can only complete their own assigned orders.';
  end if;

  if v_order.status = 'completed' then
    return jsonb_build_object('ok', true);
  end if;

  if v_completion_type not in ('office', 'factory') then
    raise exception 'Choose a valid completion action.';
  end if;

  if v_completion_type = 'factory' and not v_order.move_to_factory then
    raise exception 'Only factory-transfer entries can be marked as dropped at the factory.';
  end if;

  if v_order.picked_up_at is null then
    raise exception 'Mark the entry as picked up before dropping it off.';
  end if;

  update private.orders
  set status = 'completed',
      completed_at = now(),
      completion_type = v_completion_type,
      completed_by_user_id = v_actor.id,
      driver_flag_type = null,
      driver_flag_note = '',
      driver_flagged_at = null,
      driver_flagged_by_user_id = null,
      updated_at = now()
  where id = p_order_id
    and status <> 'completed';

  return jsonb_build_object('ok', true, 'completionType', v_completion_type);
end;
$$;

create or replace function public.delete_order(
  p_token uuid,
  p_order_id uuid
)
returns jsonb
language plpgsql
security definer
set search_path = ''
as $$
declare
  v_actor private.app_users;
begin
  v_actor := private.require_user(p_token, array['admin']);

  delete from private.orders
  where id = p_order_id;

  return jsonb_build_object('ok', true, 'deletedBy', v_actor.id);
end;
$$;

commit;
