create extension if not exists pgcrypto;

create table if not exists public.detenciones (
    id uuid primary key default gen_random_uuid(),
    identificador text not null unique,
    equipo text not null,
    fecha_detencion date not null,
    hora_inicio time,
    fecha_hora_inicio timestamp,
    turno text,
    codigo text,
    razon text,
    comentario text,
    categoria text,
    tipo_categoria text,
    descripcion_normalizada text,
    requiere_ot boolean not null default true,
    motivo_exclusion text,
    archivo_origen text,
    estado text not null default 'PENDIENTE',
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create table if not exists public.ordenes_trabajo (
    id uuid primary key default gen_random_uuid(),
    numero_ot text not null unique,
    equipo text,
    turno text,
    descripcion text,
    descripcion_normalizada text,
    archivo_origen text,
    fecha_recepcion date not null,
    fecha_hora_recepcion timestamp not null,
    estado_validacion text,
    campos_faltantes integer not null default 0,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create table if not exists public.detencion_ot (
    id uuid primary key default gen_random_uuid(),
    detencion_id uuid not null references public.detenciones(id) on delete cascade,
    ot_id uuid not null references public.ordenes_trabajo(id) on delete cascade,
    confianza numeric(5,4),
    tipo_asociacion text not null default 'MANUAL',
    confirmada boolean not null default true,
    created_at timestamptz not null default now(),
    unique(detencion_id, ot_id)
);

create index if not exists idx_detenciones_fecha on public.detenciones(fecha_detencion);
create index if not exists idx_detenciones_equipo on public.detenciones(equipo);
create index if not exists idx_ot_equipo on public.ordenes_trabajo(equipo);
create index if not exists idx_asociacion_detencion on public.detencion_ot(detencion_id);

alter table public.detenciones enable row level security;
alter table public.ordenes_trabajo enable row level security;
alter table public.detencion_ot enable row level security;

-- La aplicación usa la service_role key en Streamlit Secrets.
-- No crees políticas públicas si los datos son corporativos.
