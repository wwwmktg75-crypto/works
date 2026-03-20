-- スキルマッチングMVP向けの最小テーブル定義
create table if not exists public.members (
  id uuid primary key default gen_random_uuid(),
  displayName text not null,
  displayTitle text not null default '',
  skillTags text not null,
  hourlyRate numeric not null default 0,
  profileText text not null default '',
  portfolioUrl text not null default '',
  email text not null,
  isApproved boolean not null default false,
  createdAt timestamptz not null default now()
);

create table if not exists public.inquiries (
  id uuid primary key default gen_random_uuid(),
  memberRecordId uuid not null references public.members(id) on delete cascade,
  clientName text not null,
  clientEmail text not null,
  message text not null,
  status text not null default 'new',
  createdAt timestamptz not null default now()
);

create table if not exists public.orders (
  id uuid primary key default gen_random_uuid(),
  checkoutSessionId text not null unique,
  paymentIntentId text not null default '',
  memberRecordId uuid references public.members(id) on delete set null,
  serviceTitle text not null default '',
  amount numeric not null default 0,
  currency text not null default 'jpy',
  payerEmail text not null default '',
  status text not null default 'paid',
  createdAt timestamptz not null default now()
);

create index if not exists idxMembersIsApproved on public.members(isApproved);
create index if not exists idxInquiriesMemberRecordId on public.inquiries(memberRecordId);
create index if not exists idxOrdersMemberRecordId on public.orders(memberRecordId);
