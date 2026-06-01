-- ------------------------------------------------------------
-- FoodForward Mother and Child Programme Supabase Schema
-- ------------------------------------------------------------
-- Run this file in the Supabase SQL editor before running:
-- python load_to_supabase.py
--
-- The tables mirror the cleaned and anonymised Project 1 workbook.
-- Direct personal identifiers are not included.
-- ------------------------------------------------------------

drop table if exists ff_data_quality_log cascade;
drop table if exists ff_measurements cascade;
drop table if exists ff_participants cascade;
drop table if exists ff_food_distributed cascade;
drop table if exists ff_bo_profile cascade;
drop table if exists ff_sites cascade;

create table ff_sites (
    id bigint generated always as identity primary key,
    site_id text,
    bo_code text,
    bo_name text,
    notes text
);

create table ff_bo_profile (
    id bigint generated always as identity primary key,
    site_id text,
    bo_code text,
    province text,
    suburb text,
    bo_size text,
    bo_category text,
    total_beneficiaries numeric,
    beneficiaries_0_5 numeric,
    beneficiaries_6_18 numeric,
    beneficiaries_19_29 numeric,
    beneficiaries_30_59 numeric,
    beneficiaries_60_plus numeric,
    feeding_days text,
    serve_breakfast numeric,
    serve_lunch numeric,
    serve_supper numeric,
    feeding_scheme text,
    receive_government_grant text,
    annual_food_spending numeric,
    meal_sitdown text,
    covered_eating_area text,
    adequate_food_storage text,
    have_cooking_equipment text,
    internet_access text,
    daily_menu text
);

create table ff_food_distributed (
    id bigint generated always as identity primary key,
    distribution_id text,
    posting_date date,
    posting_month text,
    site_id text,
    bo_code text,
    line_location_code text,
    item_code text,
    description text,
    quantity numeric,
    gross_weight numeric,
    line_weight numeric,
    stock_type text,
    nutrition text,
    category text
);

create table ff_participants (
    id bigint generated always as identity primary key,
    participant_id text,
    site_id text,
    partner text,
    birth_month text,
    birth_year numeric,
    pii_removed text
);

create table ff_measurements (
    id bigint generated always as identity primary key,
    participant_id text,
    site_id text,
    partner text,
    measurement_month text,
    birth_month text,
    age_months numeric,
    weight_kg numeric,
    height_cm numeric,
    data_quality_flag text
);

create table ff_data_quality_log (
    id bigint generated always as identity primary key,
    source_sheet text,
    source_row numeric,
    participant_id text,
    measurement_month text,
    issue text,
    original_cell_value text
);

create index idx_ff_food_site on ff_food_distributed(site_id);
create index idx_ff_food_month on ff_food_distributed(posting_month);
create index idx_ff_measurements_site on ff_measurements(site_id);
create index idx_ff_measurements_participant on ff_measurements(participant_id);
create index idx_ff_measurements_month on ff_measurements(measurement_month);
create index idx_ff_participants_site on ff_participants(site_id);

-- Optional: enable row-level security after deciding who should access the data.
-- alter table ff_sites enable row level security;
-- alter table ff_bo_profile enable row level security;
-- alter table ff_food_distributed enable row level security;
-- alter table ff_participants enable row level security;
-- alter table ff_measurements enable row level security;
-- alter table ff_data_quality_log enable row level security;
