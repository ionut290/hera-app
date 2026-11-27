-- PostgreSQL 16+, apply with a migration owner, NOT the runtime role.
-- Dedicated database only; create LOGIN role varga_app separately, no BYPASSRLS.
BEGIN;
CREATE TABLE tenants (
 id uuid PRIMARY KEY, owner_uid text NOT NULL, request_id uuid NOT NULL,
 name text NOT NULL CHECK (char_length(name) BETWEEN 2 AND 120),
 join_code text NOT NULL UNIQUE CHECK(join_code ~ '^VC-[A-Z0-9]{8}$'),
 status text NOT NULL DEFAULT 'active' CHECK(status IN ('active','suspended')),
 plan text NOT NULL DEFAULT 'trial' CHECK(plan IN ('trial','full')),
 max_jobs integer NOT NULL DEFAULT 1 CHECK(max_jobs >= 1),
 max_plants integer NOT NULL DEFAULT 10 CHECK(max_plants >= 1),
 max_admins integer NOT NULL DEFAULT 1 CHECK(max_admins >= 1),
 max_operators integer NOT NULL DEFAULT 3 CHECK(max_operators >= 0),
 created_at timestamptz NOT NULL DEFAULT now(),
 UNIQUE(owner_uid,request_id)
);
CREATE TABLE members (
 tenant_id uuid NOT NULL REFERENCES tenants(id), uid text NOT NULL,
 role text NOT NULL CHECK(role IN ('admin','employee')),
 status text NOT NULL DEFAULT 'active' CHECK(status IN ('active','disabled')),
 PRIMARY KEY(tenant_id,uid)
);
CREATE INDEX members_uid_idx ON members(uid,tenant_id);
CREATE TABLE join_requests (
 tenant_id uuid NOT NULL REFERENCES tenants(id), uid text NOT NULL,
 email text NOT NULL, display_name text NOT NULL,
 status text NOT NULL DEFAULT 'pending' CHECK(status IN ('pending','accepted','rejected')),
 requested_at timestamptz NOT NULL DEFAULT now(), reviewed_at timestamptz,
 reviewed_by text, PRIMARY KEY(tenant_id,uid)
);
CREATE TABLE jobs (
 tenant_id uuid NOT NULL REFERENCES tenants(id), id uuid NOT NULL,
 name text NOT NULL CHECK(char_length(name) BETWEEN 1 AND 160),
 created_by text NOT NULL, created_at timestamptz NOT NULL DEFAULT now(),
 PRIMARY KEY(tenant_id,id)
);
CREATE TABLE plants (
 tenant_id uuid NOT NULL, id uuid NOT NULL, job_id uuid NOT NULL,
 name text NOT NULL CHECK(char_length(name) BETWEEN 1 AND 160),
 lat double precision CHECK(lat BETWEEN -90 AND 90), lng double precision CHECK(lng BETWEEN -180 AND 180),
 created_by text NOT NULL, created_at timestamptz NOT NULL DEFAULT now(),
 CHECK((lat IS NULL) = (lng IS NULL)), PRIMARY KEY(tenant_id,id),
 FOREIGN KEY(tenant_id,job_id) REFERENCES jobs(tenant_id,id)
);
CREATE INDEX plants_job_idx ON plants(tenant_id,job_id);
-- Defense in depth: operational rows always require an explicit transaction tenant.
ALTER TABLE jobs ENABLE ROW LEVEL SECURITY;
ALTER TABLE jobs FORCE ROW LEVEL SECURITY;
CREATE POLICY jobs_scope ON jobs USING(tenant_id::text = current_setting('varga.tenant_id',true)) WITH CHECK(tenant_id::text = current_setting('varga.tenant_id',true));
ALTER TABLE plants ENABLE ROW LEVEL SECURITY;
ALTER TABLE plants FORCE ROW LEVEL SECURITY;
CREATE POLICY plants_scope ON plants USING(tenant_id::text = current_setting('varga.tenant_id',true)) WITH CHECK(tenant_id::text = current_setting('varga.tenant_id',true));
REVOKE ALL ON tenants,members,join_requests,jobs,plants FROM PUBLIC;
GRANT USAGE ON SCHEMA public TO varga_app;
GRANT SELECT,INSERT ON tenants,members,join_requests,jobs,plants TO varga_app;
GRANT UPDATE(status,reviewed_at,reviewed_by) ON join_requests TO varga_app;
-- FOR UPDATE locks require UPDATE privilege; narrowly scoped column, no API mutation.
GRANT UPDATE(status) ON tenants TO varga_app;
COMMIT;
