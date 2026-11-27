-- Apply after 001_initial.sql, using the migration owner.
BEGIN;
CREATE TABLE plan_audit (
 id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
 tenant_id uuid NOT NULL REFERENCES tenants(id),
 actor_uid text NOT NULL,
 old_plan jsonb NOT NULL,
 new_plan jsonb NOT NULL,
 changed_at timestamptz NOT NULL DEFAULT now()
);
REVOKE ALL ON plan_audit FROM PUBLIC;
GRANT INSERT ON plan_audit TO varga_app;
GRANT UPDATE(plan,max_jobs,max_plants,max_operators) ON tenants TO varga_app;
COMMIT;
