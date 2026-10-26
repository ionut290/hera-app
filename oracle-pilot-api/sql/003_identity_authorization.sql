BEGIN;

CREATE OR REPLACE FUNCTION public.varga_authorize_identity(p_uid text, p_email text)
RETURNS TABLE(matched boolean, allowed boolean, role text)
LANGUAGE sql
STABLE
SECURITY DEFINER
SET search_path = pg_catalog
AS $$
  WITH candidate AS (
    SELECT
      lower(COALESCE(f.data#>>'{v,banned,v}', 'false')) = 'true' AS banned,
      lower(COALESCE(f.data#>>'{v,accountStatus,v}', f.data#>>'{v,statoAccount,v}', '')) AS account_status,
      lower(COALESCE(f.data#>>'{v,approved,v}', f.data#>>'{v,abilitato,v}', f.data#>>'{v,enabled,v}', '')) AS approval_status,
      lower(COALESCE(f.data#>>'{v,isAdmin,v}', f.data#>>'{v,admin,v}', 'false')) = 'true' AS is_admin,
      lower(COALESCE(f.data#>>'{v,role,v}', f.data#>>'{v,ruolo,v}', '')) AS source_role
    FROM public.firestore_import_test_20261026 AS f
    WHERE f.collection_name = 'platformUsers'
      AND f.path ~ '^platformUsers/[^/]+$'
      AND COALESCE(f.data#>>'{v,uid,v}', f.document_id) = p_uid
      AND lower(COALESCE(f.data#>>'{v,email,v}', f.data#>>'{v,loginEmail,v}', '')) = lower(trim(p_email))
    LIMIT 1
  )
  SELECT
    COALESCE((SELECT banned OR account_status <> '' OR approval_status <> '' FROM candidate), false) AS matched,
    COALESCE((SELECT NOT banned AND (approval_status = 'true' OR account_status IN ('attivo', 'active', 'approved')) FROM candidate), false) AS allowed,
    COALESCE((SELECT CASE WHEN is_admin OR source_role IN ('admin', 'super_admin') THEN 'admin' ELSE 'reader' END FROM candidate), 'reader') AS role;
$$;

REVOKE ALL ON FUNCTION public.varga_authorize_identity(text, text) FROM PUBLIC;
GRANT EXECUTE ON FUNCTION public.varga_authorize_identity(text, text) TO varga_reader;

COMMIT;
