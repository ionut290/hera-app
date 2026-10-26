BEGIN;

-- API-facing views expose only an explicit allow-list. The source JSON is never returned.
-- The path filter is essential: collection_name=commesse also labels imported subcollections.
CREATE OR REPLACE VIEW public.varga_catalogo_api_v1 AS
SELECT
  collection_name,
  document_id,
  COALESCE(data#>>'{v,denominazione,v}', data#>>'{v,nome,v}', data#>>'{v,name,v}') AS denominazione,
  data#>>'{v,codice,v}' AS codice,
  data#>>'{v,comune,v}' AS comune,
  data#>>'{v,indirizzo,v}' AS indirizzo,
  data#>>'{v,distretto,v}' AS distretto,
  data#>>'{v,tipologiaIntervento,v}' AS tipologia_intervento,
  COALESCE(data#>>'{v,stato,v}', data#>>'{v,attiva,v}') AS stato,
  CASE
    WHEN lower(COALESCE(data#>>'{v,deleted,v}', 'false')) = 'true' THEN true
    ELSE false
  END AS eliminato
FROM public.firestore_import_test_20261026
WHERE
  (collection_name = 'commesse' AND path ~ '^commesse/[^/]+$')
  OR
  (collection_name = 'impianti' AND path ~ '^impianti/[^/]+$');

CREATE OR REPLACE VIEW public.varga_mappa_api_v1 AS
SELECT collection_name, document_id, latitudine, longitudine
FROM public.varga_mappa_pilota
WHERE collection_name IN ('commesse', 'impianti')
  AND latitudine BETWEEN 35 AND 48
  AND longitudine BETWEEN 6 AND 19;

REVOKE ALL ON public.varga_catalogo_api_v1 FROM PUBLIC;
REVOKE ALL ON public.varga_mappa_api_v1 FROM PUBLIC;
GRANT SELECT ON public.varga_catalogo_api_v1 TO varga_reader;
GRANT SELECT ON public.varga_mappa_api_v1 TO varga_reader;

COMMIT;
