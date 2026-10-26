BEGIN;

-- API-facing views expose only an explicit allow-list. They never expose the imported JSON archive.
CREATE OR REPLACE VIEW public.varga_catalogo_api_v1 AS
WITH source AS (
  SELECT to_jsonb(v) AS row_data
  FROM public.varga_catalogo_pilota AS v
)
SELECT
  row_data->>'collection_name' AS collection_name,
  row_data->>'document_id' AS document_id,
  COALESCE(row_data->>'denominazione', row_data#>>'{data,v,denominazione,v}', row_data#>>'{data,v,nome,v}') AS denominazione,
  COALESCE(row_data->>'comune', row_data#>>'{data,v,comune,v}') AS comune,
  COALESCE(row_data->>'indirizzo', row_data#>>'{data,v,indirizzo,v}') AS indirizzo,
  COALESCE(row_data->>'stato', row_data#>>'{data,v,stato,v}') AS stato,
  CASE
    WHEN lower(COALESCE(row_data->>'eliminato', row_data#>>'{data,v,deleted,v}', 'false')) = 'true' THEN true
    ELSE false
  END AS eliminato
FROM source
WHERE row_data->>'collection_name' IN ('commesse', 'impianti')
  AND row_data->>'document_id' IS NOT NULL;

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
