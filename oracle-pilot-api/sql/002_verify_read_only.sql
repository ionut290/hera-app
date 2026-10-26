\set ON_ERROR_STOP on

SELECT current_database() AS database, current_user AS connected_as;

SELECT
  has_table_privilege('varga_reader', 'public.firestore_import_test_20261026', 'SELECT') AS archive_select_must_be_false,
  has_table_privilege('varga_reader', 'public.varga_catalogo_api_v1', 'SELECT') AS catalog_select_must_be_true,
  has_table_privilege('varga_reader', 'public.varga_mappa_api_v1', 'SELECT') AS map_select_must_be_true;

SELECT collection_name, COUNT(*) AS rows
FROM public.varga_catalogo_api_v1
GROUP BY collection_name
ORDER BY collection_name;

SELECT collection_name, COUNT(*) AS rows
FROM public.varga_mappa_api_v1
GROUP BY collection_name
ORDER BY collection_name;

SELECT COUNT(*) AS invalid_coordinates_must_be_zero
FROM public.varga_mappa_api_v1
WHERE latitudine NOT BETWEEN 35 AND 48 OR longitudine NOT BETWEEN 6 AND 19;

SELECT COUNT(*) AS missing_ids_must_be_zero
FROM public.varga_catalogo_api_v1
WHERE document_id IS NULL OR document_id = '';

SELECT
  COUNT(*) FILTER (WHERE collection_name = 'commesse' AND COALESCE(denominazione, '') <> '') AS named_commesse,
  COUNT(*) FILTER (WHERE collection_name = 'impianti' AND COALESCE(denominazione, '') <> '') AS named_impianti,
  COUNT(*) FILTER (WHERE collection_name = 'impianti' AND COALESCE(id_sap, '') <> '') AS impianti_with_id_sap
FROM public.varga_catalogo_api_v1;
