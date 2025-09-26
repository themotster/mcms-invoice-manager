BEGIN TRANSACTION;

DROP INDEX IF EXISTS idx_document_definitions_business_sort;
DROP TABLE IF EXISTS document_definitions;
DROP TABLE IF EXISTS documents;
DROP TABLE IF EXISTS jobsheet_template_overrides;
DROP TABLE IF EXISTS document_definition_bindings;

COMMIT;
