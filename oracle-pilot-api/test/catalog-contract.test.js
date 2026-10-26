import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

const viewSql = fs.readFileSync(new URL("../sql/001_api_views.sql", import.meta.url), "utf8");
const repositorySource = fs.readFileSync(new URL("../src/repository.js", import.meta.url), "utf8");
const identitySql = fs.readFileSync(new URL("../sql/003_identity_authorization.sql", import.meta.url), "utf8");

test("catalog exposes operational fields but excludes sensitive source fields", () => {
  for (const field of ["id_sap", "voce_riferimento", "frequenza_annua_minima", "consistenza_verde"]) {
    assert.match(viewSql, new RegExp(`\\b${field}\\b`));
    assert.match(repositorySource, new RegExp(`\\b${field}\\b`));
  }

  for (const forbidden of ["emailOperatore", "createdBy", "creatoDa", "sheetSpreadsheetId", "updatedBy"]) {
    assert.doesNotMatch(viewSql, new RegExp(`\\{v,${forbidden},v\\}`));
    assert.doesNotMatch(repositorySource, new RegExp(`\\b${forbidden}\\b`));
  }
});

test("identity authorization is a narrow security-definer function", () => {
  assert.match(identitySql, /SECURITY DEFINER/);
  assert.match(identitySql, /REVOKE ALL ON FUNCTION public\.varga_authorize_identity/);
  assert.match(identitySql, /GRANT EXECUTE ON FUNCTION public\.varga_authorize_identity\(text, text\) TO varga_reader/);
  assert.match(identitySql, /RETURNS TABLE\(matched boolean, allowed boolean, role text\)/);
  assert.doesNotMatch(identitySql, /RETURNS TABLE\([^)]*(?:email|display_name|telefono|codice_fiscale)/i);
});
