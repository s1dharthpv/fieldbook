# Changelog

## 0.1.0

First release.

- One Desk page, **Fieldbook**, that downloads a three-sheet Excel workbook (fields, example payload,
  JSON Schema) for any DocType: standard, custom, Single, virtual or child table.
- Required shown as Yes, No or Conditional; descriptions include help text, read only, fetched from,
  unique, set-once and default; link, select and condition hints.
- Schema with `enum`, `format` and `maxLength`; believable synthetic examples; a version stamp
  (generation date, Frappe and app versions, database engine) that never includes a site, company or user.
- Options: include standard fields, customisations only, and several DocTypes in one zip with an index.
- A preview of the file before downloading, and clear messages for an unknown DocType, no access,
  or a server or connection failure.
- Works on Frappe 15 and 16, with or without ERPNext, on MariaDB and SQLite (numbers are typed from
  the Frappe field type where the database reports only text).
- CI (tests on Frappe 15 and 16) and linter workflows.
