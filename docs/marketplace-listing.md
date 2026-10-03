# Marketplace listing (draft text)

Prepared from the Frappe Cloud Marketplace publishing guidelines. Items marked TODO need a decision
or an asset that is not in this repository yet.

- **Title:** Fieldbook
- **Short description (69 characters, one sentence):** Download an Excel spec sheet for any DocType, including custom fields
- **Long description:**

  Fieldbook turns any DocType into an Excel workbook you can hand to the vendor building your API
  integration, a client or an auditor. It documents your site as it is, custom fields included, with an
  example payload and a JSON Schema. Pick a DocType, check the preview and download. The workbook lists every field with its
  database type, whether it is required, a description and an example, and adds an example payload
  and a JSON Schema as two more sheets. Child tables are included.

  Custom fields are covered, so the file describes your site, not just the standard product. A
  "customisations only" option lists just what was added or changed, and you can export several
  DocTypes at once as a zip with an index. Optional standard fields add `docstatus`, `owner` and the
  like. Each file records the Frappe and app versions and the database it came from.

  Hidden, layout and virtual fields are left out, and so are fields without a database column: they
  hold no data in the database, so there is nothing to put in a payload.

  Fieldbook stores nothing and sends nothing off your site. It reads DocType definitions and your
  default currency, country and language, never your business records. Only System Managers can use
  it. Works on Frappe 15 and 16.

- **Category:** TODO choose the closest available (developer tools or documentation)
- **Support URL:** https://github.com/s1dharthpv/fieldbook/issues (works once the repository is public)
- **Privacy policy URL:** https://github.com/s1dharthpv/fieldbook/blob/main/PRIVACY.md (works once the repository is public)
- **Logo:** `docs/logo/fieldbook.png` (512 x 512, square, centred, no text; SVG source in `docs/logo/fieldbook.svg`)
- **Screenshots:** see `docs/screenshots`
- **Demo video:** `docs/demo/fieldbook-demo.mp4` (20 seconds, 1100 x 860), at https://github.com/s1dharthpv/fieldbook/blob/main/docs/demo/fieldbook-demo.mp4. The Frappe Cloud listing has no video field, so the link goes in the Description.
- **Publisher and contact details:** Sidharth PV, contact sidharth.thamban@gmail.com. Copyright holder in `license.txt` is Sidharth PV.
- **Name uniqueness:** no app called "Fieldbook" was in the public Marketplace list on 2026-10-03 (364 apps). Frappe confirms the name when the app is submitted.

## Publishing steps

Based on Frappe's publishing guide (https://docs.frappe.io/cloud/marketplace/publishing-an-app-to-marketplace).
The deploy branch for the Marketplace is `main`; `develop` is for work in progress and reaches `main`
through a pull request.

1. Make the GitHub repository public (the Marketplace lists open source apps).
2. In the Frappe Cloud dashboard, open the Profile tab and choose "Become a Publisher".
3. In the Marketplace tab choose "+ Add App", then "Add from GitHub" and authorise GitHub if asked.
4. Select `s1dharthpv/fieldbook` and the branch `main`. The guide does not say on which screen the
   branch is chosen, so check this when you get there.
5. After validation, pick the Frappe version the app is compatible with (15 and 16 are tested in CI)
   and click "Add to Marketplace".
6. On the Overview tab fill in the details above: title, short and long description, category,
   logo, screenshots, support URL, privacy URL and the demo video link.
7. Publish a release of the app, then wait for Frappe's review (the guide says up to 10 days).
8. To ship a change later: work on `develop`, open a pull request into `main` (CI and the linters must
   pass), merge, then publish a new release.
