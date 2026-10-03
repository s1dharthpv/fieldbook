# Privacy

Fieldbook does not collect, store or send data.

- **No data of its own.** The app has no DocTypes, tables or records, and it keeps no copy of any
  workbook it creates.
- **Nothing leaves your site.** The code makes no calls to outside services. A workbook is built on
  your server and sent only to the browser of the System Manager who asked for it.
- **Structure, not records.** To build a workbook the app reads DocType definitions, Custom Fields
  and Property Setters, the database's column types, the installed apps and their versions, and a few
  site settings (default currency, country, language) to make example values believable. It does not
  read your documents, customers or transactions, and example values are synthetic.
- **What a workbook does reveal.** Field names, labels, Select options and help texts are taken from
  your DocType definitions as they are, so they show your customisations and may contain business
  terms. The Schema sheet's version stamp states the date, the Frappe and app versions and the
  database engine and version; it contains no site name, company or user. Treat a workbook like any
  technical document before you share it.
- **Access.** Only users with the System Manager role can open the page or call its methods.

## Publisher

Fieldbook is published by Sidharth PV. Questions about this policy: sidharth.thamban@gmail.com.
