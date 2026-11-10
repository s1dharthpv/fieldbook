# Contributing to Fieldbook

Thanks for helping. Bug reports, questions and pull requests are all welcome.

## Before you start

- Read the [README](README.md) for what Fieldbook exports and how.
- For a bug or a request, open an issue first so we can agree on the approach.
- To report a security problem, follow [SECURITY.md](SECURITY.md) instead of opening an issue.

## Set up

```bash
bench get-app https://github.com/s1dharthpv/fieldbook --branch develop
bench --site <your-site> install-app fieldbook
cd apps/fieldbook
pre-commit install
```

Fieldbook reads DocType definitions and writes nothing to the site. CI tests against Frappe and ERPNext.

## Make a change

1. Branch from `develop`.
2. Keep the change small and focused, and update the README and CHANGELOG if behaviour changes.
3. Add or update tests, then run the suite:
   ```bash
   bench --site <your-site> set-config allow_tests true
   bench --site <your-site> run-tests --app fieldbook
   ```
4. Run `pre-commit run --all-files`. CI runs the same linters, plus Frappe's Semgrep rules, on every pull request.
5. Open a pull request into `develop`.

## Pull request checklist

- Tests pass and the linters are clean.
- The export stays read-only and System Manager only.
- No real customer or production data in tests, screenshots or issue text.

By contributing you agree that your work is released under the license in [license.txt](license.txt).
