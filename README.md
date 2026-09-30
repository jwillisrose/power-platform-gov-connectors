# Power Platform Connectors for US Government Clouds

A searchable GitHub Pages site for researching Power Apps and Power Automate
connector availability in GCC, GCC High, and DoD environments.

**Published site:** https://jwillisrose.github.io/power-platform-gov-connectors/

The initial dataset and downloadable artifacts were originally generated from
Microsoft Learn research collected in Microsoft Scout. Scout is being retired,
so the research pipeline now runs as a scheduled GitHub Actions workflow
directly in this repository (see [Refresh the research](#refresh-the-research)
below). The site itself is a static, dependency-free application so it can be
hosted directly on GitHub Pages.

## Features

- Search by connector name or publisher
- Filter by government cloud, Power Platform product, tier, and release status
- Link each connector to its Microsoft Learn documentation
- Download the customer PDF and PowerPoint presentation
- Responsive, accessible layout with no build step

## Run locally

The browser blocks `fetch` requests from a `file://` URL, so serve the repository
with any local static web server. For example, if Python is installed:

```powershell
python -m http.server 8000
```

Then open `http://localhost:8000`.

## Publish with GitHub Pages

1. Push this repository to GitHub.
2. Open **Settings > Pages** in the GitHub repository.
3. Under **Build and deployment**, select **GitHub Actions** as the source.
4. Run the **Deploy GitHub Pages** workflow, or push to `main`.

The workflow publishes the repository as a static Pages artifact.

## Refresh the research

Research is refreshed automatically. The
[`Refresh connector research`](.github/workflows/refresh-research.yml)
workflow runs weekly (Mondays, 13:00 UTC) and can also be triggered manually
from the **Actions** tab (`workflow_dispatch`). Each run:

1. Re-scrapes the Microsoft Learn connector reference pages
   (overview, Power Apps, and Power Automate availability tables).
2. Regenerates `data/connectors.json`, connector logos in `assets/logos/`,
   and product icons in `assets/product-icons/`.
3. Compares the new dataset against the previous `data/connectors.json` and,
   if anything changed (connectors added/removed, cloud availability,
   release status, tier, or renames), appends a dated entry to
   `data/change-log.json`.
4. Regenerates the DOCX and PDF handouts and PowerPoint deck in
   `assets/documents/` when the dataset or source dates change, including the
   latest change-log note on the handout.
5. Opens a **pull request** with all changed files — it never commits
   directly to `main`. Review the diff (especially `data/change-log.json`)
   and merge to publish; merging to `main` triggers the Pages deployment
   workflow automatically.

If no changes are found, the workflow run completes without opening a pull
request.

### Run the refresh locally

The refresh requires Python 3.12 and LibreOffice. LibreOffice converts the
generated Word handout to PDF. On Windows, install it with:

```powershell
winget install --id TheDocumentFoundation.LibreOffice --exact
```

Open a new terminal after installation, then run:

```powershell
python -m venv .venv
.\.venv\Scripts\pip.exe install -r scripts/requirements.txt
.\.venv\Scripts\python.exe scripts/build_connector_research.py
```

The script is idempotent: logos are cached by a hash of their Learn URL, so
re-running it without upstream changes produces no diff.

## Operating the GitHub automation

### Publish the current site

1. Open the repository's **Actions** tab.
2. Select **Deploy GitHub Pages**.
3. Select **Run workflow**, leave `main` selected, and confirm.
4. Open the deployment URL shown after the `deploy` job succeeds.

Every push or merged pull request to `main` also publishes automatically.

### Refresh the research now

1. Open the repository's **Actions** tab.
2. Select **Refresh connector research**.
3. Select **Run workflow**, leave `main` selected, and confirm.
4. If Microsoft Learn changed, open the pull request created by the workflow.
5. Review the data and document changes, then merge the pull request. The merge
   automatically publishes the updated site.

A successful run with no pull request means the upstream data has not changed.

### Troubleshooting

- For a failed run, open **Actions**, select the run, and expand the failed step
  to see its error.
- If refresh cannot open a pull request, check **Settings > Actions > General >
  Workflow permissions** and enable **Allow GitHub Actions to create and approve
  pull requests**.
- If Pages does not deploy, check **Settings > Pages** and confirm that
  **GitHub Actions** is the publishing source.
- The workflows require no personal access tokens or custom repository secrets.
  They use the repository-scoped `GITHUB_TOKEN`.

## Data sources and disclaimer

The data is derived from the public Microsoft Learn connector reference,
including the Power Apps and Power Automate government availability tables.
Source update timestamps are stored in `data/connectors.json` and displayed by
the site.

Availability and licensing can change. Confirm current details in official
Microsoft documentation before making architecture or purchasing decisions.

## License

The original code and content in this repository are available under the
[MIT License](LICENSE).
